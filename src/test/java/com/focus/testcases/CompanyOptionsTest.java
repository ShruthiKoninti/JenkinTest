package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.CompanyOptionsPage;
import com.focus.Pages.MastersPage;
import com.focus.base.BaseEngine;

public class CompanyOptionsTest extends BaseEngine
{
	CompanyOptionsPage COP;
	
	
	
	
	
	 @Test(priority=100)
	 public void checkDeleteCompanyOptionInSearch() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 COP=new CompanyOptionsPage(getDriver());
		 Assert.assertEquals(COP.checkDeleteCompanyOptionInSearch(), true);
	 }
	
	
	
	 @Test(priority=101)
	 public void checkRestoreCompanyForSanity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 COP=new CompanyOptionsPage(getDriver());
		 Assert.assertEquals(COP.checkRestoreCompanyForSanity(), true);
	 }
	 
	 @Test(priority=102)
	 public void checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 COP=new CompanyOptionsPage(getDriver());
		 Assert.assertEquals(COP.checkLogin(), true);
	 }
	 
	 
	 
	 @Test(priority=108)
		public void verifyAvailabilityOptionsInWrapperLayout() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
	    	  COP=new CompanyOptionsPage(getDriver());
			  Assert.assertEquals(CompanyOptionsPage.checkAvailibityOfOptionsInWrapperLayout(), true);
		}
	    
		
		@Test(priority=109)
		public void verifyAvailabilityOptionsInAboutIdQuickAccessMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
	    	  COP=new CompanyOptionsPage(getDriver());
			  Assert.assertEquals(CompanyOptionsPage.checkAboutIdQuickAccessMenu(), true);
		   
		}
	 
	 
	 @Test(priority=110)
		public void verifyClickOnAbout() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
	    	  COP=new CompanyOptionsPage(getDriver());
			  Assert.assertEquals(CompanyOptionsPage.clickOnAbout(), true);
		   
		}

		@Test(priority=111)
		public void verifySearchOptionByInputingEditCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
	    	  COP=new CompanyOptionsPage(getDriver());
			  Assert.assertEquals(CompanyOptionsPage.checkEditCompanyOptionInSearch(), true);
		   
		}
		
		
		@Test(priority=112)
		public void verifySelectEditCompanyFromSearch() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
	    	  COP=new CompanyOptionsPage(getDriver());
			  Assert.assertEquals(CompanyOptionsPage.checkSelectingEditCompanyInSearch(), true);
		   
		}
		
		@Test(priority=113)
		public void verifyAlertByUpdatingCompanyClickOnOkButton() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
	    	  COP=new CompanyOptionsPage(getDriver());
			  Assert.assertEquals(CompanyOptionsPage.checkUpdateTheCompanyAlertThroughSearch(), true);
		   
		}

		
		@Test(priority=114)
		public void verifyCloseEditCompanyPageClickOnCancelButton() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
	    	  COP=new CompanyOptionsPage(getDriver());
			  Assert.assertEquals(CompanyOptionsPage.clickOnCloseTheEditCompanyThroughSearch(),true);
		}
		
		
		@Test(priority=115)
		 public void checkLogOutAndLoginAfterEditCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 COP=new CompanyOptionsPage(getDriver());
			 Assert.assertEquals(COP.checkLogOutAndLogin(), true);
		 }
		
		

}
