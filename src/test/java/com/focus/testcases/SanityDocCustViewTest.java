package com.focus.testcases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.SanityDocCustViewPage;

import com.focus.base.BaseEngine;

public class SanityDocCustViewTest extends BaseEngine
{

	SanityDocCustViewPage dcp;


	@Test(priority=1032)
	public void checkLoginCompanyWithValidCredentials() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 dcp=new SanityDocCustViewPage(getDriver());
		 Assert.assertEquals(dcp.checkLoginCompanyWithValidCredentials(), true);
	}


	@Test(priority=1035)
	public void checkUserCreatedVoucherDisplayWithDocumentTabInDocumentCustomization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		dcp=new SanityDocCustViewPage(getDriver());
		Assert.assertEquals(dcp.checkUserCreatedVoucherDisplayWithDocumentTabInDocumentCustomization(), true);
	}

	@Test(priority=1036)
	public void checkCreatedVoucherMiscellaneousTabInDocumentCustoimization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		dcp=new SanityDocCustViewPage(getDriver());
		Assert.assertEquals(dcp.checkCreatedVoucherMiscellaneousTabInDocumentCustoimization(), true);
	}

	@Test(priority=1037)
	public void checkCreatedVoucherEditScreenTabInDocumentCustoimization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		dcp=new SanityDocCustViewPage(getDriver());
		Assert.assertEquals(dcp.checkCreatedVoucherEditScreenTabInDocumentCustoimization(), true);
	}


	@Test(priority=1049)
	public void checktheNarrationinHeaderTabEditLayoutinPVVATVIEW() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		dcp=new SanityDocCustViewPage(getDriver());
		Assert.assertEquals(dcp.checktheNarrationinHeaderTabEditLayoutinPVVATVIEW(), true);
	}

	@Test(priority=1056)
	public void checkFieldTypesinBodyGridViewTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkFieldTypesinBodyGridViewTab(), true);
	}

	@Test(priority=1060)
	public void checkDeleteinCreatedViewInViewTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkDeleteinCreatedViewInViewTab(), true);
	}

	@Test(priority=1061)
	public void checkEditinginBodyFieldAndFooterFieldinInViewTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkEditinginBodyFieldAndFooterFieldinInViewTab(), true);
	}

	@Test(priority=1063)
	public void checkRuleTabInDocumentCustomization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkRuleTabInDocumentCustomization(), true);
	}

	@Test(priority=1065)
	public void checkCreatingRuleinRuleTAbUnderDocumentCustomization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkCreatingRuleinRuleTAbUnderDocumentCustomization(), true);
	}

	@Test(priority=1066)
	public void checkNavigateToDocumentNumberingTabInPurchaseVoucherVatView() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkNavigateToDocumentNumberingTabInPurchaseVoucherVatView(), true);
	}

	@Test(priority=1077)
	public void checkSavingOnClickingUpdateOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkSavingOnClickingUpdateOption(), true);
	}

	@Test(priority=1078)
	public void checkSavedFieldsinPvVatViewVoucherDocumentNUmberingTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkSavedFieldsinPvVatViewVoucherDocumentNUmberingTab(), true);
	}




	@Test(priority=1079)
	public void checkRaiseDocumentOptionInTriggers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkRaiseDocumentOptionInTriggers(), true);
	}


	@Test(priority=1080)
	public void checkSaveOptionInTriggerTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkSaveOptionInTriggerTab(), true);
	}

	@Test(priority=1081)
	public void checkDeleteTriggerOptionInTriggers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkDeleteTriggerOptionInTriggers(), true);
	}


	@Test(priority=1082)
	public void checkSaveTriggerWithRaiseDocumentOptionInTriggers() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkSaveTriggerWithRaiseDocumentOptionInTriggers(), true);
	}




	@Test(priority=1083)
	public void checkEditingInTriggersTabUnderPVVATView() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkEditingInTriggersTabUnderPVVATView(), true);
	}


	@Test(priority=1084)
	public void checkLogoutUserDocumentCustOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkLogoutUserDocumentCustOptions(), true);
	}


	@Test(priority=1085)
	public void checkLoginAfterDocumentCustOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		dcp=new SanityDocCustViewPage(getDriver());	
		Assert.assertEquals(dcp.checkLoginCompanyWithValidCredentials(), true);
	}


	
	 //////////////////////////Creating View for adding new fields in edit screen////////
	 
	 
	 
	 
	 
	 @Test(priority=1186)
		public void checkCreatingNewFieldsinEditScreenTabatDifferentPositions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
		 dcp=new SanityDocCustViewPage(getDriver());	
			  Assert.assertEquals(dcp.checkCreatingNewFieldsinEditScreenTabatDifferentPositions(), true);
		}
	
	 
	 
	 @Test(priority=1187)
		public void  checkCreateViewforNewlyAddedFields() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
		 dcp=new SanityDocCustViewPage(getDriver());	
			  Assert.assertEquals(dcp. checkCreateViewforNewlyAddedFields(), true);
		}
	 
	  
	 
	 @Test(priority=1188)
		public void  checkFieldsinViewTabafterRefreshView() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
		 dcp=new SanityDocCustViewPage(getDriver());	
			  Assert.assertEquals(dcp. checkFieldsinViewTabafterRefreshView(), true);
		}
	
	 
	 @Test(priority=1189)
		public void   checkCreatedViewFieldsatVoucherLevel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
		 dcp=new SanityDocCustViewPage(getDriver());	
			  Assert.assertEquals(dcp.  checkCreatedViewFieldsatVoucherLevel(), true);
		}
	 
	 @Test(priority=1190)
		public void   checkDeleteAddedFieldsinEditScreenTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
		 dcp=new SanityDocCustViewPage(getDriver());	
			  Assert.assertEquals(dcp.  checkDeleteAddedFieldsinEditScreenTab(), true);
		}
	 
	 
	
	

}
