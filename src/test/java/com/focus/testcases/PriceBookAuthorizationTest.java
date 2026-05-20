package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.AccountsPage;
import com.focus.Pages.PricebookAuthorizationPage;
import com.focus.base.BaseEngine;

public class PriceBookAuthorizationTest extends BaseEngine{
	
	PricebookAuthorizationPage pap;
	
	
	
	 @Test(priority=9999)
	public void CheckLogin() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.CheckLogin(), true);	
	}
	
	 
	 
	
	 @Test(priority=10000)
		public void checkOpenPricebookAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkOpenPricebookAuthorization(), true);	
		}
		
		
		
		@Test(priority=10001)
		public void checkPricebookDropdownOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());	
			Assert.assertEquals(pap.checkPricebookDropdownOptions(), true);
		}
		
		
		
		@Test(priority=10002)
		public void checkSaveWithoutSelectionOfPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSaveWithoutSelectionOfPriceBook(), true);
		}
		

		@Test(priority=10003)
		public void checkSelectingBuyingPriceBookinPriceBookDropdown() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSelectingBuyingPriceBookinPriceBookDropdown(), true);
		}
		
		@Test(priority=10004)
		public void checkSaveWithoutName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSaveWithoutName(), true);
		}
		
		@Test(priority=10005)
		public void checkActiveCheckBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkActiveCheckBox(), true);
		}
		
		
		@Test(priority=10006)
		public void checkInputNameAndClickOnSaveWithoutAuthorizationSettings() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkInputNameAndClickOnSaveWithoutAuthorizationSettings(), true);
		}
		
		
		@Test(priority=10007)
		public void checkStartPoint() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStartPoint(), true);
		}
		
		@Test(priority=10008)
		public void checkConditionTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkConditionTabOptions(), true);
		}
		
		@Test(priority=10009)
		public void checkUserSelectionTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkUserSelectionTabOptions(), true);
		}
		
		
		@Test(priority=100010)
		public void checkUserDropdownOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkUserDropdownOptions(), true);
		}
		
		
		@Test(priority=100011)
		public void checkAlertsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAlertsTabOptions(), true);
		}
		
		
		
		@Test(priority=100012)
		public void checkCancelButtonInDefinitionScreen() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCancelButtonInDefinitionScreen(), true);
		}
		
		@Test(priority=100013)
		public void checkDescriptionAsMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDescriptionAsMandatory(), true);
		}
		
		
		
		@Test(priority=100014)
		public void checkConditionAsMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkConditionAsMandatory(), true);
		}
		
		
		@Test(priority=100015)
		public void checkUserSelectionIsManadatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkUserSelectionIsManadatory(), true);
		}
		
		
		@Test(priority=100016)
		public void checkOkButtonOnSelectingUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkOkButtonOnSelectingUser(), true);
		}
		
		
		@Test(priority=100017)
		public void checkSavingSellingPriceBookAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingSellingPriceBookAuthorization(), true);
		}
		
		
		
		@Test(priority=100018)
		public void checkCreateaSellerPriceBookafterAuthorizationSaved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCreateaSellerPriceBookafterAuthorizationSaved(), true);
		}
		
		
		@Test(priority=100019)
		public void checkStatusofSellerPricebookafterAuthorizationSaved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofSellerPricebookafterAuthorizationSaved(), true);
		}
		
		
		
		
		@Test(priority=100020)//Issue with pronghorn
		public void checkRateValueloadingbeforeAuthorizationPricebookinSalesInvoiceVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRateValueloadingbeforeAuthorizationPricebookinSalesInvoiceVoucher(), true);
		}
		
		
		
		@Test(priority=100021)
		public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		@Test(priority=100022)
		public void checkAuthorizationofSellerPricebookwithUserAllOptionsST() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationofSellerPricebookwithUserAllOptionsST(), true);
		}
		
		
		
		@Test(priority=100023)
		public void checkLogoutandloginSU() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		
		@Test(priority=100024)//Issue with pronghorn
		public void checkRateLoadingafterAuthorizeSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRateLoadingafterAuthorizeSellerPricebook(), true);
		}
		
		
		
		
		
		//@Test(priority=100018)
		public void checkCancelButton() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCancelButton(), true);
		}
		
		
		@Test(priority=100025)
		public void checkDisplayingMastersNameInListOfMasters() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
		}
		
		
		@Test(priority=100026)
		public void checkSavedNameDisplayInNameTextBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
		}

		
		@Test(priority=100027)
		public void checkEditPriceBookAuthorizationSettings() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
		}
		
		
		
		
		@Test(priority=100028)
		public void checkDateRangeCheckBoxinAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDateRangeCheckBoxinAuthorization(), true);
		}
		
		
		
		
		
		@Test(priority=100029)
		public void checkEditingthePricebookAfterClickingDateRangeinAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditingthePricebookAfterClickingDateRangeinAuthorization(), true);
		}
		
	
		
		@Test(priority=100030)
		public void checkStatusofeditPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofeditPricebook(), true);
		}
		
		
		
		
		@Test(priority=100031)
		public void checkRateLoadingforPendingandAuthorizedPricebookItemsinSalesInvoiceVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRateLoadingforPendingandAuthorizedPricebookItemsinSalesInvoiceVoucher(), true);
		}
		
		
		@Test(priority=100032)
		public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeEditedSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		
		@Test(priority=100033)
		public void checkAuthorizationforEditedRowsinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinSellerPricebook(), true);
		}
		
		

		@Test(priority=100034)
		public void checkLogoutandloginSUAgain() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		
		
		@Test(priority=100035) 
		public void checkRateLoadingforPendingandAuthorizedPricebookItemsinSalesInvoiceVoucherafterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRateLoadingforPendingandAuthorizedPricebookItemsinSalesInvoiceVoucherafterAuthorization(), true);
		}
		
		@Test(priority=100036)
		public void checkDisplayingMastersNameInListOfMastersAgain() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
		}
		
		
		@Test(priority=100037)
		public void checkSavedNameDisplayInNameTextBoAgainx() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
		}

		
		@Test(priority=100038) 
		public void checkConditionCheckbox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkConditionCheckbox(), true);
		}
		
		
		
		
		@Test(priority=100039) 
		public void checkSavingPricebookafterPricebookAuthorizewithCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingPricebookafterPricebookAuthorizewithCondition(), true);
		}
		
		@Test(priority=100040) 
		public void checkStatusofPricebookforAuthorizationWithRateCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofPricebookforAuthorizationWithRateCondition(), true);
		}
		
		@Test(priority=100041)
		public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeConditionSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		@Test(priority=100042)
		public void checkConditionedAuthorizationinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinSellerPricebook(), true);
		}
		
		

		@Test(priority=100043)
		public void checkLogoutandloginSUAgainafterAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		@Test(priority=100044) 
		public void checkStatusofPricebookAfterAuthorizationWithRateCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofPricebookAfterAuthorizationWithRateCondition(), true);
		}
		
		
		
	
		@Test(priority=100045) 
		public void checkDeletetheConditionPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheConditionPriceBook(), true);
		}
	
	
	
	//Condition as Product Name
	
	@Test(priority=100046)
	public void checkDisplayingMastersNameInListOfMastersAgainforProductName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
	}
	
	@Test(priority=100047)
	public void checkSavedNameDisplayInNameTextBoxAgain() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
	}
	
	@Test(priority=100048)
	public void checkEditingPricebookAuthorizationforProductNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkEditingPricebookAuthorizationforProductNameasCondition(), true);
	}
	
	@Test(priority=100049)
	public void checkSavingPricebookforProductnameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavingPricebookforProductnameasCondition(), true);
	}
	
	@Test(priority=100050)
	public void checkStatusofthePricebookforProductnameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkStatusofthePricebookforProductnameasCondition(), true);
	}
	
	
	@Test(priority=100051)
	public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeProductNameConditionSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
	}
	
	
	
	@Test(priority=100052)
	public void checkConditionedAuthorizationofProductNameinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkAuthorizationforEditedRowsinSellerPricebook(), true);
	}
	
	

	@Test(priority=100053)
	public void checkLogoutandloginSUAgainafterAuthforProductName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutandloginSU(), true);
	}
	
	
	
	
	
	@Test(priority=100054)
	public void checkStatusofthePricebookAfterAuthforProductnameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkStatusofthePricebookAfterAuthforProductnameasCondition(), true);
	}
	
	
	
	
	
	
	
	
	@Test(priority=100055)
	public void checkDisplayingMastersNameInListOfMastersafterEditing() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
	}
	
	@Test(priority=100056)
	public void checkSavedNameDisplayInNameTextBoxafterEditing() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
	}
	
	
	@Test(priority=100057)
	public void checkUndoAuthorizationEditinginPricebookAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkUndoAuthorizationEditinginPricebookAuthorization(), true);
	}
	
	
	@Test(priority=100058)
	public void checkEditingaPricebookforUndoAuthEditingOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkEditingaPricebookforUndoAuthEditingOption(), true);
	}
	
	
	@Test(priority=100059)
	public void checkLogoutAndLoginWithUserAllOptionsSTsafterEditAuthorizeSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
	}
	
	
	
	@Test(priority=100060)
	public void checkConditionedAuthorizationforUndoAuthEditing() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkAuthorizationforEditedRowsinSellerPricebook(), true);
	}
	
	

	@Test(priority=100061)
	public void checkLogoutandloginSUAgainafterEditingAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutandloginSU(), true);
	}
	
	
	
	@Test(priority=100062) 
	public void checkDeletetheConditionProductNamePriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDeletetheConditionProductNamePriceBook(), true);
	}
	
	
	
//Account Name as Condition	
	
	@Test(priority=100063)
	public void checkDisplayingMastersNameInListOfMastersafterEditingAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
	}
	
	@Test(priority=100064)
	public void checkSavedNameDisplayInNameTextBoxafterEditingAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
	}
	
	@Test(priority=100065)
	public void checkEditingPricebookAuthorizationforAccountNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkEditingPricebookAuthorizationforAccountNameasCondition(), true);
	}
	
	@Test(priority=100066)
	public void checkSavingPricebookforAccountNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavingPricebookforAccountNameasCondition(), true);
	}
	
	@Test(priority=100067)
	public void checkStatusofthePricebookforAccountNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkStatusofthePricebookforAccountnameasCondition(), true);
	}
	
	
	@Test(priority=100068)
	public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeAccountNameConditionSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
	}
	
	
	
	@Test(priority=100069)
	public void checkConditionedAuthorizationofAccountNameinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkAuthorizationforEditedRowsinSellerPricebook(), true);
	}
	
	

	@Test(priority=100070)
	public void checkLogoutandloginSUAgainafterAuthforAccountName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutandloginSU(), true);
	}
	
	
	
	@Test(priority=100071)
	public void checkStatusofthePricebookAfterAuthforAccountnameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkStatusofthePricebookAfterAuthforAccountnameasCondition(), true);
	}
	
	@Test(priority=100072)
	public void checkDeletetheConditionAccountNamePriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDeletetheConditionAccountNamePriceBook(), true);
	}
	
	
	//Condition as Group Account
	
		
	@Test(priority=100073)
	public void checkDisplayingMastersNameInListOfMastersafterEditingAuthforAcc() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
	}
	
	@Test(priority=100074)
	public void checkSavedNameDisplayInNameTextBoxafterEditingAuthforAcc() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
	}
	
	@Test(priority=100075)
	public void checkEditingPricebookAuthorizationforAccountNameasGroupAccCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkEditingPricebookAuthorizationforAccountNameasGroupAccCondition(), true);
	}
	
	@Test(priority=100076)
	public void checkSavingPricebookforGroupAccNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavingPricebookforGroupAccNameasCondition(), true);
	}
	
	@Test(priority=100077)
	public void checkStatusofthePricebookforGroupAccNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkStatusofthePricebookforGroupAccNameasCondition(), true);
	}
	
	
	@Test(priority=100078)
	public void checkRateValueLoadingBeforeAuthofPricebookforAccountGroup() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkRateValueLoadingBeforeAuthofPricebookforAccountGroup(), true);
	}
	
	@Test(priority=100079)
	public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeAccountGrpNameConditionSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
	}
	
	
	
	@Test(priority=100080)
	public void checkConditionedAuthorizationofAccountGrpNameinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkAuthorizationforEditedRowsinSellerPricebook(), true);
	}
	
	

	@Test(priority=100081)
	public void checkLogoutandloginSUAgainafterAuthforAccountGrpName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutandloginSU(), true);
	}
	
	
	@Test(priority=100082)
	public void checkStatusofthePricebookAfterAuthforGroupAccNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkStatusofthePricebookAfterAuthforGroupAccNameasCondition(), true);
	}
	
	
	@Test(priority=100083)
	public void checkDeletetheConditionGroupAccountNamePriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDeletetheConditionGroupAccountNamePriceBook(), true);
	}
	
	
	
	
//Condition as Min 	Qty
	
	@Test(priority=100084)
	public void checkDisplayingMastersNameInListOfMastersafterEditingAuthforQty() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
	}
	
	@Test(priority=100085)
	public void checkSavedNameDisplayInNameTextBoxafterEditingAuthforQty() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
	}
	
	@Test(priority=100086)
	public void checkEditPricebookAuthorizationforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkEditPricebookAuthorizationforMinimumQuantityasCondition(), true);
	}
	
	@Test(priority=100087)
	public void checkSavingPricebookforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavingPricebookforMinimumQuantityasCondition(), true);
	}
	
	@Test(priority=100088)
	public void checkStatusofthePricebookforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkStatusofthePricebookforMinimumQuantityasCondition(), true);
	}
	
	
	@Test(priority=100089)
	public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeMinQtyConditionSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
	}
	
	
	
	@Test(priority=100090)
	public void checkConditionedAuthorizationofMinQtyinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkAuthorizationforEditedRowsinSellerPricebook(), true);
	}
	
	

	@Test(priority=100091)
	public void checkLogoutandloginSUAgainafterAuthforMinQty() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutandloginSU(), true);
	}
	
	
	@Test(priority=100092)
	public void checkStatusofthePricebookAfterAuthforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkStatusofthePricebookAfterAuthforMinimumQuantityasCondition(), true);
	}
	
	@Test(priority=100093)
	public void checkDeletetheConditionMinQtyPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDeletetheConditionMinQtyPriceBook(), true);
	}
	
	
	
	//Condition as Tag
	
	@Test(priority=100094)
	public void checkSettingsofSellerPriceBookforTag() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSettingsofSellerPriceBookforTag(), true);
	}
	
	
	
	
	@Test(priority=100095)
	public void checkDisplayingMastersNameInListOfMastersafterEditingAuthforTag() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
	}
	
	@Test(priority=100096)
	public void checkSavedNameDisplayInNameTextBoxafterEditingAuthforTag() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
	}
	
	
	
	
	
	@Test(priority=100097)
	public void checkEditingPricebookAuthorizationforTagasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkEditingPricebookAuthorizationforTagasCondition(), true);
	}
	
	@Test(priority=100098)
	public void checkSavingPricebookforTagasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavingPricebookforTagasCondition(), true);
	}
	
	@Test(priority=100099)
	public void checkStatusofthePricebookforTagasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkStatusofthePricebookforTagasCondition(), true);
	}
	
	
	@Test(priority=100100)
	public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeTagConditionSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
	}
	
	
	
	@Test(priority=100101)
	public void checkConditionedAuthorizationofTaginSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkAuthorizationforEditedRowsinSellerPricebook(), true);
	}
	
	

	@Test(priority=100102)
	public void checkLogoutandloginSUAgainafterAuthforTag() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkLogoutandloginSU(), true);
	}
	
	
	
	
	@Test(priority=100103)
	public void checkStatusofthePricebookAfterAuthforTagasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkStatusofthePricebookAfterAuthforTagasCondition(), true);
	}
	
	
	
	
	@Test(priority=100104)
	public void checkDeletetheConditionTagPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDeletetheConditionTagPriceBook(), true);
	}
	
	
	//Alerts
	
	@Test(priority=100105)
	public void checkDisplayingMastersNameInListOfMastersafterEditingAuthforAlerts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
	}
	
	@Test(priority=100106)
	public void checkSavedNameDisplayInNameTextBoxafterEditingAuthforAlerts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
	}
	
	
	
	
	@Test(priority=100107)
	public void checkEditPriceBookAuthorizationSettingsforAlerts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		pap=new PricebookAuthorizationPage(getDriver());		
		Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
	}
	
	
		
		@Test(priority=100108)
		public void checkEditOnSelectAlertsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditOnSelectAlertsTabOptions(), true);
		}
		
		
		@Test(priority=100109)
		public void checkSaveWithAllMandatoryInputs() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSaveWithAllMandatoryInputs(), true);
		}
		
		
		
		@Test(priority=100110)
		public void checkCreatedEmailInEmailHedaerTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCreatedEmailInEmailHedaerTemplate(), true);
		}
		
		
				
		@Test(priority=100111)
		public void checkEmailTemplateExpandInFieldsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEmailTemplateExpandInFieldsTabOptions(), true);
		}
		
		
		
		@Test(priority=100112)
		public void checkFieldExpansionDetails() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkFieldExpansionDetails(), true);
		}
		
		
		
		@Test(priority=100113)
		public void checkCreationofEmailTemplateinPricebookAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCreationofEmailTemplateinPricebookAuthorization(), true);
		}
		
		@Test(priority=100114)
		public void checkCreatedEmailTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCreatedEmailTemplate(), true);
		}
		
		@Test(priority=100115)
		public void checkSaveAlertsTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSaveAlertsTemplate(), true);
		}
		
		@Test(priority=100116)
		public void checkCreatedAlertTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCreatedAlertTemplate(), true);
		}
		
		
		
		@Test(priority=100116)
		public void checkSavingAPricebookafterAlertsCreated() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingAPricebookafterAlertsCreated(), true);
		}
		
		
		
		
		@Test(priority=100117)
		public void checkStatusofPricebookforAlerts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofPricebookforAlerts(), true);
		}
		
		
		@Test(priority=100118)
		public void checkLogoutAndLoginWithUserAllOptionsSTforAlerts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		@Test(priority=100119)
		public void checkConditionedAuthorizationforAlertsk() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinSellerPricebook(), true);
		}
		
		

		@Test(priority=100120)
		public void checkLogoutandloginSUAgainafterAlertsReceived() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		@Test(priority=100121)
		public void checkStatusofPricebookforAlertsAfterAuthorize() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofPricebookforAlertsAfterAuthorize(), true);
		}
		
		@Test(priority=100122)
		public void checkDeletetheAlertPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheAlertPriceBook(), true);
		}
		
		
		@Test(priority=100123)
		public void checkDisplayingMastersNameInListOfMastersafterEditingAuthforAlertsApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
		}
		
		@Test(priority=100124)
		public void checkSavedNameDisplayInNameTextBoxafterEditingAuthforAlertApproveds() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
		}
		
		
		
		
		@Test(priority=100125)
		public void checkEditPriceBookAuthorizationSettingsforAlertsApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
		}
		
		
		
		
		
		@Test(priority=100126)
		public void checkSelectCreatedAlertTemplateforApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSelectCreatedAlertTemplateforApproved(), true);
		}
		
		
		
		
		@Test(priority=100127)
		public void checkRecipientforAlertsApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRecipientforAlertsApproved(), true);
		}
		
		
		@Test(priority=100128)
		public void checkAgainSavingAPricebookafterAlertsCreatedforApprove() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingAPricebookafterAlertsCreated(), true);
		}
		
		
		
		
		@Test(priority=100129)
		public void checkStatusofPricebookforAlertsagain() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofPricebookforAlerts(), true);
		}
		
		
		@Test(priority=100130)
		public void checkLogoutAndLoginWithUserAllOptionsSTforAlertsAgain() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		@Test(priority=100131)
		public void checkConditionedAuthorizationforAlertsApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinSellerPricebook(), true);
		}
		
		

		@Test(priority=100132)
		public void checkLogoutandloginSUAgainafterAlertsApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		
		
		@Test(priority=100133)
		public void checkAlertinOwnerAfterApprovedofAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAlertinOwnerAfterApprovedofAuthorization(), true);
		}
		
		
		@Test(priority=100134)
		public void checkDeletetheAlertPriceBookforApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheAlertPriceBook(), true);
		}
		
		
		
		//Rejected
		
		
		@Test(priority=100135)
		public void checkDisplayingMastersNameInListOfMastersafterEditingAuthforAlertsRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
		}
		
		@Test(priority=100136)
		public void checkSavedNameDisplayInNameTextBoxafterEditingAuthforAlertRejecteds() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
		}
		
		
		
		
		@Test(priority=100137)
		public void checkEditPriceBookAuthorizationSettingsforAlertsRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
		}
		
		
		
		
		
		@Test(priority=100138)
		public void checkSelectCreatedAlertTemplateforRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSelectCreatedAlertTemplateforRejected(), true);
		}
		
		
		
		
		@Test(priority=100139)
		public void checkRecipientforAlertsRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRecipientforAlertsRejected(), true);
		}
		
		
		@Test(priority=100140)
		public void checkAgainSavingAPricebookafterAlertsCreatedforRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingAPricebookafterAlertsCreated(), true);
		}
		
		
		
		
		@Test(priority=100141)
		public void checkStatusofPricebookforAlertsagainafterReject() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofPricebookforAlerts(), true);
		}
		
		
		@Test(priority=100142)
		public void checkLogoutAndLoginWithUserAllOptionsSTforAlertsAgainAfterReject() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		@Test(priority=100143)
		public void checkConditionedAuthorizationforAlertsRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRejectionforEditedRowsinSellerPricebook(), true);
		}
		
		

		@Test(priority=100144)
		public void checkLogoutandloginSUAgainafterAlertsRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		@Test(priority=100145)
		public void checkStatusofthePricebbokafterRejection() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofthePricebbokafterRejection(), true);
		}
		
		
		
		@Test(priority=100146)
		public void checkAlertinOwnerAfterRejectedofAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAlertinOwnerAfterRejectedofAuthorization(), true);
		}
		
		
		@Test(priority=100147)
		public void checkDeletetheAlertPriceBookAfterRejection() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheAlertPriceBook(), true);
		}
		
		
		
		////Stopped
		
		
		
		@Test(priority=100148)
		public void checkDisplayingMastersNameInListOfMastersafterEditingAuthforAlertsStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
		}
		
		@Test(priority=100149)
		public void checkSavedNameDisplayInNameTextBoxafterEditingAuthforAlertStoppeds() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
		}
		
		
		
		
		@Test(priority=100150)
		public void checkEditPriceBookAuthorizationSettingsforAlertsStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
		}
		
		
		
		
		
		@Test(priority=100151)
		public void checkSelectCreatedAlertTemplateforStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSelectCreatedAlertTemplateforStopped(), true);
		}
		
		
		
		
		@Test(priority=100152)
		public void checkRecipientforAlertsStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRecipientforAlertsStopped(), true);
		}
		
		
		@Test(priority=100153)
		public void checkAgainSavingAPricebookafterAlertsCreatedforStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingAPricebookafterAlertsCreated(), true);
		}
		
		
		
		
		@Test(priority=100154)
		public void checkStatusofPricebookforAlertsagainafterStoped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofPricebookforAlerts(), true);
		}
		
		
		@Test(priority=100155)
		public void checkLogoutAndLoginWithUserAllOptionsSTforAlertsAgainAfterStoped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		@Test(priority=100156)
		public void checkStopedforEditedRowsinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStopedforEditedRowsinSellerPricebook(), true);
		}
		
		

		@Test(priority=100157)
		public void checkLogoutandloginSUAgainafterAlertsStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		@Test(priority=100158)
		public void checkStatusofthePricebbokafterStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofthePricebbokafterStopped(), true);
		}
		
		
		
		@Test(priority=100160)
		public void checkAlertinOwnerAfterStoppedofAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAlertinOwnerAfterStoppedofAuthorization(), true);
		}
		
		
		@Test(priority=100161)
		public void checkDeletetheAlertPriceBookAfterStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheAlertPriceBook(), true);
		}
		
		
		
		
	/////Multiple Users with multiple rows
		
			@Test(priority=100162)
			public void checkDisplayingMastersNameInListOfMastersafterEditingAuthforMultipleUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
			}
			
			@Test(priority=100163)
			public void checkSavedNameDisplayInNameTextBoxafterEditingAuthforforMultipleUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
			}
			
			
			
			
			@Test(priority=100164)
			public void checkEditPriceBookAuthorizationSettingsforMultipleUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
			}
			
			
			
			@Test(priority=100165)
			public void checkSavingPricebookAuthorizationforMultipleUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkSavingPricebookAuthorizationforMultipleUsers(), true);
			}
			
			
			
			@Test(priority=100166)
			public void checkSavingPricebookforMultipleUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkSavingPricebookforMultipleUsers(), true);
			}
			
			
			
			@Test(priority=100167)
			public void checkStatusofPricebookforMultipleUsersBeforeAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkStatusofPricebookforMultipleUsersBeforeAuthorization(), true);
			}
			
			
			@Test(priority=100168)
			public void checkLogoutAndLoginWithUserAllOptionsSTforMultipleUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
			}
			
			@Test(priority=100169)
			public void checkAuthorizeMultipleUsersPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAuthorizeMultipleUsersPricebook(), true);
			}
			
			

			@Test(priority=100170)
			public void checkLogoutandloginSUAgainafterAuthorizationofMultipleUsersPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandloginSU(), true);
			}
			
			
			
			
			@Test(priority=100171)
			public void checkEditingthePricebookAfterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkEditingthePricebookAfterAuthorization(), true);
			}
			
			
			
			@Test(priority=100172)
			public void checkLogoutAndLoginWithUserAllOptionsSTforMultipleUsersAfterEdit() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
			}
			
			
			
			
			@Test(priority=100173)
			public void checkDeleteRowinPricebookatUserLevel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkDeleteRowinPricebookatUserLevel(), true);
			}
			
			
			
			@Test(priority=100174)
			public void checkLogoutandloginSUAgainafterDeleteRowinPricebookUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandloginSU(), true);
			}
				
			
			
			//@Test(priority=100175)
			public void checkDeletePricebookforMultipleUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkDeletePricebookforMultipleUsers(), true);
			}
			
		 
		 
		 
			//Checking Alert in Users
				

			@Test(priority=100176)
			public void checkEditingthePricebookAfterAuthorizationforAlert() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkEditingthePricebookAfterAuthorization(), true);
			}
			
			
			
			@Test(priority=100177)
			public void checkLogoutAndLoginWithUserAllOptionsSTforMultipleUsersAfterEditforAlert() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
			}
			
			
			
			@Test(priority=100178)
			public void checkAlertDisplayedinUserAllOptionsSTUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAlertDisplayedinUserAllOptionsSTUser(), true);
			}
			
			
			
			@Test(priority=100179)
			public void checkLogoutandLoginWithTestwithMultipleUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandLoginWithTestwithMultipleUsers(), true);
			}
			@Test(priority=100180)
			public void checkAuthorizeMultipleUsersPricebookfroAlert() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAuthorizeMultipleUsersPricebook(), true);
			}
			
			
			
			@Test(priority=100181)
			public void checkAlertDisplayedAfterAuthorizationforUsers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAlertDisplayedAfterAuthorization(), true);
			}
			
			@Test(priority=100182)
			public void checkLogoutAndLoginWithUserAllOptionsSTforMultipleUsersforAlertAfterAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
			}
			
			@Test(priority=100183)
			public void checkAlertDisplayedAfterAuthorizationatUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAlertDisplayedAfterAuthorization(), true);
			}
			
			@Test(priority=100184)
			public void checkLogoutandloginSUAgainafterAuthorizationofMultipleUsersPricebookforAlertAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandloginSU(), true);
			}
			
			//Adding Duplicate Row and check for Alerts
				
			
			@Test(priority=100185)
			public void checkAddingDuplicateRowinPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAddingDuplicateRowinPricebook(), true);
			}
			
			
			@Test(priority=100186)
			public void checkLogoutAndLoginWithUserAllOptionsSTforMultipleUsersAfterDuplicate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
			}
			
			
			
			
			@Test(priority=100187)
			public void checkAuthorizingtheRowfromHomePageinPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAuthorizingtheRowfromHomePageinPricebook(), true);
			}
			
					
			
			@Test(priority=100188)
			public void checkAlertDisplayedAfterAuthorizationforDuplicate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAlertDisplayedAfterAuthorization(), true);
			}
			
			
			
			@Test(priority=100189)
			public void checkLogoutandloginSUAgainafterAuthorizationofMultipleUsersPricebookforAlertAuthforduplicate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandloginSU(), true);
			}
			
			
			
			//Insert Row 
			
			@Test(priority=100190)
			public void checkInsertingRowinPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkInsertingRowinPricebook(), true);
			}
			
			@Test(priority=100191)
			public void checkLogoutAndLoginWithUserAllOptionsSTforMultipleUsersAfterInsert() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
			}
			
			
			@Test(priority=100192)
			public void checkAuthorizingRowafterInsertRowinPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAuthorizingRowafterInsertRowinPricebook(), true);
			}
			
			@Test(priority=100193)
			public void checkAlertDisplayedAfterAuthorizationforUsersforInsert() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAlertDisplayedAfterAuthorization(), true);
			}
			
			
			
			@Test(priority=100194)
			public void checkLogoutandloginSUAgainafterAuthorizationofMultipleUsersPricebookforAlertAuthforInsert() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandloginSU(), true);
			}
			
			
			
		  //Manually adding Row
		 
		 
		 @Test(priority=100195)
			public void checkAddingRowManuallyinPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAddingRowManuallyinPricebook(), true);
			}
		 
			@Test(priority=100196)
			public void checkLogoutAndLoginWithUserAllOptionsSTforMultipleUsersAfterManually() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
			}
			
		 
		 
		 @Test(priority=100197)
			public void checkAuthorizingRowAfterAddingRowManually() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAuthorizingRowAfterAddingRowManually(), true);
			}
		 
		 @Test(priority=100198)
			public void checkAlertDisplayedAfterAuthorizationforUsersforManually() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAlertDisplayedAfterAuthorization(), true);
			}
			
			
			
			@Test(priority=100199)
			public void checkLogoutandloginSUAgainafterAuthorizationofMultipleUsersPricebookforAlertAuthforManually() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandloginSU(), true);
			}
			
			
			
			//Delete Pending Records for Alert Disappear and saving msg for Pricebook
			
			
			@Test(priority=100200)
			public void checkSavingPricebookforDeletePendingRow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkSavingPricebookforDeletePendingRow(), true);
			}
			
			@Test(priority=100201)
			public void checkLogoutAndLoginWithUserAllOptionsSTforDeletePending() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
			}
			
			
			
			@Test(priority=100202)
			public void checkDeleteRowsinUserLevelforAlertDisappear() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkDeleteRowsinUserLevelforAlertDisappear(), true);
			}
			
			@Test(priority=100203)
			public void checkAlertDisplayedAfterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAlertDisplayedAfterAuthorization(), true);
			}
			
			@Test(priority=100204)
			public void checkLogoutandloginSUAgainafterAuthorizationofMultipleUsersPricebookforAlertAuthforDelete() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandloginSU(), true);
			}
			
			//Approving One Row 
			
			@Test(priority=100205)
			public void checkSavingPricebookfroApprovingOneRow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkSavingPricebookforDeletePendingRow(), true);
			}
			
			@Test(priority=100206)
			public void checkLogoutAndLoginWithUserAllOptionsSTforApprovingOneRow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
			}
			
			
			
			
			@Test(priority=100207)
			public void checkAuthorizingOneRowinSellerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAuthorizingOneRowinSellerPricebook(), true);
			}
			
			
			
			
			@Test(priority=100208)
			public void checkAlertTxtAfterApprovingOneRowinPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAlertTxtAfterApprovingOneRowinPricebook(), true);
			}
			
			@Test(priority=100209)
			public void checkLogoutandloginSUAgainafterAuthorizationofMultipleUsersPricebookforAlertAuthforApproving() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandloginSU(), true);
			}
			
			
			
			@Test(priority=100210)
			public void checkDeletingPendingRowinPricebookAfterAuthofOneRow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkDeletingPendingRowinPricebookAfterAuthofOneRow(), true);
			}
			
					
			
			//Roles
			
			@Test(priority=100211)
			public void checkDisplayingMastersNameInListOfMastersafterEditingAuthforMultipleRoles() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkDisplayingMastersNameInListOfMasters(), true);
			}
			
			@Test(priority=100212)
			public void checkSavedNameDisplayInNameTextBoxafterEditingAuthforforMultipleRoles() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkSavedNameDisplayInNameTextBox(), true);
			}
			
			
			
			
			@Test(priority=100213)
			public void checkEditPriceBookAuthorizationSettingsforMultipleRoles() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
			}
			
			
			
			@Test(priority=100214)
			public void checkSavingPricebookAuthorizationforMultipleRoles() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkSavingPricebookAuthorizationforMultipleRoles(), true);
			}
			
			
			
			@Test(priority=100215)
			public void checkSavingPricebookforMultipleRoles() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkSavingPricebookforMultipleRoles(), true);
			}
			
			
			@Test(priority=100216)
			public void checkStatusofPricebookforMultipleRolesBeforeAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkStatusofPricebookforMultipleRolesBeforeAuthorization(), true);
			}
			
			@Test(priority=100217)
			public void checkLogoutAndLoginWithUserAllOptionsSTforMultipleRoles() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
			}
			
			
			
			@Test(priority=100218)
			public void checkAuthorizeMultipleRolesPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkAuthorizeMultipleRolesPricebook(), true);
			}
			
			@Test(priority=100219)
			public void checkLogoutandloginSUAgainafterAuthorizationofMultipleRolesPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandloginSU(), true);
			}
			
			
			
			@Test(priority=100220)
			public void checkEditingthePricebookRolesAfterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkEditingthePricebookRolesAfterAuthorization(), true);
			}
			
			@Test(priority=100221)
			public void checkLogoutandLoginWithTestwithMultipleRolesAfterEdit() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandLoginWithTestwithMultipleUsers(), true);
			}
			
			
			
			
			@Test(priority=100222)
			public void checkDeleteRowinPricebookofRolesatUserLevel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkDeleteRowinPricebookofRolesatUserLevel(), true);
			}
			
			@Test(priority=100223)
			public void checkLogoutandloginSUAgainafterAuthorizationofEditMultipleRolesPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkLogoutandloginSU(), true);
			}
			
			
			
			@Test(priority=100224)
			public void checkDeletePricebookforMultipleRoles() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkDeletePricebookforMultipleRoles(), true);
			}
			
			
			
			///Update Price List
			
			
			
			@Test(priority=100225)
			public void checkStartandEndDateinPricebbokUpdatePriceList() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				pap=new PricebookAuthorizationPage(getDriver());		
				Assert.assertEquals(pap.checkStartandEndDateinPricebbokUpdatePriceList(), true);
			}
			
			//Unit New scenerios
			
			
			@Test(priority=100226)
			 public void checkRestorePricebookwithunits() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 pap=new PricebookAuthorizationPage(getDriver());
				 Assert.assertEquals(pap.checkRestorePricebookwithunits(), true);
			 }
			
			@Test(priority=100227)
			 public void checkpricebookvaluesinsalesorderforBOXUnitofGroup1Items() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 pap=new PricebookAuthorizationPage(getDriver());
				 Assert.assertEquals(pap.checkpricebookvaluesinsalesorderforBOXUnitofGroup1Items(), true);
			 }
			
			@Test(priority=100228)
			 public void checkpricebookvaluesinsalesorderforPENUnitofGroup1Items() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 pap=new PricebookAuthorizationPage(getDriver());
				 Assert.assertEquals(pap.checkpricebookvaluesinsalesorderforPENUnitofGroup1Items(), true);
			 }
			
			@Test(priority=100229)
			 public void checkpricebookvaluesinsalesorderforSTPUnitofGroup1Items() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 pap=new PricebookAuthorizationPage(getDriver());
				 Assert.assertEquals(pap.checkpricebookvaluesinsalesorderforSTPUnitofGroup1Items(), true);
			 }
			
			@Test(priority=100230)//Except India
			 public void checkDifferentDepartmentwithanyitemshouldnotshowPricebookvalues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 pap=new PricebookAuthorizationPage(getDriver());
				 Assert.assertEquals(pap.checkDifferentDepartmentwithanyitemshouldnotshowPricebookvalues(), true);
			 }
			
			@Test(priority=100231)//Except Customer A
			 public void checkDifferentCustomerAccountwithanyitemshouldnotshowPricebookvalues() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 pap=new PricebookAuthorizationPage(getDriver());
				 Assert.assertEquals(pap.checkDifferentCustomerAccountwithanyitemshouldnotshowPricebookvalues(), true);
			 }
			
			@Test(priority=100232)
			 public void checkpricebookvaluesinsalesorderforDifferentUnitofTestCItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 pap=new PricebookAuthorizationPage(getDriver());
				 Assert.assertEquals(pap.checkpricebookvaluesinsalesorderforDifferentUnitofTestCItem(), true);
			 }
		
			@Test(priority=100233)
			 public void checkpricebookvaluesinsalesorderforDifferentUnitofTestDItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
			 {
				 pap=new PricebookAuthorizationPage(getDriver());
				 Assert.assertEquals(pap.checkpricebookvaluesinsalesorderforDifferentUnitofTestDItem(), true);
			 }
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
		
	//////Buyer Pricebook Authorization
		
	/*	
		

		@Test(priority=100170)
		public void checkSavingBuyerPricebookAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingBuyerPricebookAuthorization(), true);
		}
		
	
		@Test(priority=100171)
		public void checkCreateBuyerPricebookAfterAuthorizationSaved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCreateBuyerPricebookAfterAuthorizationSaved(), true);
		}
		
		
		
		@Test(priority=100172)
		public void checkStatusofBuyerPricebookafterAuthorizationSaved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofBuyerPricebookafterAuthorizationSaved(), true);
		}
		
		
		
		@Test(priority=100173)
		public void checkRateValueloadingbeforeAuthorizationPricebookinPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRateValueloadingbeforeAuthorizationPricebookinPurchaseVoucher(), true);
		}
		
		


		@Test(priority=100174)
		public void checkLogoutAndLoginWithUserAllOptionsSTAfterBuyerPricebookSaved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		@Test(priority=100175)
		public void checkAuthorizationforEditedRowsinBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100176)
		public void checkLogoutandloginSUAfterBuyerPricebookAuthorized() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}

		
		
		
		@Test(priority=100177)
		public void checkRateLoadingafterAuthorizeBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRateLoadingafterAuthorizeBuyerPricebook(), true);
		}
		
		
		@Test(priority=100178)
		public void checkEditPriceBookAuthorizationSettingsforBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		
		
		
		@Test(priority=100179)
		public void checkDateRangeCheckBoxinAuthorizationforBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDateRangeCheckBoxinAuthorizationforBuyer(), true);
		}
		
		
		
		
		
		@Test(priority=100180)
		public void checkEditingtheBuyerPricebookAfterClickingDateRangeinAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditingtheBuyerPricebookAfterClickingDateRangeinAuthorization(), true);
		}
		
	
		
		@Test(priority=100181)
		public void checkStatusofeditBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofeditBuyerPricebook(), true);
		}
		
		
		
		
		@Test(priority=100182)
		public void checkRateLoadingforPendingandAuthorizedPricebookItemsinPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRateLoadingforPendingandAuthorizedPricebookItemsinPurchaseVoucher(), true);
		}
		
		@Test(priority=100183)
		public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeEditedBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		
		@Test(priority=100184)
		public void checkAuthorizationforEditedRowsinBuyerPricebookforDateRangeChkBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100185)
		public void checkLogoutandloginSUAgainofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		
		
		@Test(priority=100186) 
		public void checkRateLoadingforPendingandAuthorizedPricebookItemsinPurchaseVoucherafterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRateLoadingforPendingandAuthorizedPricebookItemsinPurchaseVoucherafterAuthorization(), true);
		}

		
		
		
		@Test(priority=100187)
		public void checkEditBuyerPriceBookAuthorizationSettings() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		
		
		@Test(priority=100188) 
		public void checkConditionCheckboxforBuyerAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkConditionCheckbox(), true);
		}
		
		
		
		@Test(priority=100189) 
		public void checkSavingBuyerPricebookafterPricebookAuthorizewithCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingBuyerPricebookafterPricebookAuthorizewithCondition(), true);
		}
		
		@Test(priority=100190) 
		public void checkStatusofBuyerPricebookforAuthorizationWithRateCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofBuyerPricebookforAuthorizationWithRateCondition(), true);
		}
		
		@Test(priority=100191)
		public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeConditionBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		@Test(priority=100192)
		public void checkConditionedAuthorizationinBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100193)
		public void checkLogoutandloginSUAgainafterAuthofBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		@Test(priority=100194) 
		public void checkStatusofBuyerPricebookAfterAuthorizationWithRateCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofBuyerPricebookAfterAuthorizationWithRateCondition(), true);
		}
		
		
		
	
		@Test(priority=100195) 
		public void checkDeletetheConditionBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheConditionBuyerPriceBook(), true);
		}
	
	
		
		///Product Name
		
		
		@Test(priority=100196)
		public void checkEditBuyerPriceBookAuthorizationSettingsforProduct() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		
		
		
		
		@Test(priority=100197)
		public void checkEditingBuyerPricebookAuthorizationforProductNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditingPricebookAuthorizationforProductNameasCondition(), true);
		}
		
		@Test(priority=100198)
		public void checkSavingBuyerPricebookforProductnameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingBuyerPricebookforProductnameasCondition(), true);
		}
		
		@Test(priority=100199)
		public void checkStatusoftheBuyerPricebookforProductnameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusoftheBuyerPricebookforProductnameasCondition(), true);
		}
		
		
		@Test(priority=100200)
		public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeProductNameConditionBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		@Test(priority=100201)
		public void checkConditionedAuthorizationofProductNameinBuyerrPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100202)
		public void checkLogoutandloginSUAgainafterAuthforProductNameofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		
		
		
		@Test(priority=100203)
		public void checkStatusoftheBuyerPricebookAfterAuthforProductnameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusoftheBuyerPricebookAfterAuthforProductnameasCondition(), true);
		}
		
		
		
		
		
		
		////
		
		@Test(priority=100204)
		public void checkEditBuyerPriceBookAuthorizationSettingsforProductagain() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		
		@Test(priority=100205)
		public void checkUndoAuthorizationEditinginPricebookAuthorizationofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkUndoAuthorizationEditinginPricebookAuthorization(), true);
		}
		
		
		@Test(priority=100206)
		public void checkEditingaBuyerPricebookforUndoAuthEditingOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditingaBuyerPricebookforUndoAuthEditingOption(), true);
		}
		
		
		@Test(priority=100207)
		public void checkLogoutAndLoginWithUserAllOptionsSTsafterEditAuthorizeBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		@Test(priority=100208)
		public void checkConditionedAuthorizationforUndoAuthEditingofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100209)
		public void checkLogoutandloginSUAgainafterEditingAuthofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		@Test(priority=100210)
		public void checkDeletetheConditionProductNameBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheConditionProductNameBuyerPriceBook(), true);
		}
		
		///////////////////////////
		
		
		//Account Name as Condition	
		
		
		@Test(priority=100211)
		public void checkEditBuyerPriceBookAuthorizationSettingsforAccountagain() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		
		@Test(priority=100212)
		public void checkEditingBuyerPricebookAuthorizationforAccountNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditingBuyerPricebookAuthorizationforAccountNameasCondition(), true);
		}
		
		@Test(priority=100213)
		public void checkSavingBuyerPricebookforAccountNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingBuyerPricebookforAccountNameasCondition(), true);
		}
		
		@Test(priority=100214)
		public void checkStatusoftheBuyerPricebookforAccountNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusoftheBuyerPricebookforAccountnameasCondition(), true);
		}
		
		
		@Test(priority=100215)
		public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeAccountNameConditionBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		@Test(priority=100216)
		public void checkConditionedAuthorizationofAccountNameinBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100217)
		public void checkLogoutandloginSUAgainafterAuthforAccountNameofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		
		@Test(priority=100218)
		public void checkStatusoftheBuyerPricebookAfterAuthforAccountnameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusoftheBuyerPricebookAfterAuthforAccountnameasCondition(), true);
		}
		
		@Test(priority=100219)
		public void checkDeletetheConditionAccountNameBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheConditionAccountNameBuyerPriceBook(), true);
		}
		
		
		//Condition as Group Account
		
			
		@Test(priority=100220)
		public void checkEditBuyerPriceBookAuthorizationSettingsforAccountGroupagain() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		@Test(priority=100221)
		public void checkEditingBuyerPricebookAuthorizationforAccountNameasGroupAccCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditingPricebookAuthorizationforAccountNameasGroupAccCondition(), true);
		}
		
		@Test(priority=100222)
		public void checkSavingBuyerPricebookforGroupAccNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingBuyerPricebookforGroupAccNameasCondition(), true);
		}
		
		@Test(priority=100223)
		public void checkStatusoftheBuyerPricebookforGroupAccNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusoftheBuyerPricebookforGroupAccNameasCondition(), true);
		}
		
		
		@Test(priority=100224)
		public void checkRateValueLoadingBeforeAuthofBuyerPricebookforAccountGroup() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRateValueLoadingBeforeAuthofBuyerPricebookforAccountGroup(), true);
		}
		
		@Test(priority=100225)
		public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeAccountGrpNameConditionBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		@Test(priority=100226)
		public void checkConditionedAuthorizationofAccountGrpNameinBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100227)
		public void checkLogoutandloginSUAgainafterAuthforAccountGrpNameofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		@Test(priority=100228)
		public void checkStatusoftheBuyerPricebookAfterAuthforGroupAccNameasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusoftheBuyerPricebookAfterAuthforGroupAccNameasCondition(), true);
		}
		
		
		@Test(priority=100229)
		public void checkDeletetheConditionGroupAccountNameBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheConditionGroupAccountNameBuyerPriceBook(), true);
		}
		
		
		
		
	//Condition as Min 	Qty
		
		
		@Test(priority=100230)
		public void checkEditBuyerPriceBookAuthorizationSettingsforQuantity() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		@Test(priority=100231)
		public void checkEditBuyerPricebookAuthorizationforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPricebookAuthorizationforMinimumQuantityasCondition(), true);
		}
		
		@Test(priority=100232)
		public void checkSavingBuyerPricebookforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingBuyerPricebookforMinimumQuantityasCondition(), true);
		}
		
		@Test(priority=100233)
		public void checkStatusoftheBuyerPricebookforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusoftheBuyerPricebookforMinimumQuantityasCondition(), true);
		}
		
		
		@Test(priority=100234)
		public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeMinQtyConditionBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		@Test(priority=100235)
		public void checkConditionedAuthorizationofMinQtyinBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100236)
		public void checkLogoutandloginSUAgainafterAuthforMinQtyofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		@Test(priority=100237)
		public void checkStatusoftheBuyerPricebookAfterAuthforMinimumQuantityasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusoftheBuyerPricebookAfterAuthforMinimumQuantityasCondition(), true);
		}
		
		@Test(priority=100238)
		public void checkDeletetheConditionMinQtyBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheConditionMinQtyBuyerPriceBook(), true);
		}
		
		
		//Condition as Tag
		
		@Test(priority=100239)
		public void checkSettingsofBuyerPriceBookforTag() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSettingsofBuyerPriceBookforTag(), true);
		}
		
		
		
		
		@Test(priority=100240)
		public void checkEditBuyerPriceBookAuthorizationSettingsforTag() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		
		
		@Test(priority=100241)
		public void checkEditingBuyerPricebookAuthorizationforTagasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditingBuyerPricebookAuthorizationforTagasCondition(), true);
		}
		
		@Test(priority=100242)
		public void checkSavingPricebookforTagasConditionofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingBuyerPricebookforTagasCondition(), true);
		}
		
		@Test(priority=100243)
		public void checkStatusoftheBuyerPricebookforTagasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusoftheBuyerPricebookforTagasCondition(), true);
		}
		
		
		@Test(priority=100244)
		public void checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeTagConditionBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		@Test(priority=100245)
		public void checkConditionedAuthorizationofTaginBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100246)
		public void checkLogoutandloginSUAgainafterAuthforTagofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		
		
		@Test(priority=100247)
		public void checkStatusoftheBuyerPricebookAfterAuthforTagasCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusoftheBuyerPricebookAfterAuthforTagasCondition(), true);
		}
		
		
		
		
		@Test(priority=100248)
		public void checkDeletetheConditionTagBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheConditionTagBuyerPriceBook(), true);
		}
		
		
		
		
		/////////////////Alerts
		
		
		
		@Test(priority=100249)
		public void checkEditBuyerPriceBookAuthorizationSettingsforAlerts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		
		@Test(priority=100250)
		public void checkEditPriceBookAuthorizationSettingsforAlertsofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
		}
		
		
		@Test(priority=100251)
		public void checkEditOnSelectAlertsTabOptionsofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditOnSelectAlertsTabOptions(), true);
		}
		
		
		@Test(priority=100252)
		public void checkSaveWithAllMandatoryInputsofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSaveWithAllMandatoryInputsofBuyerAuth(), true);
		}
		
		
		
		@Test(priority=100253)
		public void checkCreatedEmailInEmailHedaerTemplateofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCreatedEmailInEmailHedaerTemplateofBuyerAuth(), true);
		}
		
		
				
	/*	//@Test(priority=100254)
		public void checkEmailTemplateExpandInFieldsTabOptionsofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEmailTemplateExpandInFieldsTabOptions(), true);
		}
		
		
		
	//	@Test(priority=100112)
		public void checkFieldExpansionDetails() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkFieldExpansionDetails(), true);
		}
		
		
		
	//	@Test(priority=100113)
		public void checkCreationofEmailTemplateinPricebookAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCreationofEmailTemplateinPricebookAuthorization(), true);
		}
		
		@Test(priority=100255)
		public void checkCreatedEmailTemplateofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCreatedEmailTemplate(), true);
		}
		
		//@Test(priority=100115)
		public void checkSaveAlertsTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSaveAlertsTemplate(), true);
		}
		
	/*	@Test(priority=100256)
		public void checkCreatedAlertTemplateofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkCreatedAlertTemplateofBuyer(), true);
		}
		
		
		@Test(priority=100257)
		public void checkSavingABuyerPricebookafterAlertsCreated() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingABuyerPricebookafterAlertsCreated(), true);
		}
		
		
		
		
		@Test(priority=100258)
		public void checkStatusofBuyerPricebookforAlerts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofBuyerPricebookforAlerts(), true);
		}
		
		
		@Test(priority=100259)
		public void checkLogoutAndLoginWithUserAllOptionsSTforAlertsofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		
		
		@Test(priority=100260)
		public void checkConditionedAuthorizationforAlertofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100261)
		public void checkLogoutandloginSUAgainafterAlertsReceivedofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		@Test(priority=100262)
		public void checkStatusofBuyerPricebookforAlertsAfterAuthorize() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofBuyerPricebookforAlertsAfterAuthorize(), true);
		}
		
		@Test(priority=100263)
		public void checkDeletetheAlertBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheAlertBuyerPriceBook(), true);
		}
		
		
		//Approved
		
		@Test(priority=100264)
		public void checkEditBuyerPriceBookAuthorizationSettingsforAlertsApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		
		@Test(priority=100265)
		public void checkEditBuyerPriceBookAuthorizationSettingsforAlertsApprovedclick() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
		}
		
		
		//@Test(priority=100251)
		public void checkEditOnSelectAlertsTabOptionsofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditOnSelectAlertsTabOptions(), true);
		}
		
		
		
		
		
		
		
		@Test(priority=100266)
		public void checkSelectCreatedAlertTemplateforApprovedofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSelectCreatedAlertTemplateforApprovedofBuyerAuth(), true);
		}
		
		
		
		
		@Test(priority=100267)
		public void checkRecipientforAlertsApprovedofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkRecipientforAlertsApproved(), true);
		}
		
		
		@Test(priority=100268)
		public void checkAgainSavingABuyerPricebookafterAlertsCreatedforApprove() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkSavingABuyerPricebookafterAlertsCreated(), true);
		}
		
		
		
		
		@Test(priority=100269)
		public void checkStatusofBuyerPricebookforAlertsagain() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkStatusofBuyerPricebookforAlerts(), true);
		}
		
		
		@Test(priority=100270)
		public void checkLogoutAndLoginWithUserAllOptionsSTforAlertsAgainofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
		}
		
		@Test(priority=100271)
		public void checkConditionedAuthorizationforAlertsApprovedofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAuthorizationforEditedRowsinBuyerPricebook(), true);
		}
		
		

		@Test(priority=100272)
		public void checkLogoutandloginSUAgainafterAlertsApprovedofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkLogoutandloginSU(), true);
		}
		
		
		
		
		@Test(priority=100273)
		public void checkAlertinOwnerAfterApprovedofAuthorizationofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkAlertinOwnerAfterApprovedofAuthorizationofBuyer(), true);
		}
		
		
		@Test(priority=100274)
		public void checkDeletetheAlertBuyerPriceBookforApproved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkDeletetheAlertBuyerPriceBook(), true);
		}
		
		
		
		//Rejected
		
		
		@Test(priority=100275)
		public void checkEditBuyerPriceBookAuthorizationSettingsforAlertsRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
		}
		
		
		@Test(priority=100276)
		public void checkEditBuyerPriceBookAuthorizationSettingsforAlertsRejectedclick() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			pap=new PricebookAuthorizationPage(getDriver());		
			Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
		}
				
		
				
				
				
				
				@Test(priority=100278)
				public void checkSelectCreatedAlertTemplateforRejectedofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkSelectCreatedAlertTemplateforRejectedofBuyerAuth(), true);
				}
				
				
				
				
				@Test(priority=100279)
				public void checkRecipientforAlertsRejectedofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkRecipientforAlertsRejected(), true);
				}
				
				
				@Test(priority=100280)
				public void checkAgainSavingABuyerPricebookafterAlertsCreatedforRejected() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkSavingABuyerPricebookafterAlertsCreated(), true);
				}
				
				
				
				
				@Test(priority=100281)
				public void checkStatusofBuyerPricebookforAlertsagainafterReject() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkStatusofBuyerPricebookforAlerts(), true);
				}
				
				
				@Test(priority=100282)
				public void checkLogoutAndLoginWithUserAllOptionsSTforAlertsAgainAfterRejectofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
				}
				
				@Test(priority=100283)
				public void checkRejectionforEditedRowsinBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkRejectionforEditedRowsinBuyerPricebook(), true);
				}
				
				

				@Test(priority=100284)
				public void checkLogoutandloginSUAgainafterAlertsRejectedofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkLogoutandloginSU(), true);
				}
				
				
				@Test(priority=100285)
				public void checkStatusoftheBuyerPricebbokafterRejection() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkStatusoftheBuyerPricebbokafterRejection(), true);
				}
				
				
				
				@Test(priority=100286)
				public void checkAlertinOwnerAfterRejectedofAuthorizationofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkAlertinOwnerAfterApprovedofAuthorizationofBuyer(), true);
				}
				
				
				@Test(priority=100287)
				public void checkDeletetheAlertBuyerPriceBookAfterRejection() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkDeletetheAlertBuyerPriceBook(), true);
				}
				
				
				
				////Stopped
				
				
				
				
				@Test(priority=100288)
				public void checkEditBuyerPriceBookAuthorizationSettingsforAlertsStoppedAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettingsforBuyer(), true);
				}
				
				
				@Test(priority=100289)
				public void checkEditBuyerPriceBookAuthorizationSettingsforAlertsStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkEditPriceBookAuthorizationSettings(), true);
				}
						
				
				
				
				
				
				
				
				@Test(priority=100290)
				public void checkSelectCreatedAlertTemplateforStoppedofBuyerAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkSelectCreatedAlertTemplateforStoppedofBuyerAuth(), true);
				}
				
				
				
				
				@Test(priority=100291)
				public void checkRecipientforAlertsStoppedofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkRecipientforAlertsStopped(), true);
				}
				
				
				@Test(priority=100292)
				public void checkAgainSavingABuyerPricebookafterAlertsCreatedforStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkSavingABuyerPricebookafterAlertsCreated(), true);
				}
				
				
				
				
				@Test(priority=100293)
				public void checkStatusofBuyerPricebookforAlertsagainafterStoped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkStatusofBuyerPricebookforAlerts(), true);
				}
				
				
				@Test(priority=100294)
				public void checkLogoutAndLoginWithUserAllOptionsSTforAlertsAgainAfterStopedofBuyer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkLogoutAndLoginWithUserAllOptionsSTToAuthorizeSellerPricebook(), true);
				}
				
				@Test(priority=100295)
				public void checkStopedforEditedRowsinBuyerPricebook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkStopedforEditedRowsinBuyerPricebook(), true);
				}
				
				

				@Test(priority=100296)
				public void checkLogoutandloginSUAgainafterAlertsStoppedofBuyerAuthStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkLogoutandloginSU(), true);
				}
				
				
				@Test(priority=100297)
				public void checkStatusoftheBuyerPricebbokafterStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkStatusoftheBuyerPricebbokafterStopped(), true);
				}
				
				
				
				@Test(priority=100298)
				public void checkAlertinOwnerAfterStoppedofAuthorizationofBuyerStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkAlertinOwnerAfterApprovedofAuthorizationofBuyer(), true);
				}
				
				
				@Test(priority=100299)
				public void checkDeletetheAlertBuyerPriceBookAfterStopped() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkDeletetheAlertBuyerPriceBook(), true);
				}
				
				*/
				
				@Test(priority=100300)
				public void checkLogoutPricebookAuthorizationPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
				{
					pap=new PricebookAuthorizationPage(getDriver());		
					Assert.assertEquals(pap.checkLogoutPricebookAuthorizationPage(), true);
				}
		

}

