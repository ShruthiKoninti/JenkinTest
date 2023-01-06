package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.OpeningBalancesPage;
import com.focus.base.BaseEngine;

public class OpeningBalancesTest extends BaseEngine
{
	OpeningBalancesPage OP;

/*
	@Test(priority=230)
	public void checkOpeningBalanceCustomizationRemovingLinkStatusColumn() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkOpeningBalanceCustomizationRemovingLinkStatusColumn(), true);
	}

	@Test(priority=231)//issue not hidding Col
	public void checkOpeningBalanceCustomizationHideBalanceLinkValueColumn() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkOpeningBalanceCustomizationHideBalanceLinkValueColumn(), true);
	}

	@Test(priority=232)
	public void checkOpeningBalanceCustomizationAddingDepartmentColumn() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkOpeningBalanceCustomizationAddingDepartmentColumn(), true);
	}


*/

	@Test(priority=241)
	public void checkSavingOpeningBalanceVoucher1WithCurrencyAsINR() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkSavingOpeningBalanceVoucher1WithCurrencyAsINR(), true);
	}

	@Test(priority=242)
	public void checkVerifyAndDeleteOpeningBalanceVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkVerifyAndDeleteOpeningBalanceVoucher1(), true);
	}

	@Test(priority=243)
	public void checkSavingOpeningBalanceVoucher1WithCurrencyAsAED() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkSavingOpeningBalanceVoucher1WithCurrencyAsAED(), true);
	}

	@Test(priority=244)
	public void checkVerifyAndEditOpeningBalanceVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkVerifyAndEditOpeningBalanceVoucher1(), true);
	}

	@Test(priority=245)
	public void checkSavingOpeningBalanceVoucher2AndSuspendingFromEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkSavingOpeningBalanceVoucher2AndSuspendingFromEntryPage(), true);
	}
	
	@Test(priority=246)
	public void checkResavingOpeningBalanceVoucher2AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkResavingOpeningBalanceVoucher2AndSuspendingFromHomepage(), true);
	}
	
	@Test(priority=247)
	public void checkDeleteOptionInOpeningBalanceHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkDeleteOptionInOpeningBalanceHomePage(), true);
	}
	
	@Test(priority=248)
	public void checkVerifingDetailsOfSavedOpeningBalancesVouchersInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		OP=new OpeningBalancesPage(getDriver());
		Assert.assertEquals(OP.checkVerifingDetailsOfSavedOpeningBalancesVouchersInHomepage(), true);
	}












}
