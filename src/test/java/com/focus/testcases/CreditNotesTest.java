package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.CreditNotesPage;
import com.focus.base.BaseEngine;

public class CreditNotesTest extends BaseEngine
{
	CreditNotesPage CN;

	//@Test(priority=327)
	public void checkCustomizationOptionInCreditNotes() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		CN=new CreditNotesPage(getDriver());
		Assert.assertEquals(CN.checkCustomizationOptionInCreditNotes(), true);
	}

	@Test(priority=328)
	public void checkSavingCreditNotesVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		CN=new CreditNotesPage(getDriver());
		Assert.assertEquals(CN.checkSavingCreditNotesVoucher1(), true);
	}

	@Test(priority=329)
	public void checkEditingCreditNotesVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		CN=new CreditNotesPage(getDriver());
		Assert.assertEquals(CN.checkEditingCreditNotesVoucher1(), true);
	}

	@Test(priority=330)
	public void checkSavingCreditNotesVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		CN=new CreditNotesPage(getDriver());
		Assert.assertEquals(CN.checkSavingCreditNotesVoucher2(), true);
	}

	@Test(priority=331)
	public void checkDeletingCreditNotesVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		CN=new CreditNotesPage(getDriver());
		Assert.assertEquals(CN.checkDeletingCreditNotesVoucher2(), true);
	}

	@Test(priority=332)
	public void checkVerifingDetailsOfSavedCreditNotesVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		CN=new CreditNotesPage(getDriver());
		Assert.assertEquals(CN.checkVerifingDetailsOfSavedCreditNotesVoucherInHomepage(), true);
	}





}
