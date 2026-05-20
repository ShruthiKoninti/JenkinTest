package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.DebitNotesPage;
import com.focus.base.BaseEngine;

public class DebitNotesTest extends BaseEngine
{
	DebitNotesPage DN;
	
	
	@Test(priority=330)
	public void CheckLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		DN=new DebitNotesPage(getDriver());
		Assert.assertEquals(DN.CheckLogin(), true);
	}

	@Test(priority=332)
	public void checkCustomizationOptionInDebitNotes() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		DN=new DebitNotesPage(getDriver());
		Assert.assertEquals(DN.checkCustomizationOptionInDebitNotes(), true);
	}

	@Test(priority=333)
	public void checkSavingDebitNotesVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		DN=new DebitNotesPage(getDriver());
		Assert.assertEquals(DN.checkSavingDebitNotesVoucher1(), true);
	}

	@Test(priority=334)
	public void checkEditingDebitNotesVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		DN=new DebitNotesPage(getDriver());
		Assert.assertEquals(DN.checkEditingDebitNotesVoucher1(), true);
	}

	@Test(priority=335)
	public void checkSavingDebitNotesVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		DN=new DebitNotesPage(getDriver());
		Assert.assertEquals(DN.checkSavingDebitNotesVoucher2(), true);
	}

	@Test(priority=336)
	public void checkDeletingDebitNotesVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		DN=new DebitNotesPage(getDriver());
		Assert.assertEquals(DN.checkDeletingDebitNotesVoucher2(), true);
	}


	@Test(priority=337)
	public void checkVerifingDetailsOfSavedDebitNotesVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		DN=new DebitNotesPage(getDriver());
		Assert.assertEquals(DN.checkVerifingDetailsOfSavedDebitNotesVoucherInHomepage(), true);
	}

	
	
	@Test(priority=339)
	public void checkLogoutDebitNotesPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		DN=new DebitNotesPage(getDriver());
		Assert.assertEquals(DN.checkLogoutDebitNotesPage(), true);
	}

	
	
}
