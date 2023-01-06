package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.ReceiptsVATVoucherPage;
import com.focus.Pages.ReceiptsVATVoucherPage;
import com.focus.base.BaseEngine;

public class ReceiptsVATVoucherTest extends BaseEngine
{
	ReceiptsVATVoucherPage RV;

	@Test(priority=323)
	public void checkSavingRecepitsVATVocher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RV=new ReceiptsVATVoucherPage(getDriver());
		Assert.assertEquals(RV.checkSavingRecepitsVATVocher1(), true);
	}

	@Test(priority=324)
	public void checkEditingRecepitsVATVocher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RV=new ReceiptsVATVoucherPage(getDriver());
		Assert.assertEquals(RV.checkEditingRecepitsVATVocher1(), true);
	}

	@Test(priority=325)
	public void checkSavingRecepitsVATVocher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RV=new ReceiptsVATVoucherPage(getDriver());
		Assert.assertEquals(RV.checkSavingRecepitsVATVocher2(), true);
	}

	@Test(priority=326)
	public void checkDeletingRecepitsVATVocher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RV=new ReceiptsVATVoucherPage(getDriver());
		Assert.assertEquals(RV.checkDeletingRecepitsVATVocher2(), true);
	}

	@Test(priority=327)
	public void checkVerifingDetailsOfSavedReceiptsVATVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		RV=new ReceiptsVATVoucherPage(getDriver());
		Assert.assertEquals(RV.checkVerifingDetailsOfSavedReceiptsVATVoucherInHomepage(), true);
	}

}
