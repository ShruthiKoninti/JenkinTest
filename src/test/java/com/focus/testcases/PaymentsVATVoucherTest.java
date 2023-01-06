package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.PaymentsVATVoucherPage;
import com.focus.base.BaseEngine;

public class PaymentsVATVoucherTest extends BaseEngine
{

	PaymentsVATVoucherPage PV;

	@Test(priority=276)
	public void checkSavingPaymentVATVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		PV=new PaymentsVATVoucherPage(getDriver());
		Assert.assertEquals(PV.checkSavingPaymentVATVoucher1(), true);
	}

	@Test(priority=277)
	public void checkEditingPaymentVATVoucher1AndAdding2ndRow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		PV=new PaymentsVATVoucherPage(getDriver());
		Assert.assertEquals(PV.checkEditingPaymentVATVoucher1AndAdding2ndRow(), true);
	}

	@Test(priority=278)
	public void checkSavingPaymentVATVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		PV=new PaymentsVATVoucherPage(getDriver());
		Assert.assertEquals(PV.checkSavingPaymentVATVoucher2(), true);
	}

	@Test(priority=279)
	public void checkVerifyingSavedPaymentVATVoucher2AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		PV=new PaymentsVATVoucherPage(getDriver());
		Assert.assertEquals(PV.checkVerifyingSavedPaymentVATVoucher2AndDeleting(), true);
	}

	@Test(priority=280)
	public void checkVerifingDetailsOfSavedPaymentVATVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		PV=new PaymentsVATVoucherPage(getDriver());
		Assert.assertEquals(PV.checkVerifingDetailsOfSavedPaymentVATVoucherInHomepage(), true);
	}

}
