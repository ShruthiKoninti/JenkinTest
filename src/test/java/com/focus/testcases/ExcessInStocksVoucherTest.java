package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.ExcessInStocksVoucherPage;
import com.focus.Pages.ExcessInStocksVoucherPage;
import com.focus.base.BaseEngine;

public class ExcessInStocksVoucherTest extends BaseEngine
{
	ExcessInStocksVoucherPage EIS;

	@Test(priority=280)
	public void checkSavingExcessInStocksVoucherWithAllItems() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		EIS=new ExcessInStocksVoucherPage(getDriver());
		Assert.assertEquals(EIS.checkSavingExcessInStocksVoucherWithAllItems(), true);
	}

	@Test(priority=281)
	public void checkCopyDocumentOptionAndSavingInExcessInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		EIS=new ExcessInStocksVoucherPage(getDriver());
		Assert.assertEquals(EIS.checkCopyDocumentOptionAndSavingInExcessInStocks(), true);
	}

	@Test(priority=282)
	public void checkUpdatingSavedVoucherInExcessInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		EIS=new ExcessInStocksVoucherPage(getDriver());
		Assert.assertEquals(EIS.checkUpdatingSavedVoucherInExcessInStocks(), true);
	}

	@Test(priority=283)
	public void checkDeletingVoucherInExcessInStocks() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		EIS=new ExcessInStocksVoucherPage(getDriver());
		Assert.assertEquals(EIS.checkDeletingVoucherInExcessInStocks(), true);
	}

	@Test(priority=284)
	public void checkSavingAndSuspendingExcessInStocksVoucherInEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		EIS=new ExcessInStocksVoucherPage(getDriver());
		Assert.assertEquals(EIS.checkSavingAndSuspendingExcessInStocksVoucherInEntryPage(), true);
	}
	
	@Test(priority=285)
	public void checkResavingExcessesInStocksVoucher3AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		EIS=new ExcessInStocksVoucherPage(getDriver());
		Assert.assertEquals(EIS.checkResavingExcessesInStocksVoucher3AndSuspendingFromHomepage(), true);
	}
	
	@Test(priority=286)
	public void checkDeletingExcessesInStockVoucher3FromHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		EIS=new ExcessInStocksVoucherPage(getDriver());
		Assert.assertEquals(EIS.checkDeletingExcessesInStockVoucher3FromHomePage(), true);
	}

	@Test(priority=287)
	public void checkVerifingDetailsOfSavedExcessInStocksVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		EIS=new ExcessInStocksVoucherPage(getDriver());
		Assert.assertEquals(EIS.checkVerifingDetailsOfSavedExcessInStocksVoucherInHomepage(), true);
	}

}
