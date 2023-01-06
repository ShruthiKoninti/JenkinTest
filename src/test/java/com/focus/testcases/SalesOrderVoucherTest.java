package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.SalesOrderVoucherPage;
import com.focus.base.BaseEngine;

public class SalesOrderVoucherTest extends BaseEngine
{
	SalesOrderVoucherPage SO;
/*
	//@Test(priority=289)
	public void checkSalesOrderVoucherCustomization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkSalesOrderVoucherCustomization(), true);
	}

	@Test(priority=290)
	public void checkSavingSalesOrderVoucher1WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkSavingSalesOrderVoucher1WithBatchItem(), true);
	}

	@Test(priority=291)
	public void checkEditingSalesOrderVoucher1AndAdding2ndRow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkEditingSalesOrderVoucher1AndAdding2ndRow(), true);
	}

	@Test(priority=292)
	public void checkSavingSalesOrderVoucher2WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkSavingSalesOrderVoucher2WithBatchItem(), true);
	}

	@Test(priority=293)
	public void checkVerifyingSalesOrderVoucher2AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkVerifyingSalesOrderVoucher2AndDeleting(), true);
	}



	@Test(priority=294)
	public void checkSavingSalesOrderVoucher2WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkSavingSalesOrderVoucher2WithStockItem(), true);
	}

	@Test(priority=295)
	public void checkEditingSalesOrderVoucher2AndAdding2ndRow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkEditingSalesOrderVoucher2AndAdding2ndRow(), true);
	}

	@Test(priority=296)
	public void checkSavingSalesOrderVoucher3WithStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkSavingSalesOrderVoucher3WithStockItem(), true);
	}

	@Test(priority=297)
	public void checkVerifyingSalesOrderVoucher3AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkVerifyingSalesOrderVoucher3AndDeleting(), true);
	}

	@Test(priority=298)
	public void checkSavingSalesOrderVoucherWithBinItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkSavingSalesOrderVoucherWithBinItem(), true);
	}


	@Test(priority=299)
	public void checkEditingSalesOrderVoucher3AndAdding2ndRow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkEditingSalesOrderVoucher3AndAdding2ndRow(), true);
	}
*/

	@Test(priority=300)
	public void checkSavingSalesOrderVoucher4WithBinItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkSavingSalesOrderVoucher4WithBinItem(), true);
	}

	@Test(priority=301)
	public void checkVerifyingSalesOrderVoucher4AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkVerifyingSalesOrderVoucher4AndDeleting(), true);
	}



	@Test(priority=302)
	public void checkSavingSalesOrderVoucherWithRmaItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkSavingSalesOrderVoucherWithRmaItem(), true);
	}


	@Test(priority=303) // Negative Stocks Error Message
	public void checkEditingSalesOrderVoucher4AndAdding2ndRow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkEditingSalesOrderVoucher4AndAdding2ndRow(), true);
	}

	@Test(priority=304) // Negative Stocks Error Message
	public void checkSavingSalesOrderVoucher5WithRmaItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkSavingSalesOrderVoucher5WithRmaItem(), true);
	}

	@Test(priority=305)
	public void checkVerifyingSalesOrderVoucher5AndDeleting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkVerifyingSalesOrderVoucher5AndDeleting(), true);
	}

	@Test(priority=306)
	public void checkSavingAndSuspendingSalesOrderVoucher5FromEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkSavingAndSuspendingSalesOrderVoucher5FromEntryPage(), true);
	}
	
	@Test(priority=307)
	public void checkResavingSalesOrderVoucher5AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkResavingSalesOrderVoucher5AndSuspendingFromHomepage(), true);
	}
	
	@Test(priority=308)
	public void checkDeleteOptionInSalesOrderHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkDeleteOptionInSalesOrderHomePage(), true);
	}
	
	@Test(priority=309)
	public void checkVerifingDetailsOfSavedSalesOrderVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SO=new SalesOrderVoucherPage(getDriver());
		Assert.assertEquals(SO.checkVerifingDetailsOfSavedSalesOrderVoucherInHomepage(), true);
	}



}
