package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.SalesInvoiceVATVoucherPage;
import com.focus.base.BaseEngine;

public class SalesInvoiceVATVoucherTest extends BaseEngine
{
	SalesInvoiceVATVoucherPage SIV;
/*	
	@Test(priority=307)
	public void checkSalesInvoiceVoucherToCheckLinkPopUpScreenWithCRTlPlusL() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSalesInvoiceVoucherToCheckLinkPopUpScreenWithCRTlPlusL(), true);
	}

	@Test(priority=308)
	public void checkSalesInvoiceVoucherSearchInLinkPopupScreenAndSaveWithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSalesInvoiceVoucherSearchInLinkPopupScreenAndSaveWithBatchItem(), true);
	}

	@Test(priority=309)
	public void checkSalesInvoiceVoucherLinewiseConsumingForBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSalesInvoiceVoucherLinewiseConsumingForBatchItem(), true);
	}

	@Test(priority=310)
	public void checkSuspendOptionInSalesInvoiceEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSuspendOptionInSalesInvoiceEntryPage(), true);
	}

	@Test(priority=311)
	public void checkSalesInvoiceVoucherToCheckBinItemLinkPopUpScreenWithCRTlPlusL() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSalesInvoiceVoucherToCheckBinItemLinkPopUpScreenWithCRTlPlusL(), true);
	}

	@Test(priority=312) 
	public void checkSalesInvoiceVoucherLinewiseConsumingForBinItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSalesInvoiceVoucherLinewiseConsumingForBinItem(), true);
	}

	@Test(priority=313)
	public void checkSuspendOptionInSalesInvoiceHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSuspendOptionInSalesInvoiceHomePage(), true);
	}
*/
	@Test(priority=314) // Trans Negative Message
	public void checkSalesInvoiceVoucherToCheckRMAItemLinkPopUpScreenWithCRTlPlusL() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSalesInvoiceVoucherToCheckRMAItemLinkPopUpScreenWithCRTlPlusL(), true);
	}

	@Test(priority=315)//RMA Outward
	public void checkSalesInvoiceVoucherLinewiseConsumingForRmaItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSalesInvoiceVoucherLinewiseConsumingForRmaItem(), true);
	}

	@Test(priority=316)//check it
	public void checkDeleteOptionInSalesInvoiceEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkDeleteOptionInSalesInvoiceEntryPage(), true);
	}

	@Test(priority=317)
	public void checkSalesInvoiceVoucherToCheckStockItemLinkPopUpScreenWithCRTlPlusL() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSalesInvoiceVoucherToCheckStockItemLinkPopUpScreenWithCRTlPlusL(), true);
	}

	@Test(priority=318) 
	public void checkSalesInvoiceVoucherLinewiseConsumingForStockItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSalesInvoiceVoucherLinewiseConsumingForStockItem(), true);
	}

	@Test(priority=319)
	public void checkDeleteOptionInSalesInvoiceHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkDeleteOptionInSalesInvoiceHomePage(), true);
	}


	@Test(priority=320)
	public void checkSalesInvoiceVoucherSavingSuspenigAndDeletingThroughConvertOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSalesInvoiceVoucherSavingSuspenigAndDeletingThroughConvertOption(), true);
	}


	@Test(priority=321)
	public void checkSavingSalesInvoiceWithhAllItemsOfTransferedStockInSingleVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSavingSalesInvoiceWithhAllItemsOfTransferedStockInSingleVoucher(), true);
	}


	@Test(priority=322)
	public void checkVerifingDetailsOfSavedSalesInvoiceVATVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkVerifingDetailsOfSavedSalesInvoiceVATVoucherInHomepage(), true);
	}



	@Test(priority=323)
	public void checkSavingSalesReturnsVoucherWithConvertOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSavingSalesReturnsVoucherWithConvertOptions(), true);
	}


	@Test(priority=324)
	public void checkVerifingDetailsOfSavedSalesReturnsVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkVerifingDetailsOfSavedSalesReturnsVoucherInHomepage(), true);
	}



	@Test(priority=325)
	public void checkVoucherFooterDetailsatHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkVoucherFooterDetailsatHomePage(), true);
	}


	@Test(priority=326)
	public void checkSearchingCustomerandItemsListwithoutAstrickSigninSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSearchingCustomerandItemsListwithoutAstrickSigninSalesInvoiceVAT(), true);
	}

	
	@Test(priority=327)
	public void checkSearchingofVendorandItemsListusingAstricksigninSalesInvoiceVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		SIV=new SalesInvoiceVATVoucherPage(getDriver());
		Assert.assertEquals(SIV.checkSearchingofVendorandItemsListusingAstricksigninSalesInvoiceVAT(), true);
	}

	
	
	
	
}
