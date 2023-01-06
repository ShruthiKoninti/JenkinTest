package com.focus.testcases;

import org.testng.annotations.Test;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;

import com.focus.Pages.COGSPostingPage;
import com.focus.base.BaseEngine;

public class COGSPostingTest extends BaseEngine
{
	COGSPostingPage CPP;
	
	@Test(priority=1)
	public void checkLoginCompanyWithValidCredentials() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLoginCompanyWithValidCredentials(), true);
	}
	// Weighted Average
	
	@Test(priority=54321)
	public void checkSavingPurchasesOrderVoucher1WithAccountInEveryLineAndDetailPosting() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchasesOrderVoucher1WithAccountInEveryLineAndDetailPosting(), true);
	}
	
	//@Test(priority=54322)
	public void checkLogoutAndLoginWithUserAllOptionsSTForAuthorizingPurchasesOrdersVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLogoutAndLoginWithUserAllOptionsST(), true);
	}
	
	//@Test(priority=54323)
	public void checkAuthorizationInPurchasesOrdersVoucherWithUserAllOptionsST() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkAuthorizationInPurchasesOrdersVoucherWithUserAllOptionsST(), true);
	}
	
	//@Test(priority=54324)
	public void checkLogoutAndLoginWithSUAfterAuthorizingPurchasesOrdersVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLogoutAndLoginWithSU(), true);
	}
	
	@Test(priority=54325)
	public void checkSavingMaterialReceiptNotesWithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesWithoutUpdateFA(), true);
	}
	
	@Test(priority=54326)
	public void checkLedgerReportAfterSavingMaterialReceiptNotesWithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingMaterialReceiptNotesWithoutUpdateFA(), true);
	}
	
	@Test(priority=54327)
	public void checkReSavingMaterialReceiptNotesWithUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkReSavingMaterialReceiptNotesWithUpdateFA(), true);
	}
	
	@Test(priority=54328)
	public void checkPostingDetailsAfterSavingMaterialReceiptNotesWithUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingMaterialReceiptNotesWithUpdateFA(), true);
	}
	
	@Test(priority=54329)
	public void checkLedgerReportAfterSavingMaterialReceiptNotesWithUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingMaterialReceiptNotesWithUpdateFA(), true);
	}
	
	@Test(priority=54330)
	public void checkSavingPurchaseVoucherVoucher1WithUpdateFAAndUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucherVoucher1WithUpdateFAAndUpdateStock(), true);
	}
	
	@Test(priority=54331)
	public void checkPostingDetailsAfterSavingPurchasesVoucherWithUpdateFAAndUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesVoucherWithUpdateFAAndUpdateStock(), true);
	}
	
	@Test(priority=54332)
	public void checkLedgerReportAfterSavingPurchasesVoucherWithUpdateFAAndUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingPurchasesVoucherWithUpdateFAAndUpdateStock(), true);
	}

	@Test(priority=54333)
	public void checkDeletingPurchaseVoucherAndMaterialReceiptNotesVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkDeletingPurchaseVoucherAndMaterialReceiptNotesVoucher(), true);
	}
	
	@Test(priority=54334)
	public void checkSavingMaterialReceiptNotesWithoutUpdateFAAndWithUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesWithoutUpdateFAAndWithUpdateStock(), true);
	}
	
	@Test(priority=54335)
	public void checkStockLedgerReportAfterSavingMaterialReceiptNotesWithoutUpdateFAAndWithUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingMaterialReceiptNotesWithoutUpdateFAAndWithUpdateStock(), true);
	}
	
	@Test(priority=54336)
	public void checkSavingPurchaseVoucherWithUpdateFAAndWithoutUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucherWithUpdateFAAndWithoutUpdateStock(), true);
	}
	
	@Test(priority=54337)
	public void checkLedgerReportAfterSavingPurchasesVoucherWithUpdateFAAndWithoutUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingPurchasesVoucherWithUpdateFAAndWithoutUpdateStock(), true);
	}
	
	@Test(priority=54338)
	public void checkSavingPurchaseReturnsVoucherWithQuantity1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseReturnsVoucherWithQuantity1(), true);
	}
	
	@Test(priority=54339)
	public void checkPostingDetailsAfterSavingPurchaseReturnsVoucherWithQuantity1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchaseReturnsVoucherWithQuantity1(), true);
	}
	
	@Test(priority=54340)
	public void checkLedgerReportAfterSavingPurchasesReturnsWithQuantity1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingPurchasesReturnsWithQuantity1(), true);
	}
	
	@Test(priority=54341)
	public void checkStockLedgerReportAfterSavingPurchasesReturnsVoucherWithQuantity1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingPurchasesReturnsVoucherWithQuantity1(), true);
	}
	
	@Test(priority=54342)
	public void checkSavingDeliveryNotesVoucherWithUpdateStockAndWithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingDeliveryNotesVoucherWithUpdateStockAndWithoutUpdateFA(), true);
	}
	
	@Test(priority=54343)
	public void checkPostingDetailsAfterSavingDeliveryNotesVoucherWithUpdateStockAndWithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingDeliveryNotesVoucherWithUpdateStockAndWithoutUpdateFA(), true);
	}
	
	@Test(priority=54344)
	public void checkReSavingDeliveryNotesVoucherWithUpdateStockAndUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkReSavingDeliveryNotesVoucherWithUpdateStockAndUpdateFA(), true);
	}
	
	@Test(priority=54345)
	public void checkPostingDetailsAfterReSavingDeliveryNotesVoucherWithUpdateStockAndUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterReSavingDeliveryNotesVoucherWithUpdateStockAndUpdateFA(), true);
	}
	
	@Test(priority=54346)
	public void checkReSavingDeliveryNotesVoucherWithUpdateStockAndWithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkReSavingDeliveryNotesVoucherWithUpdateStockAndWithoutUpdateFA(), true);
	}
	
	@Test(priority=54347)
	public void checkPostingDetailsAfterReSavingDeliveryNotesVoucherWithUpdateStockAndWithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterReSavingDeliveryNotesVoucherWithUpdateStockAndWithoutUpdateFA(), true);
	}
	
	@Test(priority=54348)
	public void checkSavingSalesInvoiceVoucherWithoutUpdateStockAndWithUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesInvoiceVoucherWithoutUpdateStockAndWithUpdateFA(), true);
	}
	
	@Test(priority=54349)
	public void checkPostingDetailsAfterSavingSalesInvoiceVoucherWithoutUpdateStockAndWithUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesInvoiceVoucherWithoutUpdateStockAndWithUpdateFA(), true);
	}
	
	@Test(priority=54350)
	public void checkReSavingSalesInvoiceVoucherWithUpdateStockAndWithUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkReSavingSalesInvoiceVoucherWithUpdateStockAndWithUpdateFA(), true);
	}
	
	@Test(priority=54351)
	public void checkPostingDetailsAfterReSavingSalesInvoiceVoucherWithUpdateStockAndWithUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterReSavingSalesInvoiceVoucherWithUpdateStockAndWithUpdateFA(), true);
	}
	
	@Test(priority=54352)
	public void checkSavingSalesReturnsVoucherWithUpdateStockAndUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesReturnsVoucherWithUpdateStockAndUpdateFA(), true);
	}
	
	@Test(priority=54353)
	public void checkPostingDetailsAfterSavingSalesReturnsVoucherWithUpdateStockAndUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesReturnsVoucherWithUpdateStockAndUpdateFA(), true);
	}
	
	@Test(priority=54354)
	public void checkSavingExcessesInStocksVoucherWithUpdateStockAndUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingExcessesInStocksVoucherWithUpdateStockAndUpdateFA(), true);
	}
	
	@Test(priority=54355)
	public void checkPostingDetailsAfterSavingExcessesInStocksVoucherWithUpdateStockAndUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingExcessesInStocksVoucherWithUpdateStockAndUpdateFA(), true);
	}
	
	@Test(priority=54356)
	public void checkSavingShortageInStocksVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingShortageInStocksVoucher(), true);
	}
	
	@Test(priority=54357)
	public void checkPostingDetailsAfterSavingShortageInStocksVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingShortageInStocksVoucher(), true);
	}
	
	@Test(priority=54358)
	public void checkLedgerReportAfterSavingExcessesAndShortageInStocksVouchers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingExcessesAndShortageInStocksVouchers(), true);
	}
	
	@Test(priority=54359)
	public void checkTrailBalanceReportAfterExcessAndShortageInStocksVouchers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkTrailBalanceReportAfterExcessAndShortageInStocksVouchers(), true);
	}
	
	@Test(priority=54360)
	public void checkStockLedgerReportAfterSavingExcessesAndShortageInStocksVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingExcessesAndShortageInStocksVoucher(), true);
	}
	
	@Test(priority=54361)
	public void checkStockValuationReportAfterSavingExcessesAndShortageInStockVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockValuationReportAfterSavingExcessesAndShortageInStockVoucher(), true);
	}
	
	
	
	
	
	
	
	
	// Monthly Weighted Average
	
	@Test(priority=55321)
	public void checkSavingMaterialReceiptNotesVoucher2WithoutUpdateFAWithWarehouseAsHyderabad() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher2WithoutUpdateFAWithWarehouseAsHyderabad(), true);
	}
	
	@Test(priority=55322)
	public void checkSavingMaterialReceiptNotesVoucher3WithoutUpdateFAWithWarehouseAsSecunderabadAndDateSameAsVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher3WithoutUpdateFAWithWarehouseAsSecunderabadAndDateSameAsVoucher2(), true);
	}
	
	@Test(priority=55323)
	public void checkSavingMaterialReceiptNotesVoucher4WithoutUpdateFAWithWarehouseAsHyderabadAndDiffDateAndSameMonthFromVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher4WithoutUpdateFAWithWarehouseAsHyderabadAndDiffDateAndSameMonthFromVoucher2(), true);
	}
	
	@Test(priority=55324)
	public void checkSavingMaterialReceiptNotesVoucher5WithoutUpdateFAWithWarehouseAsHyderabadAndDiffDateAndDiffMonthFromVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher5WithoutUpdateFAWithWarehouseAsHyderabadAndDiffDateAndDiffMonthFromVoucher2(), true);
	}
	
	@Test(priority=55325)
	public void checkLedgerReportAfterSavingMaterialReceiptNotesVouchers2345WithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingMaterialReceiptNotesVouchers2345WithoutUpdateFA(), true);
	}
	
	@Test(priority=55326)
	public void checkStockLedgerReportAfterSavingMaterialReceiptNotesVouchers2345WithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingMaterialReceiptNotesVouchers2345WithoutUpdateFA(), true);
	}
	
	@Test(priority=55327)
	public void checkReSavingMaterialReceiptNotesVoucher2WithUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkReSavingMaterialReceiptNotesVoucher2WithUpdateFA(), true);
	}
	
	@Test(priority=55328)
	public void checkPostingDetailsAfterReSavingMaterialReceiptNotesVoucher2WithUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterReSavingMaterialReceiptNotesVoucher2WithUpdateFA(), true);
	}
	
	@Test(priority=55329)
	public void checkReSavingMaterialReceiptNotesVoucher2WithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkReSavingMaterialReceiptNotesVoucher2WithoutUpdateFA(), true);
	}
	
	@Test(priority=55330)
	public void checkLedgerReportAfterReSavingMaterialReceiptNotesVouchers2WithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterReSavingMaterialReceiptNotesVouchers2WithoutUpdateFA(), true);
	}
	
	@Test(priority=55331)
	public void checkSavingPurchaseVoucher2WithUpdateFAAndWithUpdateStockByLoadingMRNVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher2WithUpdateFAAndWithUpdateStockByLoadingMRNVoucher2(), true);
	}
	
	@Test(priority=55332)
	public void checkPostingDetailsAfterSavingPurchaseVoucher2WithUpdateFAAndUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchaseVoucher2WithUpdateFAAndUpdateStock(), true);
	}
	
	@Test(priority=55333)
	public void checkStockLedgerReportAfterSavingPurchaseVouchers2WithUpdateFAAndUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingPurchaseVouchers2WithUpdateFAAndUpdateStock(), true);
	}
	
	@Test(priority=55334)
	public void checkReSavingPurchaseVoucher2WithOutUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkReSavingPurchaseVoucher2WithOutUpdateStock(), true);
	}
	
	@Test(priority=55335)
	public void checkPostingDetailsAfterReSavingPurchaseVoucher2WithoutUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterReSavingPurchaseVoucher2WithoutUpdateStock(), true);
	}
	
	@Test(priority=55336)
	public void checkStockLedgerReportAfterReSavingPurchaseVouchers2WithoutUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterReSavingPurchaseVouchers2WithoutUpdateStock(), true);
	}
	
	@Test(priority=55337)
	public void checkSavingPurchaseVoucher3ByLoadingMRNVoucher3() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher3ByLoadingMRNVoucher3(), true);
	}
	
	@Test(priority=55338)
	public void checkSavingPurchaseVoucher4ByLoadingMRNVoucher4() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher4ByLoadingMRNVoucher4(), true);
	}
	
	@Test(priority=55339)
	public void checkSavingPurchaseVoucher5ByLoadingMRNVoucher5() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher5ByLoadingMRNVoucher5(), true);
	}
	
	@Test(priority=55340)
	public void checkLedgerReportAfterSavingPurchaseVouchers345() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingPurchaseVouchers345(), true);
	}
	
	@Test(priority=55341)
	public void checkStockLedgerReportAfterSavingPurchaseVouchers345() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingPurchaseVouchers345(), true);
	}
	
	
	
	
	
	
	// Erasing All Transactions
	
	@Test(priority=60000)
	public void checkEraseAllTransactionsLogoutAndLogin() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkEraseAllTransactionsLogoutAndLogin(), true);
	}
	
	
	
	// FIFO COGS Posting
	
	// Purchase Orders
	
	@Test(priority=60001)
	public void checkSavingPurchasesOrderVoucher1WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchasesOrderVoucher1WithFIFOItem(), true);
	}
	
	@Test(priority=60002)
	public void checkSavingPurchasesOrderVoucher2WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchasesOrderVoucher2WithFIFOItem(), true);
	}
	
	@Test(priority=60003)
	public void checkSavingPurchasesOrderVoucher3WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchasesOrderVoucher3WithFIFOItem(), true);
	}
	
	
	
	
	// Material Receipt Notes
	
	@Test(priority=60004)
	public void checkSavingMaterialReceiptNotesVoucher1WithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher1WithoutUpdateFA(), true);
	}
	
	@Test(priority=60005)
	public void checkSavingMaterialReceiptNotesVoucher2WithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher2WithoutUpdateFA(), true);
	}
	
	@Test(priority=60006)
	public void checkSavingMaterialReceiptNotesVoucher3WithoutUpdateFA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher3WithoutUpdateFA(), true);
	}
	
	@Test(priority=60007)
	public void checkLedgerReportAfterSavingMaterialReceiptNotesWithoutUpdateFAWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingMaterialReceiptNotesWithoutUpdateFAWithFIFOItem(), true);
	}
	
	@Test(priority=60008)
	public void checkStockLedgerReportAfterSavingMaterialReceiptNotesWithoutUpdateFAWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingMaterialReceiptNotesWithoutUpdateFAWithFIFOItem(), true);
	}
	
	
	
	
	
	// Purchases Voucher
	
	@Test(priority=60009)
	public void checkSavingPurchaseVoucher1WithoutUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher1WithoutUpdateStock(), true);
	}
	
	@Test(priority=60010)
	public void checkPostingDetailsAfterSavingPurchasesVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesVoucher1(), true);
	}
	
	@Test(priority=60011)
	public void checkSavingPurchaseVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher2(), true);
	}
	
	@Test(priority=60012)
	public void checkPostingDetailsAfterSavingPurchasesVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesVoucher2(), true);
	}
	
	@Test(priority=60013)
	public void checkSavingPurchaseVoucher3() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher3(), true);
	}
	
	@Test(priority=60014)
	public void checkPostingDetailsAfterSavingPurchasesVoucher3() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesVoucher3(), true);
	}
	
	@Test(priority=60015)
	public void checkLedgerReportAfterSavingPurchaseVouchersWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingPurchaseVouchersWithFIFOItem(), true);
	}
	
	@Test(priority=60016)
	public void checkStockLedgerReportAfterSavingPurchaseVouchersWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingPurchaseVouchersWithFIFOItem(), true);
	}
	
	
	
	
	
	// Purchase Returns
	
	@Test(priority=60017)
	public void checkSavingPurchaseReturnsVoucherWithQuantity1ByLoadingFromWorkflow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseReturnsVoucherWithQuantity1ByLoadingFromWorkflow(), true);
	}
	
	@Test(priority=60018)
	public void checkPostingDetailsAfterSavingPurchasesReturnsVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesReturnsVoucher1(), true);
	}
	
	@Test(priority=60019)
	public void checkLedgerReportAfterSavingPurchaseReturnsVoucherWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingPurchaseReturnsVoucherWithFIFOItem(), true);
	}
	
	
	@Test(priority=60020)
	public void checkStockLedgerReportAfterSavingPurchaseReturnsVoucherWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingPurchaseReturnsVoucherWithFIFOItem(), true);
	}
	
	
	
	
	// Delivery Notes
	
	@Test(priority=60021)
	public void checkSavingDeliveryNotesVoucher1WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingDeliveryNotesVoucher1WithFIFOItem(), true);
	}
	
	@Test(priority=60022)
	public void checkPostingDetailsAfterSavingDeliveryNotesVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingDeliveryNotesVoucher1(), true);
	}
	
	@Test(priority=60023)
	public void checkSavingDeliveryNotesVoucher2WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingDeliveryNotesVoucher2WithFIFOItem(), true);
	}
	
	@Test(priority=60024)
	public void checkPostingDetailsAfterSavingDeliveryNotesVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingDeliveryNotesVoucher2(), true);
	}
	
	@Test(priority=60025)
	public void checkSavingDeliveryNotesVoucher3WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingDeliveryNotesVoucher3WithFIFOItem(), true);
	}
	
	@Test(priority=60026)
	public void checkPostingDetailsAfterSavingDeliveryNotesVoucher3() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingDeliveryNotesVoucher3(), true);
	}
	
	@Test(priority=60027)
	public void checkLedgerReportAfterSavingDeliveryNotesVoucherWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingDeliveryNotesVoucherWithFIFOItem(), true);
	}
	
	
	@Test(priority=60028)
	public void checkStockLedgerReportAfterSavingDeliveryNotesVoucherWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingDeliveryNotesVoucherWithFIFOItem(), true);
	}
	
	
	
	
	
	// Sales Invoice
	
	@Test(priority=60029)
	public void checkSavingSalesInvoiceVoucher1WithoutUpdateStock() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesInvoiceVoucher1WithoutUpdateStock(), true);
	}
	
	@Test(priority=60030)
	public void checkPostingDetailsAfterSavingSalesInvoiceVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesInvoiceVoucher1(), true);
	}
	
	@Test(priority=60031)
	public void checkSavingSalesInvoiceVoucher2WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesInvoiceVoucher2WithFIFOItem(), true);
	}
	
	@Test(priority=60032)
	public void checkPostingDetailsAfterSavingSalesInvoiceVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesInvoiceVoucher2(), true);
	}
	
	@Test(priority=60033)
	public void checkSavingSalesInvoiceVoucher3WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesInvoiceVoucher3WithFIFOItem(), true);
	}
	
	@Test(priority=60034)
	public void checkPostingDetailsAfterSavingSalesInvoiceVoucher3() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesInvoiceVoucher3(), true);
	}
	
	
	
	
	
	
	// Sales Returns
	
	@Test(priority=60035)
	public void checkSavingSalesReturnsVoucher1WithUpdateStockAndUpdateFAWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesReturnsVoucher1WithUpdateStockAndUpdateFAWithFIFOItem(), true);
	}
	
	@Test(priority=60036)
	public void checkPostingDetailsAfterSavingSalesReturnsVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesReturnsVoucher1(), true);
	}
	
	@Test(priority=60037)
	public void checkSavingSalesReturnsVoucher2WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesReturnsVoucher2WithFIFOItem(), true);
	}
	
	@Test(priority=60038)
	public void checkPostingDetailsAfterSavingSalesReturnsVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesReturnsVoucher2(), true);
	}
	
	@Test(priority=60039)
	public void checkLedgerReportAfterSavingSalesInvoiceVoucherWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingSalesInvoiceVoucherWithFIFOItem(), true);
	}
	
	@Test(priority=60040)
	public void checkStockLedgerReportAfterSavingSalesInvoiceVoucherWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingSalesInvoiceVoucherWithFIFOItem(), true);
	}
	
	
	
	
	
	// Excesses In Stocks
	
	@Test(priority=60041)
	public void checkSavingExcessesInStocksVoucher1WithUpdateStockAndUpdateFAWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingExcessesInStocksVoucher1WithUpdateStockAndUpdateFAWithFIFOItem(), true);
	}
	
	@Test(priority=60042)
	public void checkPostingDetailsAfterSavingExcessesInStocksVoucher1WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingExcessesInStocksVoucher1WithFIFOItem(), true);
	}
	
	@Test(priority=60043)
	public void checkSavingExcessesInStocksVoucher2WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingExcessesInStocksVoucher2WithFIFOItem(), true);
	}
	
	@Test(priority=60044)
	public void checkPostingDetailsAfterSavingExcessesInStocksVoucher2WithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingExcessesInStocksVoucher2WithFIFOItem(), true);
	}
	
	@Test(priority=60045)
	public void checkLedgerReportAfterSavingExcessesInStocksVouchersWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingExcessesInStocksVouchersWithFIFOItem(), true);
	}
	
	
	@Test(priority=60046)
	public void checkStockLedgerReportAfterSavingExcessesInStocksVouchersWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingExcessesInStocksVouchersWithFIFOItem(), true);
	}
	
	
	
	
	
	
	// Shortage In Stocks
	
	@Test(priority=60047)
	public void checkSavingShortageInStocksVoucherWithFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingShortageInStocksVoucherWithFIFOItem(), true);
	}
	
	@Test(priority=60048)
	public void checkPostingDetailsAfterSavingShortageInStocksVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingShortageInStocksVoucher1(), true);
	}
	
	
	
	// Erasing All Transactions
	
	@Test(priority=60060)
	public void checkEraseAllTransactionsLogoutAndLoginAfterFIFOItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkEraseAllTransactionsLogoutAndLogin(), true);
	}
	
	
	
	
	
	
	
	// Batch Rate COGS Posting
	
	// Purchase Orders
	@Test(priority=60100)
	public void checkSavingPurchasesOrderVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchasesOrderVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60101)
	public void checkSavingPurchasesOrderVoucher2WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchasesOrderVoucher2WithBatchRateItem(), true);
	}
	
	
	
	
	// Material Receipts Note
	
	@Test(priority=60102)
	public void checkSavingMaterialReceiptNotesVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60103)
	public void checkSavingMaterialReceiptNotesVoucher2WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher2WithBatchRateItem(), true);
	}
	
	@Test(priority=60104)
	public void checkLedgerReportAfterSavingMaterialReceiptNotesWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingMaterialReceiptNotesWithBatchRateItem(), true);
	}
	
	@Test(priority=60105)
	public void checkStockLedgerReportAfterSavingMaterialReceiptNotesWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingMaterialReceiptNotesWithBatchRateItem(), true);
	}
	
	
	
	
	
	
	// Purchase Voucher
	
	@Test(priority=60106)
	public void checkSavingPurchaseVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60107)
	public void checkPostingDetailsAfterSavingPurchasesVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60108)
	public void checkSavingPurchaseVoucher2WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher2WithBatchRateItem(), true);
	}
	
	@Test(priority=60109)
	public void checkPostingDetailsAfterSavingPurchasesVoucher2WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesVoucher2WithBatchRateItem(), true);
	}
	
	@Test(priority=60110)
	public void checkLedgerReportAfterSavingPurchaseVouchersWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingPurchaseVouchersWithBatchRateItem(), true);
	}
	
	@Test(priority=60111)
	public void checkStockLedgerReportAfterSavingPurchaseVouchersWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingPurchaseVouchersWithBatchRateItem(), true);
	}
	
	
	
	
	
	
	// Purchase Returns
	
	@Test(priority=60112)
	public void checkSavingPurchaseReturnsVoucherWithQuantity1ByLoadingFromWorkflowWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseReturnsVoucherWithQuantity1ByLoadingFromWorkflowWithBatchRateItem(), true);
	}
	
	@Test(priority=60113)
	public void checkPostingDetailsAfterSavingPurchasesReturnsVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesReturnsVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60114)
	public void checkLedgerReportAfterSavingPurchaseReturnsVoucherWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingPurchaseReturnsVoucherWithBatchRateItem(), true);
	}
	
	@Test(priority=60115)
	public void checkStockLedgerReportAfterSavingPurchaseReturnsVoucherWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingPurchaseReturnsVoucherWithBatchRateItem(), true);
	}
	
	
//Completed	
	
	
	
	// Delivery Note
	
	@Test(priority=60116)
	public void checkSavingDeliveryNotesVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingDeliveryNotesVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60117)
	public void checkPostingDetailsAfterSavingDeliveryNotesVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingDeliveryNotesVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60118)
	public void checkSavingDeliveryNotesVoucher2WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingDeliveryNotesVoucher2WithBatchRateItem(), true);
	}
	
	@Test(priority=60119)
	public void checkPostingDetailsAfterSavingDeliveryNotesVoucher2WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingDeliveryNotesVoucher2WithBatchRateItem(), true);
	}
	
	@Test(priority=60120)
	public void checkLedgerReportAfterSavingDeliveryNotesVoucherWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingDeliveryNotesVoucherWithBatchRateItem(), true);
	}
	
	@Test(priority=60121)
	public void checkStockLedgerReportAfterSavingDeliveryNotesVoucherWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingDeliveryNotesVoucherWithBatchRateItem(), true);
	}
	
	
	
	
	
	
	
	// Sales Invoice 
	
	@Test(priority=60122)
	public void checkSavingSalesInvoiceVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesInvoiceVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60123)
	public void checkPostingDetailsAfterSavingSalesInvoiceVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesInvoiceVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60124)
	public void checkSavingSalesInvoiceVoucher2WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesInvoiceVoucher2WithBatchRateItem(), true);
	}
	
	@Test(priority=60125)
	public void checkPostingDetailsAfterSavingSalesInvoiceVoucher2WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesInvoiceVoucher2WithBatchRateItem(), true);
	}
	
	@Test(priority=60126)
	public void checkLedgerReportAfterSavingSalesInvoiceVouchersWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingSalesInvoiceVouchersWithBatchRateItem(), true);
	}
	
	@Test(priority=60127)
	public void checkStockLedgerReportAfterSavingSalesInvoiceVouchersWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingSalesInvoiceVouchersWithBatchRateItem(), true);
	}
	
	
	
	
	
	
	
	// Sales Returns
	
	@Test(priority=60128)
	public void checkSavingSalesReturnsVoucher1WithUpdateStockAndUpdateFAWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesReturnsVoucher1WithUpdateStockAndUpdateFAWithBatchRateItem(), true);
	}
	
	@Test(priority=60129)
	public void checkPostingDetailsAfterSavingSalesReturnsVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesReturnsVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60130)
	public void checkLedgerReportAfterSavingSalesReturnsVouchersWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingSalesReturnsVouchersWithBatchRateItem(), true);
	}
	
	@Test(priority=60131)
	public void checkStockLedgerReportAfterSavingSalesReturnsVouchersWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingSalesReturnsVouchersWithBatchRateItem(), true);
	}
	
	@Test(priority=60132)
	public void checkResavingSalesReturnsVoucher1Row2WithBatchNumber() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkResavingSalesReturnsVoucher1Row2WithBatchNumber(), true);
	}
	
	@Test(priority=60133)
	public void checkPostingDetailsAfterResavingSalesReturnsVoucher1WithBatchColumn() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterResavingSalesReturnsVoucher1WithBatchColumn(), true);
	}
	
	@Test(priority=60134)
	public void checkLedgerReportAfterResavingSalesReturnsVouchersWithBatchColumn() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterResavingSalesReturnsVouchersWithBatchColumn(), true);
	}
	
	@Test(priority=60135)
	public void checkStockLedgerReportAfterResavingSalesReturnsVouchersWithBatchColumn() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterResavingSalesReturnsVouchersWithBatchColumn(), true);
	}
	
	@Test(priority=60136)
	public void checkSavingSalesReturnsVoucher2WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesReturnsVoucher2WithBatchRateItem(), true);
	}
	
	@Test(priority=60137)
	public void checkPostingDetailsAfterSavingSalesReturnsVoucher2WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesReturnsVoucher2WithBatchRateItem(), true);
	}
	
	@Test(priority=60138)
	public void checkLedgerReportAfterSavingSalesReturnsVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingSalesReturnsVoucher2(), true);
	}
	
	@Test(priority=60139)
	public void checkStockLedgerReportAfterSavingSalesReturnsVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingSalesReturnsVoucher2(), true);
	}
	
	
	
	
	
	
	// Excesses in Stocks
	
	@Test(priority=60140)
	public void checkSavingExcessesInStocksVoucher1WithUpdateStockAndUpdateFAWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingExcessesInStocksVoucher1WithUpdateStockAndUpdateFAWithBatchRateItem(), true);
	}
	
	@Test(priority=60141)
	public void checkPostingDetailsAfterSavingExcessesInStocksVoucher1WithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingExcessesInStocksVoucher1WithBatchRateItem(), true);
	}
	
	@Test(priority=60142)
	public void checkLedgerReportAfterSavingExcessesInStocksVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingExcessesInStocksVoucher(), true);
	}
	
	@Test(priority=60143)
	public void checkStockLedgerReportAfterSavingExcessesInStocksVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingExcessesInStocksVoucher(), true);
	}
	
	
	
	
	
	
	// Shortage In Stocks
	
	@Test(priority=60144)
	public void checkSavingShortageInStocksVoucherWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingShortageInStocksVoucherWithBatchRateItem(), true);
	}
	
	@Test(priority=60145)
	public void checkPostingDetailsAfterSavingShortageInStocksVoucherWithBatchRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingShortageInStocksVoucherWithBatchRateItem(), true);
	}
	
	@Test(priority=60146)
	public void checkLedgerReportAfterSavingShortageInStocksVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingShortageInStocksVoucher(), true);
	}
	
	@Test(priority=60147)
	public void checkStockLedgerReportAfterSavingShorageInStocksVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingShorageInStocksVoucher(), true);
	}
	
	
	// Erasing All Transactions
	
	@Test(priority=60150)
	public void checkEraseAllTransactionsLogoutAndLoginAfterBatchItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkEraseAllTransactionsLogoutAndLogin(), true);
	}
	
	
	
	
	
	
	
	
	// Standard Item COGS Posting
	
	// Purchases Orders
	
	@Test(priority=60151)
	public void checkSavingPurchasesOrderVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchasesOrderVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60152)
	public void checkSavingPurchasesOrderVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchasesOrderVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60153)
	public void checkSavingPurchasesOrderVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchasesOrderVoucher3WithStandardRateItem(), true);
	}
	
	
	
	
	// Material Receipt Note
	
	@Test(priority=60154)
	public void checkSavingMaterialReceiptNotesVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60155)
	public void checkSavingMaterialReceiptNotesVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60156)
	public void checkSavingMaterialReceiptNotesVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingMaterialReceiptNotesVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60157)
	public void checkLedgerReportAfterSavingMaterialReceiptNotesWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingMaterialReceiptNotesWithStandardRateItem(), true);
	}
	
	@Test(priority=60158)
	public void checkStockLedgerReportAfterSavingMaterialReceiptNotesWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingMaterialReceiptNotesWithStandardRateItem(), true);
	}
	
	
	
	
	
	// Purchase Voucher
	
	@Test(priority=60159)
	public void checkSavingPurchaseVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60160)
	public void checkPostingDetailsAfterSavingPurchasesVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60161)
	public void checkSavingPurchaseVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60162)
	public void checkPostingDetailsAfterSavingPurchasesVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60163)
	public void checkSavingPurchaseVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60164)
	public void checkPostingDetailsAfterSavingPurchasesVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60165)
	public void checkLedgerReportAfterSavingPurchaseVouchersWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingPurchaseVouchersWithStandardRateItem(), true);
	}
	
	@Test(priority=60166)
	public void checkStockLedgerReportAfterSavingPurchaseVouchersWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingPurchaseVouchersWithStandardRateItem(), true);
	}
	
	
	
	
	
	
	// Purchase Returns
	
	@Test(priority=60167)
	public void checkSavingPurchaseReturnsVoucher1WithQuantity1ByLoadingFromWorkflowWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseReturnsVoucher1WithQuantity1ByLoadingFromWorkflowWithStandardRateItem(), true);
	}
	
	@Test(priority=60168)
	public void checkPostingDetailsAfterSavingPurchasesReturnsVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesReturnsVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60169)
	public void checkSavingPurchaseReturnsVoucher2WithQuantity1ByLoadingFromWorkflowWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseReturnsVoucher2WithQuantity1ByLoadingFromWorkflowWithStandardRateItem(), true);
	}
	
	@Test(priority=60170)
	public void checkPostingDetailsAfterSavingPurchasesReturnsVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesReturnsVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60171)
	public void checkSavingPurchaseReturnsVoucher3WithQuantity1ByLoadingFromWorkflowWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingPurchaseReturnsVoucher3WithQuantity1ByLoadingFromWorkflowWithStandardRateItem(), true);
	}
	
	@Test(priority=60172)
	public void checkPostingDetailsAfterSavingPurchasesReturnsVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingPurchasesReturnsVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60173)
	public void checkLedgerReportAfterSavingPurchaseReturnsVoucherWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingPurchaseReturnsVoucherWithStandardRateItem(), true);
	}
	
	@Test(priority=60174)
	public void checkStockLedgerReportAfterSavingPurchaseReturnsVoucherWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingPurchaseReturnsVoucherWithStandardRateItem(), true);
	}
	
	
	
	
	
	
	
	// Delivery Notes
	
	@Test(priority=60175)
	public void checkSavingDeliveryNotesVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingDeliveryNotesVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60176)
	public void checkPostingDetailsAfterSavingDeliveryNotesVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingDeliveryNotesVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60177)
	public void checkSavingDeliveryNotesVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingDeliveryNotesVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60178)
	public void checkPostingDetailsAfterSavingDeliveryNotesVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingDeliveryNotesVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60179)
	public void checkSavingDeliveryNotesVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingDeliveryNotesVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60180)
	public void checkPostingDetailsAfterSavingDeliveryNotesVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingDeliveryNotesVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60181)
	public void checkLedgerReportAfterSavingDeliveryNotesVoucherWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingDeliveryNotesVoucherWithStandardRateItem(), true);
	}
	
	@Test(priority=60182)
	public void checkStockLedgerReportAfterSavingDeliveryNotesVoucherWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingDeliveryNotesVoucherWithStandardRateItem(), true);
	}
	
	
	
	
	
	
	// Sales Invoice
	
	@Test(priority=60183)
	public void checkSavingSalesInvoiceVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesInvoiceVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60184)
	public void checkPostingDetailsAfterSavingSalesInvoiceVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesInvoiceVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60185)
	public void checkSavingSalesInvoiceVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesInvoiceVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60186)
	public void checkPostingDetailsAfterSavingSalesInvoiceVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesInvoiceVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60187)
	public void checkSavingSalesInvoiceVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesInvoiceVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60188)
	public void checkPostingDetailsAfterSavingSalesInvoiceVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesInvoiceVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60189)
	public void checkLedgerReportAfterSavingSalesInvoiceVouchersWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingSalesInvoiceVouchersWithStandardRateItem(), true);
	}
	
	@Test(priority=60190)
	public void checkStockLedgerReportAfterSavingSalesInvoiceVouchersWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingSalesInvoiceVouchersWithStandardRateItem(), true);
	}
	
	
	
	
	
	
	// Sales Returns
	
	@Test(priority=60191)
	public void checkSavingSalesReturnsVoucher1WithUpdateStockAndUpdateFAWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesReturnsVoucher1WithUpdateStockAndUpdateFAWithStandardRateItem(), true);
	}
	
	@Test(priority=60192)
	public void checkPostingDetailsAfterSavingSalesReturnsVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesReturnsVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60193)
	public void checkSavingSalesReturnsVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesReturnsVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60194)
	public void checkPostingDetailsAfterSavingSalesReturnsVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesReturnsVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60195)
	public void checkSavingSalesReturnsVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesReturnsVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60196)
	public void checkPostingDetailsAfterSavingSalesReturnsVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesReturnsVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60197)
	public void checkSavingSalesReturnsVoucher4WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingSalesReturnsVoucher4WithStandardRateItem(), true);
	}
	
	@Test(priority=60198)
	public void checkPostingDetailsAfterSavingSalesReturnsVoucher4WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingSalesReturnsVoucher4WithStandardRateItem(), true);
	}
	
	@Test(priority=60199)
	public void checkLedgerReportAfterSavingSalesReturnsVouchersWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingSalesReturnsVouchersWithStandardRateItem(), true);
	}
	
	@Test(priority=60200)
	public void checkStockLedgerReportAfterSavingSalesReturnsVouchersWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingSalesReturnsVouchersWithStandardRateItem(), true);
	}
	
	
	
	
	
	
	// Excesses in Stocks
	
	@Test(priority=60201)
	public void checkSavingExcessesInStocksVoucher1WithUpdateStockAndUpdateFAWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingExcessesInStocksVoucher1WithUpdateStockAndUpdateFAWithStandardRateItem(), true);
	}
	
	@Test(priority=60202)
	public void checkPostingDetailsAfterSavingExcessesInStocksVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingExcessesInStocksVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60203) // Issue
	public void checkLedgerReportAfterSavingExcessesInStocksVouchersWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingExcessesInStocksVouchersWithStandardRateItem(), true);
	}
	
	@Test(priority=60204)
	public void checkStockLedgerReportAfterSavingExcessesInStocksVouchersWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingExcessesInStocksVouchersWithStandardRateItem(), true);
	}
	
	
	
	
	
	
	// Shortage in Stocks
	
	@Test(priority=60205)
	public void checkSavingShortageInStocksVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingShortageInStocksVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60206)
	public void checkPostingDetailsAfterSavingShortageInStocksVoucher1WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingShortageInStocksVoucher1WithStandardRateItem(), true);
	}
	
	@Test(priority=60207)
	public void checkSavingShortageInStocksVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingShortageInStocksVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60208)
	public void checkPostingDetailsAfterSavingShortageInStocksVoucher2WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingShortageInStocksVoucher2WithStandardRateItem(), true);
	}
	
	@Test(priority=60209)
	public void checkSavingShortageInStocksVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingShortageInStocksVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60210)
	public void checkPostingDetailsAfterSavingShortageInStocksVoucher3WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingShortageInStocksVoucher3WithStandardRateItem(), true);
	}
	
	@Test(priority=60211)
	public void checkSavingShortageInStocksVoucher4WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingShortageInStocksVoucher4WithStandardRateItem(), true);
	}
	
	@Test(priority=60212)
	public void checkPostingDetailsAfterSavingShortageInStocksVoucher4WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingShortageInStocksVoucher4WithStandardRateItem(), true);
	}
	
	@Test(priority=60213)
	public void checkSavingShortageInStocksVoucher5WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkSavingShortageInStocksVoucher5WithStandardRateItem(), true);
	}
	
	@Test(priority=60214)
	public void checkPostingDetailsAfterSavingShortageInStocksVoucher5WithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkPostingDetailsAfterSavingShortageInStocksVoucher5WithStandardRateItem(), true);
	}
	
	@Test(priority=60215) // Issue 
	public void checkLedgerReportAfterSavingShortageInStocksVouchersWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkLedgerReportAfterSavingShortageInStocksVouchersWithStandardRateItem(), true);
	}
	
	@Test(priority=60216)// Data in Default browser and Automation browser are not matching
	public void checkStockLedgerReportAfterSavingShortageInStocksVouchersWithStandardRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		CPP=new COGSPostingPage(getDriver());
		Assert.assertEquals(CPP.checkStockLedgerReportAfterSavingShortageInStocksVouchersWithStandardRateItem(), true);
	}
	
	
	//Cogs Posting Enable and Disable Scnerios
	
	
	
	
	
	
	
	
	
	
	
	
	
	
}
