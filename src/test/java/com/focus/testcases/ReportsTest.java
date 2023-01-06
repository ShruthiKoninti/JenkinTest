package com.focus.testcases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.ReportsPage;
import com.focus.base.BaseEngine;

public class ReportsTest extends BaseEngine 
{
	ReportsPage RP;
	
	
	@Test(priority=1230)
	private void checkFinanacialLedgerReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkFinanacialLedgerReport(), true);
	}
	
	@Test(priority=1231)
	private void checkSubledgerReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkSubledgerReport(), true);
	}
	
	@Test(priority=1232)
	private void checkPdcReceiptsRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkPdcReceiptsRegisterReport(), true);
	}
	
	@Test(priority=1233)
	private void checkPdcPaymentsRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkPdcPaymentsRegisterReport(), true);
	}
	
	@Test(priority=1234)
	private void checkReceiptRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkReceiptRegisterReport(), true);
	}
	
	@Test(priority=1235)
	private void checkCustomerListingOfOutstandingBillsReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkCustomerListingOfOutstandingBillsReport(), true);
	}
	
	@Test(priority=1236)
	private void checkCustomerStatementReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkCustomerStatementReport(), true);
	}
	
	@Test(priority=1237)
	private void checkVendorAgeingSummaryReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkVendorAgeingSummaryReport(), true);
	}
	
	@Test(priority=1238)
	private void checkVendorBillWiseSummaryReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkVendorBillWiseSummaryReport(), true);
	}
	
	@Test(priority=1239)
	private void checkTrailBalanceReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkTrailBalanceReport(), true);
	}
	
	@Test(priority=1240)
	private void checkProfitAndLossReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkProfitAndLossReport(), true);
	}
	
	@Test(priority=1241)
	private void checkBalanceSheetReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkBalanceSheetReport(), true);
	}
	
	
	
	
	
	
	
	
	
	
	
	// Inventory Reports
	
	@Test(priority=1250)
	private void checkStockLedgerReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkStockLedgerReport(), true);
	}
	
	@Test(priority=1251)
	private void checkStockStatementReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkStockStatementReport(), true);
	}
	
	@Test(priority=1252)
	private void checkStockMovementReports() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkStockMovementReports(), true);
	}
	
	@Test(priority=1253)
	private void checkMultiLevelStockMovementReports() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkMultiLevelStockMovementReports(), true);
	}
	
	@Test(priority=1254)
	private void checkStockValuationReports() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkStockValuationReports(), true);
	}
	
	@Test(priority=1255)
	private void checkStockAnalysisByBatchReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkStockAnalysisByBatchReport(), true);
	}
	
	@Test(priority=1255)
	private void checkStockAnalysisByRmaReports() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkStockAnalysisByRmaReports(), true);
	}
	
	@Test(priority=1256)
	private void checkStockDetailsByBinsReports() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkStockDetailsByBinsReports(), true);
	}
	
	@Test(priority=1257)
	private void checkListingOfDocumentsReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkListingOfDocumentsReport(), true);
	}
	
	
	
	// Report Designer
	
	@Test(priority=1265)
	private void checkSavingReportDesignerOfAllTransactionsOfDocumentClass() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkSavingReportDesignerOfAllTransactionsOfDocumentClass(), true);
	}
	
	@Test(priority=1266)
	private void checkAllTransactionsOfDocumentClassReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		RP=new ReportsPage(getDriver());
		Assert.assertEquals(RP.checkAllTransactionsOfDocumentClassReport(), true);
	}
	
	
	
	
	
	

}
