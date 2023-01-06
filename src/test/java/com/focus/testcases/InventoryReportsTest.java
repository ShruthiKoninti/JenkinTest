package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.InventoryReportsPage;
import com.focus.base.BaseEngine;

public class InventoryReportsTest extends BaseEngine
{
	InventoryReportsPage IRP;
	
	
	 @Test(priority=341)
	 public void checkItemQueryReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkItemQueryReport(), true);
	 }
	 
	 
	 @Test(priority=342)
	 public void checkStockLedgerReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockLedgerReport(), true);
	 }
	 
	 @Test(priority=343)
	 public void checkStockLedgerCustomizationReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockLedgerCustomizationReport(), true);
	 }
	 
	 @Test(priority=344)
	 public void checkStockLedgerReportWithMovedItemsOnlyAndIncludeServiceTypeItemsCheckBoxs() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockLedgerReportWithMovedItemsOnlyAndIncludeServiceTypeItemsCheckBoxs(), true);
	 }
	 
	 @Test(priority=345)
	 public void checkStockLedgerReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockLedgerReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	
	 
	 @Test(priority=346)
	 public void checkOpeningStockRegisterReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkOpeningStockRegisterReport(), true);
	 }
	 
	 @Test(priority=347)
	 public void checkOpeningStockRegisterReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkOpeningStockRegisterReportFilterOption(), true);
	 } 
	
	 @Test(priority=348)
	 public void checkOpeningStockRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkOpeningStockRegisterReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	 
	 
	 @Test(priority=349)
	 public void checkStockStatementReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockStatementReport(), true);
	 } 
	 
	 @Test(priority=350)
	 public void checkStockStatementReportPrintOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockStatementReportPrintOption(), true);
	 } 
	 
	 @Test(priority=351)
	 public void checkStockStatementReportWithIncludeServiceTypeItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockStatementReportWithIncludeServiceTypeItem(), true);
	 } 
	 
	 
	 
	 
	 
	
	 
	 @Test(priority=352)
	 public void checkStockMovementReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockMovementReports(), true);
	 } 
	 
	 @Test(priority=353)
	 public void checkStockMovementReportsWithStockMovementByWarehouseAndIncludeServiceTypeItemOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockMovementReportsWithStockMovementByWarehouseAndIncludeServiceTypeItemOptions(), true);
	 } 
	 
	 @Test(priority=354)
	 public void checkStockMovementReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockMovementReportWithFilterOption(), true);
	 } 
	 
	 @Test(priority=355)
	 public void checkStockMovementReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockMovementReportBackTrackOption(), true);
	 } 
	 
	 
	
	
	 
	 
	 @Test(priority=356)
	 public void checkMultiLevelStockMovementReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkMultiLevelStockMovementReports(), true);
	 } 
	 
	 @Test(priority=357) // Filter option is not working as expected
	 public void checkMultiLevelStockMovementReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkMultiLevelStockMovementReportFilterOption(), true);
	 } 
	 
	 @Test(priority=358)
	 public void checkMultiLevelStockMovementReportWithIncludeServiceTypeItemOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkMultiLevelStockMovementReportWithIncludeServiceTypeItemOption(), true);
	 } 
	     
	 @Test(priority=359)
	 public void checkMultiLevelStockMovementReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkMultiLevelStockMovementReportBackTrackOption(), true);
	 } 
	 
	 
	 
	 
	
	 @Test(priority=360)
	 public void checkVirtualStockAnalysisReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkVirtualStockAnalysisReports(), true);
	 } 
	 
	 @Test(priority=361)
	 public void checkVirtualStockAnalysisReportWithIncludeServiceTypeItemOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkVirtualStockAnalysisReportWithIncludeServiceTypeItemOption(), true);
	 } 
	 
	 @Test(priority=362)
	 public void checkVirtualStockAnalysisReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkVirtualStockAnalysisReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	 
	 
	 @Test(priority=363)
	 public void checkStockValuationReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockValuationReports(), true);
	 }
	 
	 @Test(priority=364)
	 public void checkStockValuationCustomizationReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockValuationCustomizationReport(), true);
	 }
	
	 @Test(priority=365)
	 public void checkStockValuationReportWithStockValuationDisplayLinearReportIncludeServiceTypeItemOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockValuationReportWithStockValuationDisplayLinearReportIncludeServiceTypeItemOption(), true);
	 }
	 
	 @Test(priority=366)
	 public void checkStockValuationReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockValuationReportBackTrackOption(), true);
	 }
	 
	 
	 
	
	 
	
	 @Test(priority=367)
	 public void checkStockReportbyTagReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockReportbyTagReports(), true);
	 }
	 
	 @Test(priority=368)
	 public void checkStockReportbyTagReportPrintOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockReportbyTagReportPrintOption(), true);
	 }
	 
	 @Test(priority=369)
	 public void checkStockReportbyTagReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockReportbyTagReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
	 @Test(priority=370)
	 public void checkABCAnalysisReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkABCAnalysisReports(), true);
	 }
	 
	 @Test(priority=371)
	 public void checkABCAnalysisReportWithClassificationAndCriteriaOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkABCAnalysisReportWithClassificationAndCriteriaOptions(), true);
	 }
	 
	 
	 
	 
	
	 
	 @Test(priority=372)
	 public void checkStockAnalysisByBatchReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAnalysisByBatchReport(), true);
	 }
	 
	 @Test(priority=373)
	 public void checkStockAnalysisByBatchReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAnalysisByBatchReportFilterOption(), true);
	 }
	 
	 @Test(priority=374)
	 public void checkStockAnalysisByBatchReportWithStockTypeGoodBatches() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAnalysisByBatchReportWithStockTypeGoodBatches(), true);
	 }
	 
	 @Test(priority=375)
	 public void checkStockAnalysisByBatchReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAnalysisByBatchReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
	 @Test(priority=376)
	 public void checkStockAnalysisByRmaReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAnalysisByRmaReports(), true);
	 }
	 
	 @Test(priority=377)
	 public void checkStockAnalysisByRmaReportsPrintOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAnalysisByRmaReportsPrintOption(), true);
	 }
	 
	 
	 
	 
	 
	 @Test(priority=378)
	 public void checkStockDetailsByBinsReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockDetailsByBinsReports(), true);
	 }
	 
	 @Test(priority=379)
	 public void checkStockDetailsByBinsReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockDetailsByBinsReportFilterOption(), true);
	 }
	 
	 @Test(priority=380)
	 public void checkStockDetailsByBinsReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockDetailsByBinsReportBackTrackOption(), true);
	 }
	 
	
	 
	 
	 
	 
	 @Test(priority=381)
	 public void checkStockBalanaceByBinsReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockBalanaceByBinsReports(), true);
	 }
	 
	 @Test(priority=382)
	 public void checkStockBalanaceByBinsReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockBalanaceByBinsReportFilterOption(), true);
	 }
	 
	 
	 @Test(priority=383)
	 public void checkStockBalanaceByBinsReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockBalanaceByBinsReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	 
	 
	 @Test(priority=384)
	 public void checkStockBalancesByItemByBinsReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockBalancesByItemByBinsReports(), true);
	 }
	 
	 @Test(priority=385)
	 public void checkStockBalancesByItemByBinsReportPrintOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockBalancesByItemByBinsReportPrintOption(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
	 @Test(priority=386)
	 public void checkExpiredStockByBinsReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkExpiredStockByBinsReports(), true);
	 }
	 
	 @Test(priority=387)
	 public void checkExpiredStockByBinsReportCustomizationOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkExpiredStockByBinsReportCustomizationOption(), true);
	 }
	 
	 
	 
	 
	 
	 
	 @Test(priority=388)
	 public void checkStockBalanceByWarehouseReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockBalanceByWarehouseReports(), true);
	 }
	 
	 @Test(priority=389)
	 public void checkStockBalanceByWarehouseReportIncludeServiceTypeItemOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockBalanceByWarehouseReportIncludeServiceTypeItemOption(), true);
	 }
	 
	 @Test(priority=390)
	 public void checkStockBalanceByWarehouseReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockBalanceByWarehouseReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	 
	
	 @Test(priority=391)
	 public void checkStockAgeingAnalysisReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAgeingAnalysisReports(), true);
	 }
	 
	 @Test(priority=392)
	 public void checkStockAgeingAnalysisReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAgeingAnalysisReportFilterOption(), true);
	 }
	 
	 @Test(priority=393)
	 public void checkStockAgeingAnalysisReportWithOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAgeingAnalysisReportWithOptions(), true);
	 }
	 
	 @Test(priority=394)
	 public void checkStockAgeingAnalysisReportWithBackTrackOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAgeingAnalysisReportWithBackTrackOptions(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
	 
	 @Test(priority=395)
	 public void checkAgeingAnalysisByBatchReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingAnalysisByBatchReports(), true);
	 }
	 
	 @Test(priority=396)
	 public void checkStockAnalysisByBatchReportPrintOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAnalysisByBatchReportPrintOption(), true);
	 }
	 
	 @Test(priority=397)
	 public void checkAgeingAnalysisByBatchReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingAnalysisByBatchReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions(), true);
	 }
	 
	 @Test(priority=398)
	 public void checkAgeingAnalysisByBatchReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingAnalysisByBatchReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
	 @Test(priority=399)
	 public void checkAgeingAnalysisByRMAReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingAnalysisByRMAReport(), true);
	 }
	 
	 @Test(priority=400)
	 public void checkAgeingAnalysisByRMAReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingAnalysisByRMAReportWithFilterOption(), true);
	 }
	 
	 @Test(priority=401)
	 public void checkAgeingAnalysisByRMAReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingAnalysisByRMAReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions(), true);
	 }
	 
	 @Test(priority=402)
	 public void checkAgeingAnalysisByRMAReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingAnalysisByRMAReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
	 @Test(priority=403)
	 public void checkAgeingStockByBinsReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingStockByBinsReports(), true);
	 }
	 
	 @Test(priority=404)
	 public void checkStockAnalysisByBinReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockAnalysisByBinReportWithFilterOption(), true);
	 }
	 
	 @Test(priority=405)
	 public void checkAgeingAnalysisByBinReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingAnalysisByBinReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions(), true);
	 }
	 
	 @Test(priority=406)
	 public void checkAgeingAnalysisByBinReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingAnalysisByBinReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	
	 @Test(priority=407)
	 public void checkAgeingByItemByBinsReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingByItemByBinsReports(), true);
	 }
	 
	 @Test(priority=408)
	 public void checkStockByItemByBinReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockByItemByBinReportWithFilterOption(), true);
	 }
	 
	 @Test(priority=409)
	 public void checkAgeingByItemByBinReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingByItemByBinReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions(), true);
	 }
	 
	 @Test(priority=410)
	 public void checkAgeingByItemByBinReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkAgeingByItemByBinReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	
	 
	 @Test(priority=411)
	 public void checkReorderReports() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkReorderReports(), true);
	 }
	 
	 @Test(priority=412)
	 public void checkReorderReportPrintOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkReorderReportPrintOption(), true);
	 }

	 
	 
	 
	 
	 @Test(priority=413)
	 public void checkFastMovingItemsReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkFastMovingItemsReport(), true);
	 }
	 
	 @Test(priority=414)
	 public void checkFastMovingItemsReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkFastMovingItemsReportWithFilterOption(), true);
	 }
	 
	 @Test(priority=415)
	 public void checkFastMovingItemsReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkFastMovingItemsReportBackTrackOption(), true);
	 }
	 
	
	 
	 
	 
	 
	 @Test(priority=416)
	 public void checkSlowMovingItemsReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkSlowMovingItemsReport(), true);
	 }
	 
	 @Test(priority=417)
	 public void checkSlowMovingItemsReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkSlowMovingItemsReportWithFilterOption(), true);
	 }
	 
	 @Test(priority=418)
	 public void checkSlowMovingItemsReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkSlowMovingItemsReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
	 @Test(priority=419)
	 public void checkPeakAndLowBalancesReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkPeakAndLowBalancesReport(), true);
	 }
	 
	 @Test(priority=420) // Filter is not working
	 public void checkPeakAndLowBalancesReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkPeakAndLowBalancesReportWithFilterOption(), true);
	 }
	 
	 @Test(priority=421) 
	 public void checkPeakAndLowBalancesReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkPeakAndLowBalancesReportBackTrackOption(), true);
	 }
	 
	 
	 
	
	 
	 
	 
	 @Test(priority=422)
	 public void checkBestSellingItemReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkBestSellingItemReport(), true);
	 }
	 
	 @Test(priority=423)
	 public void checkBestSellingItemReportWithSortingOnAndIncludeAllVouchersChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkBestSellingItemReportWithSortingOnAndIncludeAllVouchersChkBox(), true);
	 }
	 
	 @Test(priority=424)
	 public void checkBestSellingItemReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkBestSellingItemReportBackTrackOption(), true);
	 }
	 
	 
	 
	 
	
	 
	 @Test(priority=425)
	 public void checkStockTransferRegisterReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockTransferRegisterReport(), true);
	 }
	 
	 @Test(priority=426)
	 public void checkStockTransferRegisterReportPrintOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockTransferRegisterReportPrintOption(), true);
	 }
	 
	 @Test(priority=427)
	 public void checkStockTransferRegisterReportWithFromAndToWarehouseAndDetailsOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockTransferRegisterReportWithFromAndToWarehouseAndDetailsOptions(), true);
	 }
	 
	 @Test(priority=428)
	 public void checkStockTransferRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockTransferRegisterReportBackTrackOption(), true);
	 }
	 
	 
	 
	
	 @Test(priority=429)
	 public void checkStockReservationRegisterReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockReservationRegisterReport(), true);
	 }
	 
	 @Test(priority=430)
	 public void checkStockReservationRegisterReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockReservationRegisterReportWithFilterOption(), true);
	 }
	 
	 @Test(priority=431)
	 public void checkStockReservationRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockReservationRegisterReportBackTrackOption(), true);
	 }
	 
	
	 
	 
	
	 @Test(priority=432)
	 public void checkTransactionTypeWiseStockReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkTransactionTypeWiseStockReport(), true);
	 }
	 
	 @Test(priority=433)
	 public void checkTransactionTypeWiseStockReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkTransactionTypeWiseStockReportWithFilterOption(), true);
	 }
	 
	  @Test(priority=434)
	 public void checkTransactionTypeWiseStockReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkTransactionTypeWiseStockReportBackTrackOption(), true);
	 }
	 
	 
	 
	  
	  
	 @Test(priority=435)
	 public void checkActualConsumptionReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkActualConsumptionReport(), true);
	 }
	 
	 @Test(priority=436)
	 public void checkActualConsumptionReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkActualConsumptionReportWithFilterOption(), true);
	 }
	 
	 @Test(priority=437)
	 public void checkActualConsumptionReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkActualConsumptionReportBackTrackOption(), true);
	 }
	
	 
		 
	 
	 @Test(priority=438)
	 public void checkLoginWithUserForItemRestriction() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkLoginWithUserForItemRestriction(), true);
	 }
	 
	 @Test(priority=439) 
	 public void checkStockReportbyTagReportsWithUserForItemRestriction() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkStockReportbyTagReportsWithUserForItemRestriction(), true);
	 }
	 
	 @Test(priority=440)
	 public void checkLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 IRP=new InventoryReportsPage(getDriver());
		 Assert.assertEquals(IRP.checkLoginWithSU(), true);
	 }
	
	 
	 
	 
	 
	
	
	 
	 
	 
	 
	 
	 
	 
	 
	 
	 
}
