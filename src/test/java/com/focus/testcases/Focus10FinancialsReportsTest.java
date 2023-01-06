package com.focus.testcases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.Focus10FinancialsReportsPage;

/*import com.focus.Pages.Focus10InventoryReportsPage;*/
import com.focus.Pages.ReportsPage;
import com.focus.base.BaseEngine;

public class Focus10FinancialsReportsTest extends BaseEngine{
	
	 Focus10FinancialsReportsPage FRP;
	

	 @Test(priority=1500)
		
		public void checkLoginTOAutomationCompany() throws InterruptedException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkLoginTOAutomationCompany(),true);
		}
	

	 //Ledger
	
	 @Test(priority=1501)
		private void checkFinancialLedgerReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkFinancialLedgerReport(), true);
		}
	 
	 @Test(priority=1502)
		private void checkFinancialLedgerReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkFinancialLedgerReportFilterOption(), true);
		}
	 
	  @Test(priority=1503)
		private void checkFinancialLedgerReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkFinancialLedgerReportBackTrackOption(), true);
		}
	
	
	//Ledger Details
	 
	 
	 @Test(priority=1504)
			private void checkFinancialLedgerDetailsReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkFinancialLedgerDetailsReport(), true);
			}
		 		 
		 @Test(priority=1505)
			private void checkFinancialLedgerDetailsReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkFinancialLedgerDetailsReportFilterOption(), true);
			}
				  
		 @Test(priority=1506)
			private void checkFinancialLedgerDetailsReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkFinancialLedgerDetailsReportBackTrackOption(), true);
			}
	
	
	 
	//SubLedger 
	 
		 @Test(priority=1507)
			private void checkFinancialSubLedgerReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkFinancialSubLedgerReport(), true);
			}
		 
		 
		 @Test(priority=1508)
			private void checkFinancialSubLedgerReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkFinancialSubLedgerReportFilterOption(), true);
			}
		 
		 
		 @Test(priority=1509)
			private void checkFinancialSubLedgerReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkFinancialSubLedgerReportBackTrackOption(), true);
			}
	 
		 @Test(priority=1510)
			private void checkFinancialSubLedgerReportwithSortingOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkFinancialSubLedgerReportwithSortingOption(), true);
			}
	
	 
	//Cash And Bank Book 
	 
		 @Test(priority=1511)
			private void checkCashBookReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkCashBookReport(), true);
			}
	 
		 @Test(priority=1512)
			private void checkCashBookReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkCashBookReportFilterOption(), true);
			}
		 @Test(priority=1513)
			private void checkCashBookReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkCashBookReportBackTrackOption(), true);
			}
	 
		 	 @Test(priority=1514)
			private void  checkCashBookReportwithShowconsolidatedamountsOptionandBalancingType() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP. checkCashBookReportwithShowconsolidatedamountsOptionandBalancingType(), true);
			}
		 	
		 	 @Test(priority=1515)
				private void checkCashBookReportwithAscendingOrderColumn() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
				{
					FRP=new Focus10FinancialsReportsPage(getDriver());
					
					Assert.assertEquals(FRP.checkCashBookReportwithAscendingOrderColumn(),true);

				}

	 
	 
	
	 
	 @Test(priority=1516)
		private void checkBankBookReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkBankBookReport(), true);
		}

	 @Test(priority=1517)
		private void checkBankBookReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkBankBookReportFilterOption(), true);
		}
	 @Test(priority=1518)
		private void checkBankBookReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkBankBookReportBackTrackOption(), true);
		}
	
	 
	 @Test(priority=1519)
		private void  checkBankBookReportwithShowconsolidatedamountsOptionandBalancingType() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP. checkBankBookReportwithShowconsolidatedamountsOptionandBalancingType(), true);
		}

	 
	 
	 @Test(priority=1520)
		private void checkPettyCashBookReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPettyCashBookReport(), true);
		}

	 @Test(priority=1521)
		private void checkPettyCashBookReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPettyCashBookReportFilterOption(), true);
		}
	 @Test(priority=1522)
		private void checkPettyCashBookReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPettyCashBookReportBackTrackOption(), true);
		}
	
	 
	 @Test(priority=1523)
		private void  checkPettyCashBookReportwithShowconsolidatedamountsOptionandBalancingType() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP. checkPettyCashBookReportwithShowconsolidatedamountsOptionandBalancingType(), true);
		}
		
	
	 
	 @Test(priority=1524)
		private void checkDayBookReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkDayBookReport(), true);
		}

	 @Test(priority=1525)
		private void checkDayBookReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkDayBookReportFilterOption(), true);
		}
	 @Test(priority=1526)
		private void checkDayBookReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkDayBookReportBackTrackOption(), true);
		}
	
	 
	 @Test(priority=1527)
		private void  checkDayBookReportwithClubCashSalesfortheDay() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP. checkDayBookReportwithClubCashSalesfortheDay(), true);
		}
	 
	
	 
	 @Test(priority=1528)
		private void checkVirtualBankLedgerBookReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVirtualBankLedgerBookReport(), true);
		}

	 @Test(priority=1529)
		private void checkVirtualBankLedgerBookReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVirtualBankLedgerBookReportFilterOption(), true);
		}
	
	 @Test(priority=1530)
		private void  checkVirtualBankLedgerPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVirtualBankLedgerPrintOption(), true);
		}
	 
	 @Test(priority=1531)
		private void checkVirtualBankLedgerBookReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVirtualBankLedgerBookReportBackTrackOption(), true);
		}
	
	 @Test(priority=1532)
			private void checkBankReconciliationReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkBankReconciliationReport(),true);
			}
		 
		 @Test(priority=1533)
			private void checkCustomerReconciliationReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkCustomerVendorReconciliationReport(),true);
			}
		// @Test(priority=1534)//
			private void checkCustomerReconcilitionReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				FRP.checkCustomerReconcilitionReportBackTrackOption();
			}
		 
		 
		 @Test(priority=1535)
			private void checkBankReconcilationStatementReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkBankReconcilationStatementReport(),true);
			}
		 
		 
		 @Test(priority=1536)
			private void checkBankReconcilationStatementReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkBankReconcilationStatementReportWithFilterOption(),true);
			}
		 
		 @Test(priority=1537)
			private void checkBankReconciliationStatementReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkBankReconciliationStatementReportBackTrackOption(),true);
			}
		 
		 
		 @Test(priority=1538)
			private void checkBankReconcilationStatementReportwithIncludeOptionsCheckbox() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkBankReconcilationStatementReportwithIncludeOptionsCheckbox(),true);
			}
	
	

	
	
	 
	 
	
	 //Sales Report
	 
	 @Test(priority=1540)
		private void checkSalesRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesRegisterReports(), true);
		}

	
	
	 @Test(priority=1541)
		private void checkSalesRegisterReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesRegisterReportWithFilterOption(), true);
		}
	 
	 @Test(priority=1542)
		private void  checkSalesRegisterPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesRegisterPrintOption(), true);
		}
	
		 
	 @Test(priority=1543)
		private void checkSalesRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesRegisterReportBackTrackOption(), true);
		}
	 
	@Test(priority=1544)
		private void  checkSalesRegisterReportswithDisplayMonthlyandSalesReturnVocherOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesRegisterReportswithDisplayMonthlyandSalesReturnVocherOption(), true);
		}
	 
	 
	 
	 @Test(priority=1545)
		private void checkSalesReturnRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesReturnsRegisterReports(), true);
		}

	
	
	 @Test(priority=1546)
		private void checkSalesReturnRegisterReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesReturnRegisterReportWithFilterOption(), true);
		}
	 
	 @Test(priority=1547)
		private void  checkSalesReturnRegisterPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesReturnRegisterPrintOption(), true);
		}
	
		 
	 @Test(priority=1548)
		private void checkSalesReturnRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesReturnRegisterReportBackTrackOption(), true);
		}
	 
	
	 
	 
	 @Test(priority=1549)
		private void checkSummarySalesBookReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSummarySalesBookReports(), true);
		}

	
	
	 @Test(priority=1550)
		private void checkSummarySalesBookReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSummarySalesBookReportWithFilterOption(), true);
		}
	 
	
	
		 
	 @Test(priority=1551)
		private void checkSummarySalesBookReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSummarySalesBookReportBackTrackOption(), true);
		}
	 
	 @Test(priority=1552)
		private void 	 checkSummarySalesBookReportswithDisplayMonthlyandSalesReturnVocherOption
() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSummarySalesBookReportswithDisplayMonthlyandSalesReturnVocherOption
(), true);
		}
	 
	 
	 
	 @Test(priority=1553)
		private void checkMonthlySalesBookReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkMonthlySalesBookReports(), true);
		}

	
	 @Test(priority=1554)
		private void  checkMonthlySalesBookPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkMonthlySalesBookPrintOption(), true);
		}
	
	 @Test(priority=1555)
		private void checkMonthlySalesBookReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkMonthlySalesBookReportWithFilterOption(), true);
		}
	 
	
	 
	 
	 @Test(priority=1556)
		private void checkTopCustomersListReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkTopCustomersListReports(), true);
		}

	
	 @Test(priority=1557)
		private void checkTopCustomersListReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkTopCustomersListReportWithFilterOption(), true);
		}
	 
	 @Test(priority=1558)
		private void  checTopCustomersListReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checTopCustomersListReportBackTrackOption(), true);
		}
	
	 
	//Purchase Report 
	 
	 @Test(priority=1559)
		private void checkPurchaseRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseRegisterReports(), true);
		}

	
	
	 @Test(priority=1560)
		private void checkPurchaseRegisterReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseRegisterReportWithFilterOption(), true);
		}
	 
	 @Test(priority=1561)
		private void  checkPurchaseRegisterPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseRegisterPrintOption(), true);
		}
	
		 
	// @Test(priority=1562)
		private void checkPurchaseRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseRegisterReportBackTrackOption(), true);
		}
	 
	@Test(priority=1563)
		private void  checkPurchaseRegisterReportswithDisplayMonthlyandSalesReturnVocherOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseRegisterReportswithDisplayMonthlyandSalesReturnVocherOption(), true);
		}
	 

	 
	 
	 
	 @Test(priority=1564)//no data
		private void checkPurchaseReturnRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseReturnRegisterReports(), true);
		}

	
	
	 @Test(priority=1565)
		private void checkPurchaseReturnRegisterReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseReturnRegisterReportWithFilterOption(), true);
		}
	 
	 @Test(priority=1566)
		private void  checkPurchaseReturnRegisterPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseReturnRegisterPrintOption(), true);
		}
	
		 
	 @Test(priority=1567)
		private void checkPurchaseReturnRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseReturnRegisterReportBackTrackOption(), true);
		}
	 
	 
	 
	 
	 
	 @Test(priority=1568)
		private void checkSummaryPurchaseBookReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSummaryPurchaseBookReports(), true);
		}

	
	
	 @Test(priority=1569)
		private void checkSummaryPurchaseBookReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSummaryPurchaseBookReportWithFilterOption(), true);
		}
	 
		
		 
	 @Test(priority=1570)
		private void checkSummaryPurchaseBookReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSummaryPurchaseBookReportBackTrackOption(), true);
		}
	 
	@Test(priority=1571)
		private void  checkSummaryPurchaseBookReportswithDisplayMonthlyandSalesReturnVocherOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSummaryPurchaseBookReportswithDisplayMonthlyandSalesReturnVocherOption(), true);
		}
	
	
	
	
	 
	//VAT 
	 
	 @Test(priority=1572)
		private void checkVATPurchaseAccountReports() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATPurchaseAccountReports(), true);
		}

	
	
	 @Test(priority=1573)
		private void checkPurchaseAccountReportFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseAccountReportWithFilterOption(), true);
		}
	 
				 
	 @Test(priority=1574)
		private void checkPurchaseAccountReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseAccountReportBackTrackOption(), true);
		}
	 
	@Test(priority=1575)
		private void  checkPurchaseAccountReportswithIncludePurchaseReturnVocherOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseAccountReportswithIncludePurchaseReturnVocherOption(), true);
		}
	

	 
	 
	 
	 @Test(priority=1576)
		private void checkVATDetailedReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATDetailedReport(), true);
		}

	
	
	 @Test(priority=1577)
		private void checkVATDetailedReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATDetailedReportWithFilterOption(), true);
		}
	 
				 
	 @Test(priority=1578)
		private void checkVATDetailedReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATDetailedReportBackTrackOption(), true);
		}
	 
	@Test(priority=1579)
		private void  checkVATDetailedReportswithLocalAmountBasedFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATDetailedReportswithLocalAmountBasedFilterOption(), true);
		}
	
	
	
	 
	 
	 @Test(priority=1580)
		private void checkVATSalesAccountReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATSalesAccountReport(), true);
		}

	
	
	 @Test(priority=1581)
		private void checkVATSalesAccountReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATSalesAccountReportWithFilterOption(), true);
		}
	 
				 
	 @Test(priority=1582)
		private void checkVATSalesAccountReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATSalesAccountReportBackTrackOption(), true);
		}
	 
	@Test(priority=1583)
		private void  checkVATSalesAccountReportswithLocalAmountBasedFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATSalesAccountReportswithIncludeSalesReturnVocher(), true);
		}
	
	
	 
	 
	 
	 @Test(priority=1584)
		private void checkVATSalesByCustomerReportwithVocherTypeSaleseandDelearsType() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATSalesByCustomerReportwithVocherTypeSalesandDelearsType(), true);
		}

	
	
	 @Test(priority=1585)
		private void checkVATSalesByCustomerReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATSalesByCustomerReportWithFilterOption(), true);
		}
	 
				 
	 @Test(priority=1586)
		private void checkVATSalesByCustomerReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATSalesByCustomerReportBackTrackOption(), true);
		}
	 
	 
	 @Test(priority=1587)
		private void  checkVATSalesByCustomerReportwithVocherTypePurchaseandDelearsType() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkVATSalesByCustomerReportwithVocherTypePurchaseandDelearsType(), true);
		}
	 
	 
	 
	 
	 @Test(priority=1588)
		private void checkSalesAdvanceVATReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesAdavanceVATReport(), true);
		}

	
	
	 @Test(priority=1589)
		private void checkSalesAdvanceVATReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesAdvanceVATReportWithFilterOption(), true);
		}
	 
				 
	 @Test(priority=1590)
		private void checkSalesAdvanceVATReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesAdvanceVATReportBackTrackOption(), true);
		}
	 
	
	 
//Register Menu
	 
	 @Test(priority=1591)
		private void checkOpeningBalanceRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkOpeningBalanceRegisterReport(), true);
		}

	
	
	 @Test(priority=1592)
		private void checkOpeningBalanceRegisterReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkOpeningBalanceRegisterReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1593)
		private void checkOpeningBalanceRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkOpeningBalanceRegisterReportBackTrackOption(), true);
		}
	 
	
	 
	 @Test(priority=1594)
		private void checkJournalEntriesRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkJournalEntriesRegisterReport(), true);
		}

	
	
	 @Test(priority=1595)
		private void checkJournalEntriesRegisterReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkJournalEntriesRegisterReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1596)
		private void checkJournalEntriesRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkJournalEntriesRegisterReportBackTrackOption(), true);
		}
	
	
	 
	 @Test(priority=1597)
		private void checkCreditNoteRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkCreditNoteRegisterReport(), true);
		}

	
	
	 @Test(priority=1598)
		private void checkCreditNoteRegisterReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkCreditNoteRegisterReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1599)
		private void checkCreditNoteRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkCreditNoteRegisterReportBackTrackOption(), true);
		}
	 
	
	
	
	 @Test(priority=1600)
		private void checkDebitNoteRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkDebitNoteRegisterReport(), true);
		}

	
	
	 @Test(priority=1601)
		private void checkDebitNoteRegisterReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkDebitNoteRegisterReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1602)
		private void checkDebitNoteRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkDebitNoteRegisterReportBackTrackOption(), true);
		}
	
	
	
	 
	 
	
	 @Test(priority=1603)
		private void checkReceiptRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkReceiptRegisterReport(), true);
		}

	
	
	 @Test(priority=1604)
		private void checkReceiptRegisterReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkReceiptRegisterReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1605)
		private void checkReceiptRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkReceiptRegisterReportBackTrackOption(), true);
		}
	
	
	
	 @Test(priority=1606)
		private void checkPaymentRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPaymentRegisterReport(), true);
		}

	
	
	 @Test(priority=1607)
		private void checkPaymentRegisterReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPaymentRegisterReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1608)
		private void checkPaymentRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPaymentRegisterReportBackTrackOption(), true);
		}
		
	
	 
	 @Test(priority=1609)
		private void checkPDCReceiptsRegisterReportwithAllIncludePDC() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPDCReceiptsRegisterReportwithAllIncludePDC(), true);
		}
	 
	 @Test(priority=1610)
		private void checkPDCReceiptRegisterReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPDCReceiptRegisterReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1611)
		private void checkPDCReceiptsRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPDCReceiptsRegisterReportBackTrackOption(), true);
		}
	
	
	 @Test(priority=1612)
		private void checkPDCReceiptsRegisterReportwithUnConvertedIncludePDCandincludeConvertedPDCandConsolidatedCheckboxes() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPDCReceiptsRegisterReportwithUnConvertedIncludePDCandincludeConvertedPDCandConsolidatedCheckboxes(), true);
		}
	 
	 
	 
	 
	 @Test(priority=1613)
		private void checkPDCPaymentRegisterReportwithAllIncludePDC() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPDCPaymentRegisterReportwithAllIncludePDC(), true);
		}
	 
	 @Test(priority=1614)
		private void checkPDCPaymentRegisterReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPDCPaymentRegisterReportWithFilterOption(), true);
		}
	 
				 
	 @Test(priority=1615)
		private void checkPDCPaymentRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPDCPaymentRegisterReportBackTrackOption(), true);
		}
	
	
	 @Test(priority=1616)
		private void checkPDCPaymentRegisterReportwithUnConvertedIncludePDCandincludeConvertedPDCandConsolidatedCheckboxes() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPDCPaymentRegisterReportwithUnConvertedIncludePDCandincludeConvertedPDCandConsolidatedCheckboxes(), true);
		}
	 
	 
	 
	 @Test(priority=1617)
		private void checkEntryJournalRegisterReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkEntryJournalRegisterReport(), true);
		}
	 
	 @Test(priority=1618)
		private void checkEntryJournalRegisterReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkEntryJournalRegisterReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1619)
		private void checkEntryJournalRegisterReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkEntryJournalRegisterReportBackTrackOption(), true);
		}
	
	 
	 @Test(priority=1620)
		private void checkEntryJournalDetailReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkEntryJournalDetailReport(), true);
		}
	 
	 @Test(priority=1621)
		private void checkEntryJournalDetailsReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkEntryJournalDetailsReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1622)
		private void checkEntryJournalDetailsReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkEntryJournalDetailsReportBackTrackOption(), true);
		}
		
		
	
	
	//Sales And Purchases Analysis 
	 
	 
	 @Test(priority=1623)
		private void checkSalesGroupedByCustomerReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesGroupedByCustomerReport(), true);
		}
	 
	// @Test(priority=1624)//not working as expected
		private void checkSalesGroupedByCustomerReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesGroupedByCustomerReportWithFilterOption(), true);
		}
	 
				 
	//@Test(priority=1625)
		private void checkSalesGroupedByCustomersReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesGroupedByCustomerReportBackTrackOption(), true);
		}
	
	@Test(priority=1626)
	private void checkSalesGroupedByCustomerReportswithSalesReturnVocherOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		FRP=new Focus10FinancialsReportsPage(getDriver());
		Assert.assertEquals(FRP.checkSalesGroupedByCustomerReportswithSalesReturnVocherOption(), true);
	}
	
	
	
	 
	 
	 @Test(priority=1627)
		private void checkSalesGroupedByItemReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesGroupedByItemReport(), true);
		}
	 
	 @Test(priority=1628)
		private void checkSalesGroupedByItemReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesGroupedByItemReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1629)
		private void checkSalesGroupedByItemPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesGroupedByItemPrintOption(), true);
		}
	
	@Test(priority=1630)
	private void checkSalesGroupedByItemReportswithSalesReturnVocherOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		FRP=new Focus10FinancialsReportsPage(getDriver());
		Assert.assertEquals(FRP.checkSalesGroupedByItemReportswithSalesReturnVocherOption(), true);
	}
	
	
	
	 @Test(priority=1631)
		private void checkSalesGroupedByDepartmentReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesGroupedByDepartmentReport(), true);
		}
		
	 
	 @Test(priority=1632)
		private void checkSalesGroupedByDepartmentReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesGroupedByDepartmentReportWithFilterOption(), true);
		}
	 
			 
	@Test(priority=1633)
		private void checkSalesGroupedByDepartmentPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkSalesGroupedByDepartmentPrintOption(), true);
		}
	
	@Test(priority=1634)
	private void checkSalesGroupedByDepartmentReportswithSalesReturnVocherOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		FRP=new Focus10FinancialsReportsPage(getDriver());
		Assert.assertEquals(FRP.checkSalesGroupedByDepartmentReportswithSalesReturnVocherOption(), true);
	}
	

	
	 
	 @Test(priority=1635)
		private void checkPurchasesGroupedByVendorReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchasesGroupedByVendorReport(), true);
		}
	 
	 @Test(priority=1636)
		private void checkPurchasesGroupedByVendorReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchasesGroupedByVendorReportWithFilterOption(), true);
		}
	
				 
	//@Test(priority=1637)
		private void checkPurchasesGroupedByVendorsReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchasesGroupedByVendorReportBackTrackOption(), true);
		}
	
	@Test(priority=1638)
	private void checkPurchasesGroupedByVendorReportswithPurchasesReturnVocherOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		FRP=new Focus10FinancialsReportsPage(getDriver());
		Assert.assertEquals(FRP.checkPurchasesGroupedByVendorReportswithPurchasesReturnVocherOption(), true);
	}

	
	 
	 
	 @Test(priority=1639)
		private void checkPurchasesGroupedByItemReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchasesGroupedByItemReport(), true);
		}
	 
	 @Test(priority=1640)
		private void checkPurchasesGroupedByItemReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchaseGroupedByItemReportWithFilterOption(), true);
		}
	 
		 
	@Test(priority=1641)
		private void checkPurchasesGroupedByItemPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchasesGroupedByItemPrintOption(), true);
		}
	
	@Test(priority=1642)
	private void checkPurchasesGroupedByItemReportswithPurchasesReturnVocherOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		FRP=new Focus10FinancialsReportsPage(getDriver());
		Assert.assertEquals(FRP.checkPurchasesGroupedByItemReportswithPurchasesReturnVocherOption(), true);
	}
	
	
	
	 @Test(priority=1643)
		private void checkPurchasesGroupedByDepartmentReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchasesGroupedByDepartmentReport(), true);
		}
	 
	 @Test(priority=1644)
		private void checkPurchasesGroupedByDepartmentReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchasesGroupedByDepartmentReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1645)
		private void checkPurchasesGroupedByDepartmentPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPurchasesGroupedByDepartmentPrintOption(), true);
		}
	
	@Test(priority=1646)
	private void checkPurchasesGroupedByDepartmentReportswithPurchasesReturnVocherOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		FRP=new Focus10FinancialsReportsPage(getDriver());
		Assert.assertEquals(FRP.checkPurchasesGroupedByDepartmentReportswithPurchasesReturnVocherOption(), true);
	}


	 
	//Miscellaneous Analysis 
	 
	 @Test(priority=1647)
		private void checkPeakAndLowBalancesAmountReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPeakAndLowBalancesAmountReport(), true);
		}
	 
	 @Test(priority=1648)
		private void checkPeakAndLowBalancesReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPeakAndLowBalancesReportWithFilterOption(), true);
		}
	 
			 
	@Test(priority=1649)
		private void checkPeakAndLowBalancesReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPeakAndLowBalancesReportBackTrackOption(), true);
		}
	
	 @Test(priority=1650)
		private void checkPeakAndLowBalancesAmountReportwithSelectAll() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkPeakAndLowBalancesReportwithSelectAll(), true);
		}
	 
 
	 
	 @Test(priority=1651)
			private void checkComparitiveAnalysisReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkComparitiveAnalysisReport(), true);
			}
		 
		 @Test(priority=1652)
			private void checkComparitiveAnalysisReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkComparitiveAnalysisReportWithFilterOption(), true);
			}
		 
					 
		@Test(priority=1653)
			private void checkComparitiveAnalysisReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkComparitiveAnalysisReportBackTrackOption(), true);
			}
		
		 @Test(priority=1654)
			private void checkComparitiveAnalysisAmountReportwithSelectAll() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkComparitiveAnalysisReportwithSelectAll(), true);
			}
		 
	 
	 
	 @Test(priority=1655)
		private void checkTransactionsTypeAnalysisReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkTransactionsTypeAnalysisReport(), true);
		}
	 
	 @Test(priority=1656)
		private void checkTransactionsTypeAnalysisReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkTransactionsTypeAnalysisReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1657)
		private void checkTransactionsTypeAnalysisReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkTransactionsTypeAnalysisReportBackTrackOption(), true);
		}
	
	 @Test(priority=1658)
		private void checkTransactionsTypeAnalysisAmountReportwithSelectAll() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkTransactionsTypeAnalysisReportwithSelectAll(), true);
		}
	 
		 
	
	 @Test(priority=1659)
		private void checkABCAnalysiswithReportBasedonBalance() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkABCAnalysiswithReportBasedonBalance(), true);
		}
		
		 @Test(priority=1660)
			private void checkABCAnalysisReportwithSelectAll() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkABCAnalysisReportwithSelectAll(), true);
			}
	 @Test(priority=1661)
		private void checkABCAnalysiswithReportBasedonTurnover() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkABCAnalysiswithReportBasedonTurnover(), true);
		}
	 
	 @Test(priority=1662)
		private void checkABCAnalysisReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkABCAnalysisReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1663)
		private void checkABCAnalysisReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkABCAnalysisReportBackTrackOption(), true);
		}
	
	
		 
	 
	 
	 
	 @Test(priority=1664)
		private void checkInterestCalculationReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkInterestCalculationReport(), true);
		}
	 @Test(priority=1665)
		private void checkInterestCalculationReportwithSelectAll() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkInterestCalculationReportwithSelectAll(), true);
		}
	
	 @Test(priority=1666)
		private void checkInterestCalculationReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkInterestCalculationReportWithFilterOption(), true);
		}
	 
				 
	@Test(priority=1667)
		private void checkInterestCalculationReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkInterestCalculationReportBackTrackOption(), true);
		}
	
	
	
	 
	 //Audit Trail
	 
	 @Test(priority=1668)
		private void checkAuditTrailTransactionsReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkAuditTrailTransactionsReport(), true);
		}
	 @Test(priority=1669)
		private void checkAuditTrailTransactionsPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkAuditTrailTransactionsPrintOption(), true);
		}
	 
	//@Test(priority=1670)//not getting rows
		private void checkAuditTrailTransactionsReportWithFilterOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkAuditTrailTransactionsReportWithFilterOption(), true);
		}
	 
	
				 
	@Test(priority=1671)
		private void checkAuditTrailTransactionsReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkAuditTrailTransactionsReportBackTrackOption(), true);
		}
	
	
	 @Test(priority=1672)
		private void checkAuditTrailLoginsReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkAuditTrailLoginsReport(), true);
		}
	 
	 @Test(priority=1673)
		private void checkAuditTrailLoginsPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkAuditTrailLoginsPrintOption(), true);
		}
	 

	
	
	 //Account Query
	 
	 @Test(priority=1674)
		private void checkAccountQueryReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkAccountQueryReport(),true);
		}
	 
	
	 
	//Transaction Authorization Report 
	 
	 @Test(priority=1675)
		private void checkTransactionAuthorizationReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkTransactionAuthorizationReport(), true);
		}
	 
	@Test(priority=1676)
		private void checkTransactionAuthorizationReportPrintOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkTransactionAuthorizationReportPrintOption(), true);
		}
	 
	 @Test(priority=1677)
		private void checkTransactionAuthorizationReportBackTrackOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			Assert.assertEquals(FRP.checkTransactionAuthorizationReportBackTrackOption(), true);
		}
		
	
	// @Test(priority=1678)//not implemented upto 31/01/2021
		private void checkVATSummaryReport() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			FRP=new Focus10FinancialsReportsPage(getDriver());
			
			Assert.assertEquals(FRP.checkVATSummaryReport(),true);

		}
	 
	//chequeDiscounting
		
		 @Test(priority=1679)
			private void checkChequeDiscountingReportforBankAccountasBank() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkChequeDiscountingReportforBankAccountasBank(), true);
			}
	 
		 @Test(priority=1680)
			private void checkChequeDiscountingReportforBankAccountasHDFC() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkChequeDiscountingReportforBankAccountasHDFC(), true);
			}
		
	//MasterAuthorization
		 @Test(priority=1681)
			private void checkMasterAuthorizationReportforMasterTypeasWarehouse() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkMasterAuthorizationReportforMasterTypeasWarehouse(), true);
			}
		 
		 //LedgerDetails
		 
		 @Test(priority=1682)
			private void checkFinancialLedgerDetailsReportwithFilterOptionasRefine() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkFinancialLedgerDetailsReportwithFilterOptionasRefine(), true);
			}
		 
		 //Financial--FinalAccount--TrailBalance
		 @Test(priority=1683)
			private void checkTrailBalanceReportinFinalAccountswithFilterOptionasRefineforParticulars() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkTrailBalanceReportinFinalAccountswithFilterOptionasRefineforParticulars(), true);
			}
		
		 
		 //ReceivableandPayable--CustomerDetails--CustomerStatemets
		 @Test(priority=1684)
			private void checkCustomerDetailsReportinReceivablePayableMenuwithFilterOptionasRefineforVouchers() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkCustomerDetailsReportinReceivablePayableMenuwithFilterOptionasRefineforVouchers(), true);
			}
		 
		 
		 @Test(priority=1685)
			private void checkVendorSummaryReportinReceivablePayableMenuwithFilterOptionasRefineforAccount() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
			{
				FRP=new Focus10FinancialsReportsPage(getDriver());
				Assert.assertEquals(FRP.checkVendorSummaryReportinReceivablePayableMenuwithFilterOptionasRefineforAccount(), true);
			}
}






