package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.BuyerSellerPricebookAndVoucherPrintPDFEmailValidationsPage;
import com.focus.Pages.MasterPreloadPage;
import com.focus.base.BaseEngine;

public class MasterPreloadTest extends BaseEngine 
{

	MasterPreloadPage MPP;

	// Master Preload Values in Purchase Voucher
	
	
	@Test(priority=2999)
	public void checkRestoreAutomationCompany() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkRestoreAutomationCompany(), true);
	}
	
	@Test(priority=3000)
	public void checkCreatingCostCenterWithStockCostCenter() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkCreatingCostCenterWithStockCostCenter(), true);
	}
	
	@Test(priority=3001)
	public void checkCreatingCostCenterWithSTDRateCostCenter() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkCreatingCostCenterWithSTDRateCostCenter(), true);
	}
	
	@Test(priority=3002)
	public void checkAddingExtraFieldInGeneralTabAsCostCenter() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkAddingExtraFieldInGeneralTabAsCostCenter(), true);
	}
	
	@Test(priority=3003)
	public void checkAssigningCostCenterToStockItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkAssigningCostCenterToStockItem(), true);
	}
	
	@Test(priority=3004)
	public void checkAssigningCostCenterToSTDRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkAssigningCostCenterToSTDRateItem(), true);
	}

	@Test(priority=3005)
	public void checkAddingFieldsInEditLayoutBodyTabAsCostCenterBodyInPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkAddingFieldsInEditLayoutBodyTabAsCostCenterBodyInPurchaseVoucher(), true);
	}
	
	@Test(priority=3006)
	public void CheckSavingPurchaseVoucherWithStockItemAndSTDRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.CheckSavingPurchaseVoucherWithStockItemAndSTDRateItem(), true);
	}
	
	@Test(priority=3007)
	public void CheckCostCenterInSavedVoucherWithStockItemAndSTDRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.CheckCostCenterInSavedVoucherWithStockItemAndSTDRateItem(), true);
	}
	
	@Test(priority=3008)
	public void CheckEditingPurchaseVoucher1WithStockItemAndSTDRateItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.CheckEditingPurchaseVoucher1WithStockItemAndSTDRateItem(), true);
	}
	
	@Test(priority=3009)
	public void checkSavedPurchaseVoucherAfterEditingVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkSavedPurchaseVoucherAfterEditingVoucher(), true);
	}
	
	
	
	
	// Inventory Item Rate Options
	
	@Test(priority=3015)
	public void checkUpdatingItemRateCannotBeBelowCostForSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkUpdatingItemRateCannotBeBelowCostForSalesInvoiceVAT(), true);
	}
	
	@Test(priority=3016)
	public void checkValidatingCannotBeBelowCostDropdownAndSavingVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkValidatingCannotBeBelowCostDropdownAndSavingVoucher(), true);
	}
	
	@Test(priority=3017)
	public void checkUpdatingItemRateCannotBeChangedForSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkUpdatingItemRateCannotBeChangedForSalesInvoiceVAT(), true);
	}
	
	@Test(priority=3018)
	public void checkValidatingCannotBeChangedInSavedVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkValidatingCannotBeChangedInSavedVoucher(), true);
	}
	
	@Test(priority=3019)
	public void checkUpdatingItemRateEditableForSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkUpdatingItemRateEditableForSalesInvoiceVAT(), true);
	}
	
	
	
	
	
	@Test(priority=30000)
	public void checkEraseAllTransactionsAfterCompletingAllValidations() throws Exception
	{
		MPP=new MasterPreloadPage(getDriver());
		Assert.assertEquals(MPP.checkEraseAllTransactionsAfterCompletingAllValidations(), true);
	}
	
	
	
	
	
	
	
	
	
	
	
	
}