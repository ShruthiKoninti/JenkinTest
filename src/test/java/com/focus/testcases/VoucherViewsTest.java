package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.VoucherViewsPage;
import com.focus.base.BaseEngine;

public class VoucherViewsTest extends BaseEngine {
	VoucherViewsPage vvp;
	
	
	@Test(priority=3000)
	public void checkLogin() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkLogin(), true);
	}
	
	@Test(priority=3001)
	public void checkEditingtheFieldsinViewWithoutSavingView() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkEditingtheFieldsinViewWithoutSavingView(), true);
	}
	
	@Test(priority=3002)
	public void checkCaptionBeforeEditingViewinVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkCaptionBeforeEditingViewinVoucher(), true);
	}
	
	@Test(priority=3003)
	public void checkCreatingViewinViewTabandChangeCaption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkCreatingViewinViewTabandChangeCaption(), true);
	}
	
	@Test(priority=3004)
	public void checkCaptionAfterEditingViewinVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkCaptionAfterEditingViewinVoucher(), true);
	}
	
	@Test(priority=3005)
	public void checkSavingViewbySearchingUserfromSerchTxtinViewsTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewbySearchingUserfromSerchTxtinViewsTab(), true);
	}
	
	@Test(priority=3006)
	public void checkEditingEditScreenFieldsinViews() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkEditingEditScreenFieldsinViews(), true);
	}
	
	@Test(priority=3007)
	public void checkUpwardandDownwardArrowsinViews() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkUpwardandDownwardArrowsinViews(), true);
	}
	
	//@Test(priority=3008)
	public void checkAddingNewFieldinEditLayoutTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkAddingNewFieldinEditLayoutTab(), true);
	}
	
	//@Test(priority=3009)
	public void checkRefreshViewAfterAddingNewFieldinEditLayoutTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkRefreshViewAfterAddingNewFieldinEditLayoutTab(), true);
	}
	
	@Test(priority=3010)
	public void checkCreatingViewinViewsTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkCreatingViewinViewsTab(), true);
	}
	
	@Test(priority=3011)
	public void checkVoucherAfterMultipleViewsCreatedatHomeScreen() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkVoucherAfterMultipleViewsCreatedatHomeScreen(), true);
	}
	
	@Test(priority=3012)
	public void checkImportandExportViewsinViewsTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkImportandExportViewsinViewsTab(), true);
	}
	
	@Test(priority=3013)
	public void checkSavingViewwithDefaultValue() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewwithDefaultValue(), true);
	}
	
	@Test(priority=3014)
	public void checkSavingVoucherwithDefaultValueinPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingVoucherwithDefaultValueinPurchaseVoucher(), true);
	}
	
	@Test(priority=3015)
	public void checkSavingViewwithLoadDefaults() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewwithLoadDefaults(), true);
	}
	
	@Test(priority=3016)
	public void checkSavingViewWithDefaultValueasGroupID() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewWithDefaultValueasGroupID(), true);
	}
	
	@Test(priority=3017)
	public void checkSavingPurchseVoucherwithDefaultValueasGroup() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingPurchseVoucherwithDefaultValueasGroup(), true);
	}
	
	@Test(priority=3018)
	public void checkSavingViewWithDefaultValueasGroupandSubGroup() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewWithDefaultValueasGroupandSubGroup(), true);
	}
	
	
	@Test(priority=3019)
	public void checkSavingPurchaseVoucherwithDefaultValuesofGroupandSubGroup() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingPurchaseVoucherwithDefaultValuesofGroupandSubGroup(), true);
	}
	
	@Test(priority=3020)
	public void checkSavingViewforFooterFieldsWithSpecialCharacters() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewforFooterFieldsWithSpecialCharacters(), true);
	}
	
	@Test(priority=3021)
	public void checkSavingViewWithInsertInformationField() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewWithInsertInformationField(), true);
	}
	
	@Test(priority=3022)
	public void checkSavingPurchaseVoucherwithInsertInfoFieldinView() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingPurchaseVoucherwithInsertInfoFieldinView(), true);
	}
	
	@Test(priority=3023)
	public void checkSavingViewWithInsertInformationColumn() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewWithInsertInformationColumn(), true);
	}
	
	@Test(priority=3024)
	public void checkSavingPurchaseVoucherwithInsertInfoColumninView() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingPurchaseVoucherwithInsertInfoColumninView(), true);
	}
	
	@Test(priority=3025)
	public void checkSavingViewWithMultipleRowsinFilterCriteria() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewWithMultipleRowsinFilterCriteria(), true);
	}
	
	@Test(priority=3026)
	public void checkSavedFilterCriteriainPurchaseVoucherforPurchaseAC() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavedFilterCriteriainPurchaseVoucherforPurchaseAC(), true);
	}
	
	@Test(priority=3027)
	public void checkSavingViewAfterEditingFilterCriteria() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewAfterEditingFilterCriteria(), true);
	}
	
	@Test(priority=3028)
	public void checkEditedFilterCriteriainPurchaseVoucherforPurchaseAC() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkEditedFilterCriteriainPurchaseVoucherforPurchaseAC(), true);
	}
	
	@Test(priority=3029)
	public void checkSavingViewwithDefaultValueinSalesReturnsVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewwithDefaultValueinSalesReturnsVoucher(), true);
	}
	
	@Test(priority=3030)
	public void checkLogoutandLoginwithUseralloptionsst() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkLogoutandLoginwithUseralloptionsst(), true);
	}
	
	@Test(priority=3031)
	public void checkSavingSalesReturnsVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingSalesReturnsVoucher(), true);
	}
	
	@Test(priority=3032)
	public void checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkLogoutAndLoginWithSU(), true);
	}
	
	@Test(priority=3033)
	public void checkSavingOpeningStockVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingOpeningStockVoucher(), true);
	}
	
	@Test(priority=3034)
	public void checkSavingViewinSalesInvoiceVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingViewinSalesInvoiceVoucher(), true);
	}
	
	@Test(priority=3035)
	public void checkSavingSalesOrderVoucherwithStockItem() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingSalesOrderVoucherwithStockItem(), true);
	}
	
	@Test(priority=3036)
	public void checkLogoutandLoginwithUseralloptionsstAgain() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkLogoutandLoginwithUseralloptionsst(), true);
	}
	
	@Test(priority=3037)
	public void checkSavingSalesInvoiceVATthroughConvertOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkSavingSalesInvoiceVATthroughConvertOption(), true);
	}
	
	@Test(priority=3038)
	public void checkLogoutVoucherViewsPage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		vvp=new VoucherViewsPage(getDriver());
		Assert.assertEquals(vvp.checkLogoutVoucherViewsPage(), true);
	}
	
}
