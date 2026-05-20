package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.ItemPropertiesPage;
import com.focus.base.BaseEngine;

public class ItemPropertiesTest extends BaseEngine{
	
	ItemPropertiesPage IPP;
	
	
	
	@Test(priority=500)
	public void checkLogin() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		IPP=new ItemPropertiesPage(getDriver());
		Assert.assertEquals(IPP.checkLogin(), true);
	}
	
	
	@Test(priority=501)
	public void checkSavingMaterialReceiptNotesVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		IPP=new ItemPropertiesPage(getDriver());
		Assert.assertEquals(IPP.checkSavingMaterialReceiptNotesVoucher(), true);
	}

	@Test(priority=502)
	public void checkStockLedgerReportWithoutEnablingDonotUpdateStockChkBox() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		IPP=new ItemPropertiesPage(getDriver());
		Assert.assertEquals(IPP.checkStockLedgerReportWithoutEnablingDonotUpdateStockChkBox(), true);
	}
	
	
	
	
	@Test(priority=503)
	public void checkEnablingDonotUpdateStockChkBoxinItemProperties() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		
		IPP=new ItemPropertiesPage(getDriver());
		Assert.assertEquals(IPP.checkEnablingDonotUpdateStockChkBoxinItemProperties(), true);
	}
	
	
	
	@Test(priority=504)
	public void checkStockLedgerReportWithEnablingDonotUpdateStockChkBox() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		IPP=new ItemPropertiesPage(getDriver());
		Assert.assertEquals(IPP.checkStockLedgerReportWithEnablingDonotUpdateStockChkBox(), true);
	}
	
	@Test(priority=505)//voucher should save without negative statement
	public void checkSalesInvoiceVATByEnablingDonotUpdateStockChkBox() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		IPP=new ItemPropertiesPage(getDriver());
		Assert.assertEquals(IPP.checkSalesInvoiceVATByEnablingDonotUpdateStockChkBox(), true);
	}
	
	@Test(priority=506)
	public void checkLogoutItemPropertiesPage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		IPP=new ItemPropertiesPage(getDriver());
		Assert.assertEquals(IPP.checkLogoutItemPropertiesPage(), true);
	}
	
	
}
