package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.InfoPanelPage;
import com.focus.base.BaseEngine;

public class InfoPanelTest extends BaseEngine{
	
	InfoPanelPage Ipp;
	
	
	
	@Test(priority=700)
	public void checkLogin() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkLogin(), true);
	
	}
	

	//@Test(priority=701)//exist in Backup
	public void EnableDoNotUseRealTimeRate() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkEnableDoNotUseRealTimeRate(), true);
	
	}
	
	
	
	@Test(priority=701)
	public void checkCreatingInfoPanelinOpeningStockVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkCreatingInfoPanelinOpeningStockVoucher(), true);
	
	}
	
	
	
	@Test(priority=702)
	public void checkDragandDropofCreatedInfoPanel() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkDragandDropofCreatedInfoPanel(), true);
	
	}
	
	
	
	
	@Test(priority=703)
	public void checkSavingOpeningStockVoucher1WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkSavingOpeningStockVoucher1WithBatchItem(), true);
	
	}
	
	
	
	
	@Test(priority=704)
	public void checkQuantityValueinInfoPanelinOpeningStocksVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkQuantityValueinInfoPanelinOpeningStocksVoucher(), true);
	
	}
	
	
	
	@Test(priority=705)
	public void checkCreatingInfoPanelforAvgRateinOpeningStocksVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkCreatingInfoPanelforAvgRateinOpeningStocksVoucher(), true);
	
	}
	
	
	
	@Test(priority=706)
	public void checkDragandDropofCreatedInfoPanelofAvgRate() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkDragandDropofCreatedInfoPanelofAvgRate(), true);
	
	}

	
	@Test(priority=707)
	public void checkSavingOpeningStockVoucherWithBinItemWithAutoAllocationOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkSavingOpeningStockVoucherWithBinItemWithAutoAllocationOption(), true);
	
	}
	
	
	
	@Test(priority=708)
	public void checkAverageRateinInfoPanelinOpeningStockVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkAverageRateinInfoPanelinOpeningStockVoucher(), true);
	
	}

	
	
	
	@Test(priority=709)
	public void checkLogoutInfoPanelPage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Ipp=new InfoPanelPage(getDriver());
		Assert.assertEquals(Ipp.checkLogoutInfoPanelPage(), true);
	
	}
}
