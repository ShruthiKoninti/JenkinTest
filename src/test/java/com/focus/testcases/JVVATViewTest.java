package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.JVVATViewPage;
import com.focus.base.BaseEngine;

public class JVVATViewTest extends BaseEngine
{
	JVVATViewPage JV;

	@Test(priority=337)
	public void checkSavingJVVATVocher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		JV=new JVVATViewPage(getDriver());
		Assert.assertEquals(JV.checkSavingJVVATVocher1(), true);
	}

	@Test(priority=338)
	public void checkEditingJVVATVocher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		JV=new JVVATViewPage(getDriver());
		Assert.assertEquals(JV.checkEditingJVVATVocher1(), true);
	}	 

	@Test(priority=339)
	public void checkSavingJVVATVocher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		JV=new JVVATViewPage(getDriver());
		Assert.assertEquals(JV.checkSavingJVVATVocher2(), true);
	}

	@Test(priority=340)
	public void checkDeletingJVVATViewVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		JV=new JVVATViewPage(getDriver());
		Assert.assertEquals(JV.checkDeletingJVVATViewVoucher2(), true);
	}


	@Test(priority=341)
	public void checkVerifingDetailsOfSavedJVVATViewVoucherInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		JV=new JVVATViewPage(getDriver());
		Assert.assertEquals(JV.checkVerifingDetailsOfSavedJVVATViewVoucherInHomepage(), true);
	}


}
