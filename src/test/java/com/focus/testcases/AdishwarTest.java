package com.focus.testcases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;
import com.focus.Pages.AdishwarPage;
import com.focus.base.BaseEngine;

public class AdishwarTest extends BaseEngine
{

	static AdishwarPage ap;
	
	
	
	@Test(priority=80)
	public void checkLogin() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		ap=new AdishwarPage(getDriver());
		Assert.assertEquals(ap.checkLogin(), true);	
	}

	@Test(priority=81)
	public void checkLoop() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		ap=new AdishwarPage(getDriver());
		Assert.assertEquals(ap.checkLoop(), true);	
	}
	
	
	
	
}
