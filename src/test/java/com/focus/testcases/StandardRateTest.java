package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.StandardRatePage;
import com.focus.base.BaseEngine;

public class StandardRateTest extends BaseEngine
{
	StandardRatePage SRP;
	
	
	 @Test(priority=215)
	 public void CheckLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SRP=new StandardRatePage(getDriver());
		 Assert.assertEquals(SRP.CheckLogin(), true);
	 }
	
	 
	 @Test(priority=217)
	 public void checkStandardRateDeleteOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SRP=new StandardRatePage(getDriver());
		 Assert.assertEquals(SRP.checkStandardRateDeleteOption(), true);
	 }
	 
	 @Test(priority=218)
	 public void checkSavingStandardRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SRP=new StandardRatePage(getDriver());
		 Assert.assertEquals(SRP.checkSavingStandardRate(), true);
	 }
	 
	 @Test(priority=219)
	 public void checkEditingStandardRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SRP=new StandardRatePage(getDriver());
		 Assert.assertEquals(SRP.checkEditingStandardRate(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=221)
	 public void checkLogoutStandardRatePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 SRP=new StandardRatePage(getDriver());
		 Assert.assertEquals(SRP.checkLogoutStandardRatePage(), true);
	 }
	 
}
