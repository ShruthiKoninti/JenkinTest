package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.CurrencyPage;
import com.focus.base.BaseEngine;

public class CurrencyTest extends BaseEngine
{
	CurrencyPage CP;
	
	
	 @Test(priority=134)
	 public void checkOpenCurrencyMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 CP=new CurrencyPage(getDriver());
		 Assert.assertEquals(CP.checkOpenCurrencyMenu(), true);
	 }
	 
	 @Test(priority=135)
	 public void checkUpdatingISOCurrencyCodeSLLCoinName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 CP=new CurrencyPage(getDriver());
		 Assert.assertEquals(CP.checkUpdatingISOCurrencyCodeSLLCoinName(), true);
	 }
	 
	 @Test(priority=136)
	 public void checkUpdatingAgainISOCurrencyCodeSLLCoinName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 CP=new CurrencyPage(getDriver());
		 Assert.assertEquals(CP.checkUpdatingAgainISOCurrencyCodeSLLCoinName(), true);
	 }
	 

}
