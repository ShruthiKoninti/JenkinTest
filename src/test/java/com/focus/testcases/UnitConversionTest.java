package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.UnitConversionPage;
import com.focus.base.BaseEngine;

public class UnitConversionTest extends BaseEngine 
{
	UnitConversionPage UCP;
	
	

	 @Test(priority=166)
	 public void checkDeletingUnitConversion() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 UCP=new UnitConversionPage(getDriver());
		 Assert.assertEquals(UCP.checkDeletingUnitConversion(), true);
	 }
	
	 @Test(priority=167)
	 public void checkSavingUnitConversion() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 UCP=new UnitConversionPage(getDriver());
		 Assert.assertEquals(UCP.checkSavingUnitConversion(), true);
	 }

	 @Test(priority=168)
	 public void checkUpdatingUnitConversion() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 UCP=new UnitConversionPage(getDriver());
		 Assert.assertEquals(UCP.checkUpdatingUnitConversion(), true);
	 }
	 
	
	

}
