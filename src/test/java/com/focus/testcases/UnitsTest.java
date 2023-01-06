package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.UnitsPage;
import com.focus.Pages.UnitsPage;
import com.focus.base.BaseEngine;

public class UnitsTest extends BaseEngine
{
	UnitsPage UP;
	

	 @Test(priority=128)
	 public void checkUnitMasterInMAsterMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 UP=new UnitsPage(getDriver());
		 Assert.assertEquals(UP.checkUnitMasterInMAsterMenu(), true);
	 }
	
	 @Test(priority=129)
	 public void checkSavingUnitsInUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 UP=new UnitsPage(getDriver());
		 Assert.assertEquals(UP.checkSavingUnitsInUnitsMaster(), true);
	 }
	
	 @Test(priority=130)
	 public void checkEditingInSavedUnitsINUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 UP=new UnitsPage(getDriver());
		 Assert.assertEquals(UP.checkEditingInSavedUnitsINUnitsMaster(), true);
	 }
	
	 @Test(priority=131)
	 public void checkUpdatedUnitsMasterDisplayInUnits() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 UP=new UnitsPage(getDriver());
		 Assert.assertEquals(UP.checkUpdatedUnitsMasterDisplayInUnits(), true);
	 }
	
	 @Test(priority=132)
	 public void checkDeleteUnitsInUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 UP=new UnitsPage(getDriver());
		 Assert.assertEquals(UP.checkDeleteUnitsInUnitsMaster(), true);
	 }
	 
}
