package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.DepartmentPage;
import com.focus.Pages.ItemsPage;
import com.focus.base.BaseEngine;

public class DepartmentTest extends BaseEngine
{
	DepartmentPage DP;
	
	 @Test(priority=146)
	 public void CheckLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.CheckLogin(), true);
	 }
	
	 @Test(priority=148)
	 public void checkDepartmentMasterInMAsterMenu() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkDepartmentMasterInMAsterMenu(), true);
	 }
	
	 @Test(priority=149)
	 public void checkSavingDepartmentInDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkSavingDepartmentInDepartmentMaster(), true);
	 }
	
	 @Test(priority=150)
	 public void checkEditingInSavedDepartmentINDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkEditingInSavedDepartmentINDepartmentMaster(), true);
	 }
	
	 @Test(priority=151)
	 public void checkUpdatedDepartmentMasterDisplayInDepartment() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkUpdatedDepartmentMasterDisplayInDepartment(), true);
	 }
	
	 @Test(priority=152)
	 public void checkDeleteInDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkDeleteInDepartmentMaster(), true);
	 }
	 
	 /*
	 @Test(priority=153)
	 public void checkDepartmentsListBeforeImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkDepartmentsListBeforeImportingDataFromExcel(), true);
	 }
	 
	 
	 @Test(priority=154)
	 public void checkAdvanceMasterImportExportDepartment() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkAdvanceMasterImportExportDepartment(), true);
	 }
	 
	 @Test(priority=155)
	 public void checkImportDatainAdvanceMasterImportExportDepartment() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkImportDatainAdvanceMasterImportExportDepartment(), true);
	 }
	 
	 @Test(priority=156)
	 public void checkDepartmentsListAfterImportingDataFromExcel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkDepartmentsListAfterImportingDataFromExcel(), true);
	 }
	 
	 @Test(priority=157)
	 public void checkImportedDepartmentDelete() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkImportedDepartmentDelete(), true);
	 }
	 

	 */
	 ////////////////////////Checking Time in extra field////////////////////
	 
	 
	 @Test(priority=158)
	 public void CreateExtraTabAndExtraFieldForCheckingTimeInDepartment() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.CreateExtraTabAndExtraFieldForCheckingTimeInDepartment(), true);
	 }
	 
	 @Test(priority=159)
		public void checkLogoutAfterExtraTabAdded() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		 DP=new DepartmentPage(getDriver());
			Assert.assertEquals(DP.checkLogoutDepartmentPage(), true);
		}
		
		
	 @Test(priority=160)
		 public void CheckLoginAfterExtraTabAdded() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		 {
			 DP=new DepartmentPage(getDriver());
			 Assert.assertEquals(DP.CheckLogin(), true);
		 }
	 
	 @Test(priority=161)
	 public void CheckPresentTimeInCreatedTabDepartment() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.CheckPresentTimeInCreatedTabDepartment(), true);
	 }
	 
	 ////////////////////////Checking New Groups and SubGroups added/////////////////////////////
	 
		
	 @Test(priority=162)
	 public void checkSavingDepartmentGroupinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkSavingDepartmentGroupinDepartmentMaster(), true);
	 }
	 
	 
	 @Test(priority=163)
	 public void checkSavingDepartmentSubGroupinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkSavingDepartmentSubGroupinDepartmentMaster(), true);
	 }
	 

	 @Test(priority=164)
	 public void checkCreatedGroupDisplayedonLeftPanel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkCreatedGroupDisplayedonLeftPanel(), true);
	 }
	 
	
	 @Test(priority=165)
	 public void checkSavingDepartmentsinSubGroupsinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkSavingDepartmentsinSubGroupsinDepartmentMaster(), true);
	 }
	
	 @Test(priority=166)
	 public void checkDepartmentsSavedinGroupsandSubGroupsinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkDepartmentsSavedinGroupsandSubGroupsinDepartmentMaster(), true);
	 }
	 
	 
	 
	 @Test(priority=167)
	 public void checkDeleteGroupDepartmentinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkDeleteGroupDepartmentinDepartmentMaster(), true);
	 }
	 
	 @Test(priority=168)
	 public void CheckEditingTimeInCreatedTabDepartment() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.CheckEditingTimeInCreatedTabDepartment(), true);
	 }
	 
	 
///////////////Creating Extra Fields//////////////////////
	 
/*	 
	 
		@Test(priority=169)
		public void checkCreatingExtraFieldsinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		DP=new DepartmentPage(getDriver());
		Assert.assertEquals(DP.checkCreatingExtraFieldsinDepartmentMaster(), true);
		}
		
		
		
		@Test(priority=170)
		public void checkAddedExtraFieldinGeneralTabofDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		DP=new DepartmentPage(getDriver());
		Assert.assertEquals(DP.checkAddedExtraFieldinGeneralTabofDepartmentMaster(), true);
		}
		
		
		
		
		@Test(priority=171)
		public void checkEditandUpdateExtraFieldinCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		DP=new DepartmentPage(getDriver());
		Assert.assertEquals(DP.checkEditandUpdateExtraFieldinCustomizeMaster(), true);
		}
		
		
		
		
		
		
		@Test(priority=172)
		public void checkUpdatedExtraFieldinGeneralTabofDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		DP=new DepartmentPage(getDriver());
		Assert.assertEquals(DP.checkUpdatedExtraFieldinGeneralTabofDepartmentMaster(), true);
		}
		
		
		
		
		@Test(priority=173)
		public void checkDeleteExtraFieldinCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		DP=new DepartmentPage(getDriver());
		Assert.assertEquals(DP.checkDeleteExtraFieldinCustomizeMaster(), true);
		}
		
		
		@Test(priority=174)
		public void checkDeletedExtraFieldinGeneralTabofDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		DP=new DepartmentPage(getDriver());
		Assert.assertEquals(DP.checkDeletedExtraFieldinGeneralTabofDepartmentMaster(), true);
		}
		*/
		
			 
		
		
		@Test(priority=175)
		public void checkLogoutDepartmentPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
		{
		DP=new DepartmentPage(getDriver());
		Assert.assertEquals(DP.checkLogoutDepartmentPage(), true);
		}
	 
	 
	 
}
