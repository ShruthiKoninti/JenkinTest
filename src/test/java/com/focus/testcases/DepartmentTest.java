package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.DepartmentPage;
import com.focus.base.BaseEngine;

public class DepartmentTest extends BaseEngine
{
	DepartmentPage DP;
	
	

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
	 

	 
	 ////////////////////////Checking New Groups and SubGroups added/////////////////////////////
	 
		
	 @Test(priority=158)
	 public void checkSavingDepartmentGroupinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkSavingDepartmentGroupinDepartmentMaster(), true);
	 }
	 
	 
	 @Test(priority=159)
	 public void checkSavingDepartmentSubGroupinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkSavingDepartmentSubGroupinDepartmentMaster(), true);
	 }
	 

	 @Test(priority=160)
	 public void checkCreatedGroupDisplayedonLeftPanel() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkCreatedGroupDisplayedonLeftPanel(), true);
	 }
	 
	
	 @Test(priority=161)
	 public void checkSavingDepartmentsinSubGroupsinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkSavingDepartmentsinSubGroupsinDepartmentMaster(), true);
	 }
	
	 @Test(priority=162)
	 public void checkDepartmentsSavedinGroupsandSubGroupsinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkDepartmentsSavedinGroupsandSubGroupsinDepartmentMaster(), true);
	 }
	 
	 
	 
	 @Test(priority=163)
	 public void checkDeleteGroupDepartmentinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DP=new DepartmentPage(getDriver());
		 Assert.assertEquals(DP.checkDeleteGroupDepartmentinDepartmentMaster(), true);
	 }
	 
	 
	 
	 
}
