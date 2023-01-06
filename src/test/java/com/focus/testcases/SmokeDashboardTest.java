package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.SmokeDashboardPage;
import com.focus.base.BaseEngine;

public class SmokeDashboardTest extends BaseEngine 
{
	
	SmokeDashboardPage HPDP;

	 @Test(priority=143)
	 public void verifyDashBboardOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkDashboard(), true);
	 }
	 	
	 @Test(priority=154)
	 public void verifySaveInDashBboardNewOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkSaveBtnWithInputDashboardName(), true);
	 }
	
	 @Test(priority=157)
	 public void verifySaveDashboardInSelectDashboard() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkSavedDashboardFromSelectDashboard(), true);
	 }	
	 
	 @Test(priority=160)
	 public void verifyEditDescriptionInEditDashboard() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkUpdateDashboardByClickOnSaveBtnInDashboard(), true);
	 }	
	 	 
	
	 @Test(priority=161)
	 public void verifyLogoutandLoginWithNewUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkLogoutandLoginWithNewUser(), true);
		
	 }	
	 
	 @Test(priority=162)
	 public void verifyLogoutandLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkLogoutAndLoginWithSU(), true);
	 }	
	
	 @Test(priority=178)
	 public void verifyCustomizationOptionsOfDashboards() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkCustomizationWithSelectingDashboard(), true);
	 } 
	 
	 @Test(priority=179)
	 public void verifyGraphCustomizationOptionsOfDashboards() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkGraphOptionsInCustomizeDashboard(), true);
	 } 
	 
	 @Test(priority=188)
	 public void verifySaveWithMandatoryFieldsInDefineGraph() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkSaveInTheDefineGraphInDefineGraph(), true);
	 } 
	 
	 @Test(priority=189)
	 public void verifyDragAndDropGraphNew() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkDragAndDropGraphAfterDeletingInCreatedashlet(), true);
	 }
	 
	 @Test(priority=191)
	 public void verifyUpdateGraphCreatedIsMessageDisplayedInDashlet() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkUpdateGraphCreatedInDashletsInDefineGraph(), true);
	 } 

	 @Test(priority=201)
	 public void verifyDashletGraphCreatedInDashboard() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkDashletGraphCreatedInDashboardAndDeleteDashlet(), true);
	 }
	 
	 @Test(priority=202)
	 public void verifyYesBtnInDeleteAfterClickOnNoGraphInDashlet() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkYesBtnInDeleteConfirmationMessagGraphCreatedInDashletsInDefineGraph(), true);
	 }
	 
	 @Test(priority=203)
	 public void checkAddingGraphDashlet() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkAddingGraphDashlet(), true);
	 }
	 
	 
	 @Test(priority=204)
	 public void verifyDashBoardDeleted() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.checkDashBoardDeleted(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=250)
	 public void TakeBackupForTransactions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 HPDP=new SmokeDashboardPage(getDriver());
		 Assert.assertEquals(SmokeDashboardPage.TakeBackupForTransactions(), true);
	 }
	 
	 
	 
	 
	 
	 
	 
}
