package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.DesignWorkFlowPage;
import com.focus.base.BaseEngine;

public class DesignWorkFlowTest extends BaseEngine
{
	DesignWorkFlowPage DWP;
	
	
	
	 @Test(priority=213)
	 public void CheckLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DWP=new DesignWorkFlowPage(getDriver());
		 Assert.assertEquals(DWP.CheckLogin(), true);
	 }

	 
	 @Test(priority=214)
	 public void checkSavingTestWorkFlow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DWP=new DesignWorkFlowPage(getDriver());
		 Assert.assertEquals(DWP.checkSavingTestWorkFlow(), true);
	 }
	 
	 @Test(priority=215)
	 public void checkEditingSavedTestWorkFlow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DWP=new DesignWorkFlowPage(getDriver());
		 Assert.assertEquals(DWP.checkEditingSavedTestWorkFlow(), true);
	 }
	 
	 @Test(priority=216)
	 public void checkDeletingTestWorkFlow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DWP=new DesignWorkFlowPage(getDriver());
		 Assert.assertEquals(DWP.checkDeletingTestWorkFlow(), true);
	 }
	 
	 
	 @Test(priority=217)
	 public void checkEditingPurchasesFlow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DWP=new DesignWorkFlowPage(getDriver());
		 Assert.assertEquals(DWP.checkEditingPurchasesFlow(), true);
	 }
	 
	 @Test(priority=219)
	 public void checkLogoutDesignWorkFlowPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 DWP=new DesignWorkFlowPage(getDriver());
		 Assert.assertEquals(DWP.checkLogoutDesignWorkFlowPage(), true);
	 }
	 
	 
	 
}
