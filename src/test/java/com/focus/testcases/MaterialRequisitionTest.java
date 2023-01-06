package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.MaterialRequisitionPage;
import com.focus.Pages.MaterialRequisitionPage;
import com.focus.base.BaseEngine;

public class MaterialRequisitionTest extends BaseEngine
{
	MaterialRequisitionPage MRP;
	
	
	 @Test(priority=245)
	 public void checkSavingMaterialRequisitionVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MRP=new MaterialRequisitionPage(getDriver());
		 Assert.assertEquals(MRP.checkSavingMaterialRequisitionVoucher1(), true);   
	 }
	 
	 @Test(priority=246)
	 public void checkVerifyingAndDeletingMaterialRequisitionVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MRP=new MaterialRequisitionPage(getDriver());
		 Assert.assertEquals(MRP.checkVerifyingAndDeletingMaterialRequisitionVoucher1(), true);
	 }
	 
	 @Test(priority=247)
	 public void checkReSavingMaterialRequisitionVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MRP=new MaterialRequisitionPage(getDriver());
		 Assert.assertEquals(MRP.checkReSavingMaterialRequisitionVoucher1(), true);
	 }
	 
	 @Test(priority=248)
	 public void checkVerifyingAndUpdatingMaterialRequisitionVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MRP=new MaterialRequisitionPage(getDriver());
		 Assert.assertEquals(MRP.checkVerifyingAndUpdatingMaterialRequisitionVoucher1(), true);
	 }
	 
	 @Test(priority=249)
	 public void checkVerifyingMaterialRequisitionVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MRP=new MaterialRequisitionPage(getDriver());
		 Assert.assertEquals(MRP.checkVerifyingMaterialRequisitionVoucher1(), true);
	 }
	 
	 @Test(priority=250)
	 public void checkSavingMaterialRequisitionVoucher2AndSuspendingFromEntryPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MRP=new MaterialRequisitionPage(getDriver());
		 Assert.assertEquals(MRP.checkSavingMaterialRequisitionVoucher2AndSuspendingFromEntryPage(), true);
	 }
	 
	 @Test(priority=251)
	 public void checkResavingMaterialRequisitionVoucher2AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MRP=new MaterialRequisitionPage(getDriver());
		 Assert.assertEquals(MRP.checkResavingMaterialRequisitionVoucher2AndSuspendingFromHomepage(), true);
	 }
	 
	 @Test(priority=252)
	 public void checkDeleteOptionInMaterialRequisitionHomePage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MRP=new MaterialRequisitionPage(getDriver());
		 Assert.assertEquals(MRP.checkDeleteOptionInMaterialRequisitionHomePage(), true);
	 }
	 
	 @Test(priority=253)
	 public void checkVerifingDetailsOfSavedMaterialRequisitionVouchersInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 MRP=new MaterialRequisitionPage(getDriver());
		 Assert.assertEquals(MRP.checkVerifingDetailsOfSavedMaterialRequisitionVouchersInHomepage(), true);
	 }
	 
	 
	 

}
