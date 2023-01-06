package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.DocumentCustomizationPage;
import com.focus.Pages.DocumentNumberingPage;
import com.focus.base.BaseEngine;

public class DocumentNumberingTest  extends BaseEngine{
	
	DocumentNumberingPage dnp;
	
	
	 @Test(priority=150)
	 public void checkNavigateToDocumentNumberingTabInReceiptsVATVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 dnp=new DocumentNumberingPage(getDriver());
		 Assert.assertEquals(dnp.checkNavigateToDocumentNumberingTabInReceiptsVATVoucher(), true);
	 }

	 
	 
	 
	 @Test(priority=151)
	 public void checkSavingRecepitsVATVocher1withDocumentNumber() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 dnp=new DocumentNumberingPage(getDriver());
		 Assert.assertEquals(dnp.checkSavingRecepitsVATVocher1withDocumentNumber(), true);
	 }
	 
	
	 
	 
	 
	 @Test(priority=152)
	 public void checkSavingRecepitsVATVocher2withChangedDocumentNumber() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 dnp=new DocumentNumberingPage(getDriver());
		 Assert.assertEquals(dnp.checkSavingRecepitsVATVocher2withChangedDocumentNumber(), true);
	 }
	 
	 
		 
	 
	 
	 @Test(priority=153)
	 public void checkSavingRecepitsVATVocher3withChangedDocumentNumber() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 dnp=new DocumentNumberingPage(getDriver());
		 Assert.assertEquals(dnp.checkSavingRecepitsVATVocher3withChangedDocumentNumber(), true);
	 }

}
