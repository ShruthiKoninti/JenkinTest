package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.TransactionAuthorizationPage;
import com.focus.base.BaseEngine;

public class TransactionAuthorizationTest extends BaseEngine
{
	TransactionAuthorizationPage TP;
	
	
	 @Test(priority=208)
	 public void checkSavingTransactionAuthorizationInPurchaseReturns() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 TP=new TransactionAuthorizationPage(getDriver());
		 Assert.assertEquals(TP.checkSavingTransactionAuthorizationInPurchaseReturns(), true);
	 }
	 
	 @Test(priority=209)
	 public void checkEditingTransactionAuthorizationInPurchaseReturns() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 TP=new TransactionAuthorizationPage(getDriver());
		 Assert.assertEquals(TP.checkEditingTransactionAuthorizationInPurchaseReturns(), true);
	 }
	 
	 @Test(priority=210)
	 public void checkDeletingTransactionAuthorizationInPurchaseReturns() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 TP=new TransactionAuthorizationPage(getDriver());
		 Assert.assertEquals(TP.checkDeletingTransactionAuthorizationInPurchaseReturns(), true);
	 } 
	 
}
