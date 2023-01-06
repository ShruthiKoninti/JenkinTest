package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.DebitNotesPage;
import com.focus.Pages.TransactionImportPage;
import com.focus.base.BaseEngine;

public class TransactionImportTest extends BaseEngine
{

	TransactionImportPage TIP;
	
	@Test(priority=1000201)
	public void checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		TIP=new TransactionImportPage(getDriver());
	 	Assert.assertEquals(TIP.checkLogin(), true);
	}
	
	@Test(priority=1000202)
	public void checkImportingSupplierDebitNoteFinAndInvInTransationImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		TIP=new TransactionImportPage(getDriver());
	 	Assert.assertEquals(TIP.checkImportingSupplierDebitNoteFinAndInvInTransationImport(), true);
	}
	
	
	
	
	
	//checkClickonImportOptionInTransationImport
	
	

}
