package com.focus.testcases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;
import com.focus.Pages.ConfigureMasterPage;
import com.focus.base.BaseEngine;

public class ConfigureMasterTest extends BaseEngine
{
	
	ConfigureMasterPage cmp;
	
	@Test(priority = 101)
	public void checkconfigureMasters() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		 cmp=new ConfigureMasterPage(getDriver());
		  Assert.assertEquals(cmp.checkconfigureMasters(), true);
	}
	
	@Test(priority = 102)
	public void checkSavingMasterInConfigureMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		 cmp=new ConfigureMasterPage(getDriver());
		  Assert.assertEquals(cmp.checkSavingMasterInConfigureMaster(), true);
	}
	
	@Test(priority = 103)
	public void checkSavedMasterAndUpdatingTheMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		 cmp=new ConfigureMasterPage(getDriver());
		  Assert.assertEquals(cmp.checkSavedMasterAndUpdatingTheMaster(), true);
	}

	@Test(priority = 104)
	public void checkUpdatedMasterAndDeleteThatMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		 cmp=new ConfigureMasterPage(getDriver());
		  Assert.assertEquals(cmp.checkUpdatedMasterAndDeleteThatMaster(), true);
	}
	
	@Test(priority = 105)
	public void checkSavingMasterForRecordCreation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		 cmp=new ConfigureMasterPage(getDriver());
		  Assert.assertEquals(cmp.checkSavingMasterForRecordCreation(), true);
	}
	
	
	@Test(priority = 106)
	public void checkSavedMaserInMastersListAndSaveRecord() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		 cmp=new ConfigureMasterPage(getDriver());
		  Assert.assertEquals(cmp.checkSavedMaserInMastersListAndSaveRecord(), true);
	}
	
	
	
	
	@Test(priority = 108)
	public void checkAddingGroupMasterInConfigureMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		 cmp=new ConfigureMasterPage(getDriver());
		  Assert.assertEquals(cmp.checkAddingGroupMasterInConfigureMaster(), true);
	}
	
	@Test(priority = 109)
	public void checkAddingChildMasterInGroupMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		 cmp=new ConfigureMasterPage(getDriver());
		  Assert.assertEquals(cmp.checkAddingChildMasterInGroupMaster(), true);
	}
	
	@Test(priority = 110)
	public void checkDeletingSavedGroupmasterWithChildmaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		 cmp=new ConfigureMasterPage(getDriver());
		  Assert.assertEquals(cmp.checkDeletingSavedGroupmasterWithChildmaster(), true);
	}
	
	@Test(priority = 111)
	public void checkAddingGroupMasterForDeletion() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		 cmp=new ConfigureMasterPage(getDriver());
		  Assert.assertEquals(cmp.checkAddingGroupMasterForDeletion(), true);
	}
	
	
	
	
	
	
}
