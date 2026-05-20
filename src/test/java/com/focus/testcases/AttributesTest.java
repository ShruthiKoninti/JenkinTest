package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.AttributesPage;
import com.focus.base.BaseEngine;

public class AttributesTest extends BaseEngine{
	
	static AttributesPage AP;
	
	
	
	
	
	@Test(priority=101)
	public void checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkLogin(), true);
	}

	
	@Test(priority=102)
	public void checkConfigureTransactionMastersUnderSettings() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkConfigureTransactionMastersUnderSettings(), true);
	}

	@Test(priority=103)
	public static void checkItemAttributeValuesinMastersUnderSettingsConfigureTransactions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkItemAttributeValuesinMastersUnderSettingsConfigureTransactions(), true);
	}

	@Test(priority=104)
	public void checkCloseOptionInMastersScreenUnderSettingsMenu() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkCloseOptionInMastersScreenUnderSettingsMenu(), true);	
	}

	@Test(priority=105)
	public void CheckAttributesTabinItemMenu() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.CheckAttributesTabinItemMenu(), true);	
	}
	
	@Test(priority=106)
	public void checkItemsAddingforAttribute0inAttributesTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkItemsAddingforAttribute0inAttributesTab(), true);	
	}

	@Test(priority=107)
	public void checkItemsAddingforAttribute1inAttributesTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkItemsAddingforAttribute1inAttributesTab(), true);	
	}
	
	@Test(priority=108)
	public void checkItemsAddingforAttribute2inAttributesTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkItemsAddingforAttribute2inAttributesTab(), true);	
	}
	
	@Test(priority=109)
	public void checkItemsAddingforAttribute3inAttributesTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkItemsAddingforAttribute3inAttributesTab(), true);	
	}
	
	@Test(priority=110)
	public void checkSavedAttributesListofAttribute0inAttributesTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkSavedAttributesListofAttribute0inAttributesTab(), true);	
	}
	
	@Test(priority=111)
	public void checkSavedAttributesListofAttribute1inAttributesTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkSavedAttributesListofAttribute1inAttributesTab(), true);	
	}
	
	@Test(priority=112)
	public void checkSavedAttributesListofAttribute2inAttributesTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkSavedAttributesListofAttribute2inAttributesTab(), true);	
	}
	
	@Test(priority=113)
	public void checkSavedAttributesListofAttribute3inAttributesTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkSavedAttributesListofAttribute3inAttributesTab(), true);	
	}
	
	@Test(priority=114)
	public void checkSavingItemwithSelectedAttributesList() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkSavingItemwithSelectedAttributesList(), true);	
	}
	
	@Test(priority=115)
	public void checkSavedAttributeItemwithAllAttributeItemList() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkSavedAttributeItemwithAllAttributeItemList(), true);	
	}
	
	@Test(priority=116)
	public void checkControlTypeofDataTypeFieldinItemMasterCustomization() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkControlTypeofDataTypeFieldinItemMasterCustomization(), true);	
	}
	
	@Test(priority=117)
	public void checkLogoutAttributesPage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		AP=new AttributesPage(getDriver());
		Assert.assertEquals(AP.checkLogoutAttributesPage(), true);	
	}
	
}
