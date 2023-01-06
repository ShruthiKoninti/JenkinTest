package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.SanityFixedAssetsPage;
import com.focus.base.BaseEngine;

public class SanityFixedAssetsTest extends BaseEngine
{

	SanityFixedAssetsPage fsp;

	 //@Test(priority=101)
	 public void checkLoginAndRestoreOptionsCompanyAndLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkLoginAndRestoreOptionsCompanyAndLogin(), true);
	 }
	
	 @Test(priority=103)
	 public void checkFixedAssetsTabUnderPreference() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkFixedAssetsTabUnderPreference(), true);
	 }
	
	 @Test(priority=104)
	 public void checkSavingGroupAsBasedOnUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkSavingGroupAsBasedOnUsage(), true);
	 }
	
	 @Test(priority=105)
	 public void checkSavingGroupAsStraightLine() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkSavingGroupAsStraightLine(), true);
	 }
	 
	 @Test(priority=106)
	 public void checkDeletionGroupInFixedAssets() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkDeletionGroupInFixedAssets(), true);
	 }
	
	 
	 @Test(priority=107)
	 public void checkConditionsinFixedAssetsScreen() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkConditionsinFixedAssetsScreen(), true);
	 }
	 
	 @Test(priority=108)
	 public void checkSavedOptionsinFixedAssetsScreen() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkSavedOptionsinFixedAssetsScreen(), true);
	 }
	
	
	 @Test(priority=109)
	 public void checkPurchaseAccountListAfterEnableFixedAssetsOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkPurchaseAccountListAfterEnableFixedAssetsOption(), true);
	 }
	
	 @Test(priority=110)
	 public void checkCreationOfPurchaseAccountAndItemFromPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkCreationOfPurchaseAccountAndItemFromPurchaseVoucherVAT(), true);
	 }
	
	 
	 @Test(priority=112)
	 public void checkSavingVoucherinPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkSavingVoucherinPurchaseVoucher(), true);
	 }
	 
	 @Test(priority=113)
	 public void checkFixedAssetsPopUpScreenGenaralTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkFixedAssetsPopUpScreenGenaralTab(), true);
	 }
	 
	

	 @Test(priority=114)
	 public void checkFixedAssetsPopUpScreenDepreciationTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkFixedAssetsPopUpScreenDepreciationTab(), true);
	 }
	 
	
	 @Test(priority=116)
	 public void checkAddingFixedAssetInHeaderOFFixedAssetDepreciationVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkAddingFixedAssetInHeaderOFFixedAssetDepreciationVoucher(), true);
	 }
	  
	 
	 @Test(priority=117)
	 public void checkSavingVoucherInAssertUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkSavingVoucherInAssertUsage(), true);
	 }
	 
	 @Test(priority=118)
	 public void checkEditingInSavedVoucherInAssertUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkEditingInSavedVoucherInAssertUsage(), true);
	 }
	 
	 @Test(priority=119)
	 public void checkDeletingVoucherInAssertUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkDeletingVoucherInAssertUsage(), true);
	 }
	 
	 
	 @Test(priority=120)
	 public void checkSavingVoucherInAssertUsageAfterDeletion() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkSavingVoucherInAssertUsageAfterDeletion(), true);
	 }
	 
	 
	 @Test(priority=121)
	 public void checkSavingVoucherInAssertUsageWithSameAssertItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkSavingVoucherInAssertUsageWithSameAssertUsage(), true);
	 }
	
	 @Test(priority=122)
	 public void checkSavedVoucherInAssertUsage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkSavedVoucherInAssertUsage(), true);
	 }
	  
	 @Test(priority=123)
	 public void checkPostingInPostDepreciationEntries() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkPostingInPostDepreciationEntries(), true);
	 }
	  
	 
	 @Test(priority=124)
	 public void checkPostedVoucherInFixedAssetDepreciationVoucher() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkPostedVoucherInFixedAssetDepreciationVoucher(), true);
	 }
	       
	                             
	 @Test(priority=125)
	 public void checkAddingAccountUnderGroupInFixedAssets() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkAddingAccountUnderGroupInFixedAssets(), true);
	 }
	
	
  
	 @Test(priority=127)
	 public void checkPostingWithStraightLineInPostDepreciationEntries() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkPostingWithStraightLineInPostDepreciationEntries(), true);
	 }
	 
	 
	 @Test(priority=128)
	 public void checkPostedVoucherInFixedAssetDepreciationVoucherWithStraightLine() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkPostedVoucherInFixedAssetDepreciationVoucherWithStraightLine(), true);
	 }
	
	
	 @Test(priority=129)
	 public void checkAssetUsageReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkAssetUsageReport(), true);
	 }
	 
	 
	 @Test(priority=130)
	 public void checkFixedAssetReport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 fsp=new SanityFixedAssetsPage(getDriver());
		 Assert.assertEquals(fsp.checkFixedAssetReport(), true);
	 }
	 
	 
	 
	 
	 
	 
}
