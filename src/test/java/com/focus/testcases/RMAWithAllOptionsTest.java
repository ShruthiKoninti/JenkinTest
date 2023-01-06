package com.focus.testcases;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.RMAWithAllOptionsPage;
import com.focus.base.BaseEngine;

public class RMAWithAllOptionsTest extends BaseEngine
{
	
	RMAWithAllOptionsPage rma;
	 
	
	 @Test(priority=60000)
	 public void checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkLogin(), true);
	 }
	
	 ///@Test(priority=60000)
	 public void checkLoginAndRestoreOptionsCompanyAndLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkLoginAndRestoreOptionsCompanyAndLogin(), true);
	 }
	 
	 // RMA Options Starts Here

	 @Test(priority=60106)
	 public void checkConfigureSelectAndUpdateRMAAlwaysNumericChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkConfigureSelectAndUpdateRMAAlwaysNumericChkBox(), true);
	 }
	
	 @Test(priority=60107)
	 public void checkLogoutAndLoginAfterUpdatingRMAAlwaysNumeric() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkLogoutAndLoginAfterUpdatingRMAAlwaysNumeric(), true);
	 }
	 
	 @Test(priority=60108)
	 public void checkPurchaseVoucherVATVoucher1RMAAsAlphabets() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkPurchaseVoucherVATVoucher1RMAAsAlphabets(), true);
	 }
	 
	 @Test(priority=60109)
	 public void checkPurchaseVoucherVATVoucher1RMAAsNumeric() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkPurchaseVoucherVATVoucher1RMAAsNumeric(), true);
	 }
	 
	 @Test(priority=60110)
	 public void checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantityInPurchaseVoucherVAT1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantityInPurchaseVoucherVAT1(), true);
	 }
	 
	 @Test(priority=60111)
	 public void checkEraseAllTransactionsAfterCompletingRMAAlwaysNumeric() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkEraseAllTransactionsAfterCompletingRMAAlwaysNumeric(), true);
	 }
	 
	 
	 @Test(priority=60112)
	 public void checkConfigureRMAUniqueAcrossAllItemsChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkConfigureRMAUniqueAcrossAllItemsChkBox(), true);
	 }
	 
	 @Test(priority=60113)
	 public void checkLogoutAndLoginAfterRMAUniqueAcrossAllItems() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkLogoutAndLoginAfterRMAUniqueAcrossAllItems(), true);
	 }
	 
	 @Test(priority=60114)
	 public void checkSavingPurchaseVoucherVATVoucher1() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingPurchaseVoucherVATVoucher1(), true);
	 }
	 
	 @Test(priority=60115)//no alert message coming as error message
	 public void checkSavingPurchaseVoucherVATVoucher2WithRMAUnique() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingPurchaseVoucherVATVoucher2WithRMAUnique(), true);
	 }
	 
	 @Test(priority=60116)
	 public void checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantity(), true);
	 }
	 
	 @Test(priority=60117)
	 public void checkEraseAllTransactionsAfterCompletingRMANumbersAreUniqueAcrossAllItems() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkEraseAllTransactionsAfterCompletingRMANumbersAreUniqueAcrossAllItems(), true);
	 }
	 
	 
	 
	 @Test(priority=60118)
	 public void checkInputRMABeforeQuantityChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkInputRMABeforeQuantityChkBox(), true);
	 }
	 
	 @Test(priority=60119)
	 public void checkLogoutAndLoginAfterRMABeforeQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkLogoutAndLoginAfterRMABeforeQuantity(), true);
	 }
	 
	 @Test(priority=60120)
	 public void checkSavingPurchaseVoucherVATVoucher1RMABeforeQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingPurchaseVoucherVATVoucher1RMABeforeQuantity(), true);
	 }
	 
	 @Test(priority=60121)
	 public void checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantityWithRMABeforeQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantityWithRMABeforeQuantity(), true);
	 }
	 
	 @Test(priority=60122)
	 public void checkEraseAllTransactionsAfterCompletingRMANumberBeforeQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkEraseAllTransactionsAfterCompletingRMANumberBeforeQuantity(), true);
	 }
	 
	 
	
	 @Test(priority=60123)
	 public void checkInputRMANoCannotBeReuseChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkInputRMANoCannotBeReuseChkBox(), true);
	 }
	 
	 @Test(priority=60124)
	 public void checkLogoutAndLoginAfterRMACantReuse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkLogoutAndLoginAfterRMACantReuse(), true);
	 }
	 
	 @Test(priority=60125)
	 public void checkSavingPurchaseVoucherVATVoucher1WithRMACantReuse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingPurchaseVoucherVATVoucher1WithRMACantReuse(), true);
	 }
	 
	 @Test(priority=60126)
	 public void checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantityWithRMACantReuse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingSalesInvoiceVATVoucher1ConsumingAllAvailableQuantityWithRMACantReuse(), true);
	 }
	 
	 @Test(priority=60127)
	 public void checkSavingPurchaseVoucherVATVoucher2WithAlreadyUsedRMA() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingPurchaseVoucherVATVoucher2WithAlreadyUsedRMA(), true);
	 }
	 
	 @Test(priority=60128)
	 public void checkEraseAllTransactionsAfterCompletingRMACantReuse() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkEraseAllTransactionsAfterCompletingRMACantReuse(), true);
	 }
	 
	//////
	 
	 
	 @Test(priority=60129)
	 public void checkPrefixVoucherNumberToTheSerialNunberInInwardsChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkPrefixVoucherNumberToTheSerialNunberInInwardsChkBox(), true);
	 }
	 
	 @Test(priority=60130)
	 public void checkLogoutAndLoginAfterRMAPrefixVoucherNo() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkLogoutAndLoginAfterRMAPrefixVoucherNo(), true);
	 }
	 
	 @Test(priority=60131)
	 public void checkSavingPurchaseVoucherVATVoucher1RMAPrefixVoucherNo() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingPurchaseVoucherVATVoucher1RMAPrefixVoucherNo(), true);
	 }
	 
	 @Test(priority=60132)
	 public void checkSavingSalesInvoiceVATVoucher1ConsumeAllAvailableQuantityRMAPrefixVoucherNo() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingSalesInvoiceVATVoucher1ConsumeAllAvailableQuantityRMAPrefixVoucherNo(), true);
	 }
	 
	 @Test(priority=60133)
	 public void checkEraseAllTransactionsAfterCompletingPrefixVoucherNumberToTheSerialNumberInInwards() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkEraseAllTransactionsAfterCompletingPrefixVoucherNumberToTheSerialNumberInInwards(), true);
	 }
	 
	 /////
	
	 @Test(priority=60134)
	 public void checkRMABasedAlternateQuantityChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkRMABasedAlternateQuantityChkBox(), true);
	 }
	 
	 @Test(priority=60135)
	 public void checkSelectingMaintainInventoryInAlternateUnitCheckBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSelectingMaintainInventoryInAlternateUnitCheckBox(), true);
	 }
	 
	 @Test(priority=60136)
	 public void checkLogoutAndLoginAfterRMABasedAlternateQty() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkLogoutAndLoginAfterRMABasedAlternateQty(), true);
	 }
	 
	 @Test(priority=60137)
	 public void checkSavingPurchaseVoucherVATVoucher1WithRMAAlternateQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingPurchaseVoucherVATVoucher1WithRMAAlternateQuantity(), true);
	 }
	 
	 @Test(priority=60138)
	 public void checkSavingSalesInvoiceVATVoucher1WithAlternateQuantityConsuming12Quantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingSalesInvoiceVATVoucher1WithAlternateQuantityConsuming12Quantity(), true);
	 }
	 
	 @Test(priority=60139)
	 public void checkSavingSalesInvoiceVATVoucher2WithAlternateQuantityConsuming12Quantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingSalesInvoiceVATVoucher2WithAlternateQuantityConsuming12Quantity(), true);
	 }
	 
	 @Test(priority=60140)
	 public void checkUnCheckRMABasedAlternateQuantityChkBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkUnCheckRMABasedAlternateQuantityChkBox(), true);
	 }
	 
	 @Test(priority=60141)
	 public void checkUnCheckMaintainInventoryInAlternateUnitCheckBox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkUnCheckMaintainInventoryInAlternateUnitCheckBox(), true);
	 }
	 
	 @Test(priority=60142)
	 public void checkEraseAllTransactionsAfterCompletingRMABasedAlternateQuantity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkEraseAllTransactionsAfterCompletingRMABasedAlternateQuantity(), true);
	 }
	 
	 ////uptoexcel
	 @Test(priority=60151)
	 public void checkLoginAndEnableReservationIsMandatoryChkbox() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkLoginAndEnableReservationIsMandatoryChkbox(), true);
	 }
	 
	 @Test(priority=60152)
	 public void checkSavingPurchaseVoucherVATWithBRItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingPurchaseVoucherVATWithBRItem(), true);
	 }
	 
	 @Test(priority=60152)
	 public void checkSavingRMAItemInPVVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingRMAItemInPVVAT(), true);
	 }
	 
	 @Test(priority=60153)
	 public void checkClickOnSaveButtonInsalesaOrderWithoutreservaion() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkClickOnSaveButtonInsalesaOrderWithoutreservaion(), true);
	 }
	 
	 @Test(priority=60154)
	 public void checkMandatoryoptionWithWACogsItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkMandatoryoptionWithWACogsItem(), true);
	 }
	 
	 
	 @Test(priority=60155)
	 public void checkEnableDoNotAllowPartialreservation() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkEnableDoNotAllowPartialreservation(), true);
	 }
	 
	 @Test(priority=60156)
	 public void checkInputMoreQtyInReservationpopupLessThanQtyInAQ() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkInputMoreQtyInReservationpopupLessThanQtyInAQ(), true);
	 }
	 
	 
	 
	 
	 
	 
	 ////////////
	
	 
	 
	 
	 /*
	 @Test(priority=60153)
	 public void checkSavingSalesInvoiceVATWithBatch() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingSalesInvoiceVATWithBatch(), true);
	 }
	 
	 @Test(priority=60154)
	 public void checkChangingOptionAsStopInPreferenceUnderInventory() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkChangingOptionAsStopInPreferenceUnderInventory(), true);
	 }
	 
	 @Test(priority=60155)
	 public void checkResavingSalesInvoiceVATAfterChangigInSettingWithLogoutAndLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkResavingSalesInvoiceVATAfterChangigInSettingWithLogoutAndLogin(), true);
	 }
	 
	
	 @Test(priority=60156)
	 public void checkEraseAllDataAndSavingPurchaseVoucherAfterEnableReservationMadatory() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkEraseAllDataAndSavingPurchaseVoucherAfterEnableReservationMadatory(), true);
	 }
	  
	
	 @Test(priority=60157)
	 public void checkSavingSalesInvoiceDirectlyAfterEnableReserVationManadatory() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		 rma=new RMAWithAllOptionsPage(getDriver());
		 Assert.assertEquals(rma.checkSavingSalesInvoiceDirectlyAfterEnableReserVationManadatory(), true);
	 }
	 
	*/
	 
	
	
	
	
	
	
	
}
