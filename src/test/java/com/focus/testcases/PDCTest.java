package com.focus.testcases;


import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.PDCPage;
import com.focus.base.BaseEngine;

public class PDCTest extends BaseEngine
{
	static PDCPage spdcp;

	@Test(priority=700398)
	public void checkPDCOptionsUnderSettings() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkPDCOptionsUnderSettings(),true);
	}

	@Test(priority=700399)
	public void checkConvertedMaturePDCScreenWithNoPDCExists() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkConvertedMaturePDCScreenWithNoPDCExists(),true);
	}

	@Test(priority=700404)
	public void checkSavingVoucherToPostDatedReceipts() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkSavingVoucherToPostDatedReceipts(),true);
	}


	@Test(priority=700414)
	public void checkSavedVoucherInPDRVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkSavedVoucherInPDRVAT(),true);
	}

	@Test(priority=700415)
	public void checkEditingVoucherAndSavingInPDRVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkEditingVoucherAndSavingInPDRVAT(),true);
	}

	@Test(priority=700416)
	public void checkSavedVouchersInPDCVoucherScreenWithOutPostOnDate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkSavedVouchersInPDCVoucherScreenWithOutPostOnDate(),true);
	}

	@Test(priority=700417)
	public void checkConvertingVouchersInPDCVoucherScreenWithOutPostOnDate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkConvertingVouchersInPDCVoucherScreenWithOutPostOnDate(),true);
	}

	@Test(priority=700418)
	public void checkAddExtraFieldPDCInHeaderFieldOFReceipts() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkAddExtraFieldPDCInHeaderFieldOFReceipts(),true);
	}


	@Test(priority=700419)
	public void checkCovertingVoucherInPDC() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkCovertingVoucherInPDC(),true);
	}

	@Test(priority=700424)
	public void checkConvertedPDCVoucherInReceipts() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkConvertedPDCVoucherInReceipts(),true);
	}

	@Test(priority=700425)
	public void checkDeleteOfConvertedVoucherinRecepits() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkDeleteOfConvertedVoucherinRecepits(),true);
	}
	

	@Test(priority=700426)
	public void checkClickOnConvertedPDCVouchersAfterDeletionOfConvertedVoucherInRecepits() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkClickOnConvertedPDCVouchersAfterDeletionOfConvertedVoucherInRecepits(),true);
	}

	@Test(priority=700427)
	public void checkUpdateinPDCUnderSettingMenuForEnableReversePostingInRecepits() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkUpdateinPDCUnderSettingMenuForEnableReversePostingInRecepits(),true);
	}

	

	@Test(priority=700428)
	public void checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkLogoutAndLoginWithSU(),true);
	}
	
	
	
	@Test(priority=700429)
	public void checkCreationOfNewCashAndBankAccountType() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkCreationOfNewCashAndBankAccountType(),true);
	}


	@Test(priority=700430)
	public void checkSavingVoucherInPostDatedReceiptsWithCustomerC() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkSavingVoucherInPostDatedReceiptsWithCustomerC(),true);
	}

	@Test(priority=700431)
	public void checkDeleteOptionInPDRVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkDeleteOptionInPDRVAT(),true);
	}


	@Test(priority=700435)
	public void checkCovertingVoucherInPDCWithReversePosting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkCovertingVoucherInPDCWithReversePosting(),true);
	}

	@Test(priority=700436)
	public void checkSavedCovertedVoucherInReceiptsAfterReversePosting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkSavedCovertedVoucherInReceiptsAfterReversePosting(),true);
	}
	
	@Test(priority=700437)
	public void checkUpdateinPDCUnderSettingMenuForDisableEnableReversePostingInRecepits() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkUpdateinPDCUnderSettingMenuForDisableEnableReversePostingInRecepits(),true);
	}


	//PDP VAT

	@Test(priority=700500)
	public void checkConvertMaturedPDCScreenForPDPVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkConvertMaturedPDCScreenForPDPVAT(),true);
	}


	@Test(priority=700501)
	public void CheckSavingVoucherInPDPVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.CheckSavingVoucherInPDPVAT(),true);
	}



	@Test(priority=700502)
	public void checkSavingNewVoucherInPDPVATOnFutureDateWithCurrencyAEDAndDepartmentINDIA() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkSavingNewVoucherInPDPVATOnFutureDateWithCurrencyAEDAndDepartmentINDIA(),true);
	}

	@Test(priority=700503)
	public void checkSavedVoucherInPDPVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkSavedVoucherInPDPVAT(),true);
	}

	@Test(priority=700504)
	public void checkEditingVoucherAndSavingInPDPVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkEditingVoucherAndSavingInPDPVAT(),true);
	}

	@Test(priority=700505)
	public void checkDeleteOptionInPDPVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkDeleteOptionInPDPVAT(),true);
	}

	@Test(priority=700506)
	public void CheckSavedVoucherInPDPVATInCovertedPDCScreen() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.CheckSavedVoucherInPDPVATInCovertedPDCScreen(),true);
	}

	@Test(priority=700516)
	public void checkConvertingVouchersPDPInPDCVoucherScreenWithOutPostOnDate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkConvertingVouchersPDPInPDCVoucherScreenWithOutPostOnDate(),true);
	}

	@Test(priority=700517)
	public void checkAddExtraFieldPDCInHeaderFieldOFPayments() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkAddExtraFieldPDCInHeaderFieldOFPayments(),true);
	}

	@Test(priority=700520)
	public void checkConvertingPDPVoucherInPDCScreenAfterExtraField() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkConvertingPDPVoucherInPDCScreenAfterExtraField(),true);
	}

	@Test(priority=700523) 
	public void checkConvertedPDCVoucherInPayments() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkConvertedPDCVoucherInPayments(),true);
	}



	@Test(priority=700524)
	public void checkUpdateinPDCUnderSettingMenuForEnableReversePostingInPayments() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkUpdateinPDCUnderSettingMenuForEnableReversePostingInPayments(),true);
	}

	@Test(priority=700525)
	public void checkSavingVoucherInPDPVATWithVendorB() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkSavingVoucherInPDPVATWithVendorB(),true);
	}


	@Test(priority=700526)
	public void checkCovertingPDPVoucherInPDCWithReversePosting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkCovertingPDPVoucherInPDCWithReversePosting(),true);
	}


	@Test(priority=700527)
	public void checkSavedCovertedVoucherInPaymentsAfterReversePosting() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		spdcp=new PDCPage(getDriver());
		Assert.assertEquals(spdcp.checkSavedCovertedVoucherInPaymentsAfterReversePosting(),true);
	}



}
