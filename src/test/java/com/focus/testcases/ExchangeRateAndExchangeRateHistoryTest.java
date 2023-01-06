package com.focus.testcases;

import java.awt.AWTException;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.focus.Pages.ExchangeRateAndExchangeRateHistoryPage;
import com.focus.Pages.ExchangeRateAndExchangeRateHistoryPage;
import com.focus.base.BaseEngine;

public class ExchangeRateAndExchangeRateHistoryTest extends BaseEngine
{
	ExchangeRateAndExchangeRateHistoryPage ER;
	
	 @Test(priority=140)
	 public void checkOpenExchangeRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkOpenExchangeRate(), true);
	 }
	 
	 @Test(priority=141)
	 public void checkDeleteOptionInExchangeRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkDeleteOptionInExchangeRate(), true);
	 }
	 
	 @Test(priority=142)
	 public void checkSavingCurrencyExchnageRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkSavingCurrencyExchnageRate(), true);
	 }
	 
	 @Test(priority=143)
	 public void checkOpenExchangeRateHistory() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkOpenExchangeRateHistory(), true);
	 }
	 
	 @Test(priority=144)
	 public void checkExchangeRateHistoryLoadButtonWithBaseCurrency() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkExchangeRateHistoryLoadButtonWithBaseCurrency(), true);
	 }
	 
	 @Test(priority=145)
	 public void checkUpdateExchangeRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkUpdateExchangeRate(), true);
	 }
	 
	 @Test(priority=146)
	 public void checkExchangeRateHistoryAfterUpdating() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkExchangeRateHistoryAfterUpdating(), true);
	 }
	 
	 
	 
	 
	 @Test(priority=147)
	 public void checkDateRangeOptionAsOnDateinExchangeRateHistory() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkDateRangeOptionAsOnDateinExchangeRateHistory(), true);
	 }
	 
	 @Test(priority=148)
	 public void checkDateOptionasCurrentMonthinExchangeRateHistory() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkDateOptionasCurrentMonthinExchangeRateHistory(), true);
	 }
	 
	 @Test(priority=149)
	 public void checkDateOptionasLastMonthinExchangeRateHistory() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkDateOptionasLastMonthinExchangeRateHistory(), true);
	 }
	 
	 
	 @Test(priority=150)
	 public void checkDateOptionasCurrentWeekinExchangeRateHistory() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkDateOptionasCurrentWeekinExchangeRateHistory(), true);
	 }
	 
	 
	 @Test(priority=151)
	 public void checkDateOptionasLastWeekinExchangeRateHistory() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkDateOptionasLastWeekinExchangeRateHistory(), true);
	 }
	 
	 
	 @Test(priority=152)
	 public void checkDateOptionasTodayinExchangeRateHistory() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkDateOptionasTodayinExchangeRateHistory(), true);
	 }
	 
	 
	 @Test(priority=153)
	 public void checkDateOptionasCurrentYearinExchangeRateHistory() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkDateOptionasCurrentYearinExchangeRateHistory(), true);
	 }
	 
	 
	 @Test(priority=154)
	 public void checkDateOptionasLastYearinExchangeRateHistory() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
		 ER=new ExchangeRateAndExchangeRateHistoryPage(getDriver());
		 Assert.assertEquals(ER.checkDateOptionasLastYearinExchangeRateHistory(), true);
	 }
	 
	 
	 
	 

}
