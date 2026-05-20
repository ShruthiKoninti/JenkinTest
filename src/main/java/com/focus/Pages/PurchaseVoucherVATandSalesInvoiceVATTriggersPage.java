package com.focus.Pages;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.focus.base.BaseEngine;
import com.testautomationguru.utility.PDFUtil;

public class PurchaseVoucherVATandSalesInvoiceVATTriggersPage extends BaseEngine{
	
	
	
	@FindBy(xpath="//*[@id='id_search_menu']/input")
    private static WebElement serachMenuTextHomePage;
	
	@FindBy(xpath="(//li/span)[1]")
    private static WebElement searchMenuTextClick;
	
	@FindBy(xpath="//div[@id='docCustomization']")
	private static WebElement  documentCustomization45Btn;
	
	  @FindBy(xpath="//a[@id='769']")
	  private static WebElement  purchaseVouchersVATBtn;
	
	  @FindBy(xpath="//span[@class='dropdown icon-menuicon1 icon-font6 theme_color-inverse pull-right']")
	  private static WebElement  toogleExpandBtn;
	  
	  @FindBy(xpath="(//*[contains(text(),'Triggers')])[1]")
	  private static WebElement  triggersBtn;
	  
	 @FindBy(xpath="//input[@id='triggers_chksuspend']")
	private static WebElement  triggersSuspendedChkboxs;
	 
	 @FindBy(xpath="//input[@id='triggers_chksuspend']/following-sibling::span")
		private static WebElement  triggersSuspendedChkboxsSelected;
	  
	  @FindBy(xpath="//input[@id='triggers_existTriggers']")
	  private static WebElement  editLayoutTriggerName;
	  
	  @FindBy(xpath="//td[@id='triggers_docTable_col_1-1']")
      private static WebElement  trigger1stRow1stCol;
	  
	  @FindBy(xpath="//td[@id='triggers_docTable_col_1-2']")
      private static WebElement  trigger1stRow2ndCol;
	  
	  @FindBy(xpath="//select[@id='triggers_docTableFieldDropDown']")
      private static WebElement  triggerSelectFieldDropdown;
	  
	  @FindBy(xpath="//select[@id='triggers_docTableMapDropDown']")
      private static WebElement  triggerSelectMapDropdown;
	  
		@FindBy(xpath="//*[@id='trigger_Controlbuttons']/a[2]/i")
		private static WebElement  editLayoutTriggersSaveBtn;
	  
		@FindBy(xpath="//a[@id='updateButton']/i")
		private static WebElement  updateBtn;
	  
		 @FindBy(xpath="//*[@id='triggers_chkraiseNewDoc']")
		  private static WebElement editLayoutTriggerRaiseDoc;  	
		  
		 @FindBy(xpath="//*[@id='triggers_chkraiseNewDoc']/following-sibling::span")
		  private static WebElement editLayoutTriggerRaiseDocSelected;  	
		  
		  @FindBy(xpath="//select[@id='triggers_allVoucherDropDown']")
		  private static WebElement selectRaiseaDocument;  	
		  
		  
		  @FindBy(xpath="//select[@id='triggers_allVoucherDropDown']")
		  private static WebElement editLayoutTriggersAllVouchers;  	
		  
		  
			@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[1]/select")
			private static WebElement  triggersConjuctionDrpdwn;
			
			@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[2]/input")
			private static WebElement  triggersSelectFieldTxt;
			
			@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[3]/select")
			private static WebElement  triggersSelectOperdrpdwn;
			
			@FindBy(xpath="//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[4]/select")
			private static WebElement  triggersCompareWithdrpdwn;
			
		//	@FindBy(xpath="//input[@id='advancefilter_master_17_1_']")
			@FindBy(xpath="//*[@id='advancefilter_master_17_1_']")
			private static WebElement  triggersValueTxt;
			
			
			
			
			@FindBy(xpath="(//*[@id='4']/span)[1]")
			private static WebElement  triggerVendorACExpansionBtn;
		  
			@FindBy(xpath="(//*[@id='5002'])[2]")
			private static WebElement  triggerVendorAcName;
			
			@FindBy(xpath="(//*[@id='5002'])[1]")
			private static WebElement  triggerCustomerAcName;
			
			@FindBy(xpath="//*[@id='4']")
			private static WebElement  triggerVendorAc;
			
			@FindBy(xpath="//*[@id='5002']")
			private static WebElement  triggerNameUnderVendor;
		  
			//@FindBy(xpath="//td[@id='advancefilter_master_17_1__input_image']/span")
			@FindBy(xpath="//*[@id='0_1_AdvanceFilter_']/table/tbody/tr/td[5]")
			private static WebElement  triggersValueToEnterTxt;
			
			@FindBy(xpath="//td[@id='advancefilter_master_17_1__input_image']/span")
			private static WebElement  triggersValueTxt1;
			
			
			@FindBy(xpath="//*[@id='id_Pick']/a/span")
			private static WebElement pickBtn;

			@FindBy(xpath="//input[@id='txtNewReference']")
			private static WebElement newReferenceTxt;

			@FindBy(xpath="//*[@id='id_Ok']/a/span")
			private static WebElement Bill_OkBtn;
			
			public boolean checkEraseAllTransactions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
				homeMenu.click();
				Thread.sleep(4000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
				dataMangementMenu.click();
				Thread.sleep(4000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAll));
				eraseAll.click();
				Thread.sleep(6000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseTranscationsRadio));
				eraseTranscationsRadio.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAllOkBtn));
				eraseAllOkBtn.click();
				
				if(getIsAlertPresent())
				{
					getWaitForAlert();
					
					getAlert().accept();
				}
				Thread.sleep(2000);
				String expMessage = "Data deleted successfully.";
				
				String actMessage = checkValidationMessage(expMessage);
				
				System.out.println("************************************* checkEraseAllTransactionsAfterCompletingAllValidations  *********************************");
				System.out.println("Message  :  "+actMessage +" Value Expected : "+expMessage);
				
				if (actMessage.equalsIgnoreCase(expMessage))
				{
					System.out.println("Voucher Saved");
					return true;
				}
				else
				{
					System.out.println("Voucher Not Saved");
					return false;
				}
			}
			 @FindBy (xpath="//table[@id='editScreen_bodyTable']/tbody/tr[5]/td/span[2]")
			 private static WebElement editScreenBodyRow5Delete;				
			
	public boolean checkDeleteVATFieldinEditScreenofPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Configure Transactions");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();	  
		
		Thread.sleep(2000);
		
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
		purchaseVouchersVATBtn.click();
						 	    	
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyTab));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyRow5Delete));
		editScreenBodyRow5Delete.click();
		
		getWaitForAlert();
		getAlert().accept();
		
		System.out.println("*******************************************checkDeleteFieldInEditScreenBodyInTrialPurchaseUpdate*******************************************");
			
		String expMessage = "Data deleted successfully";
		String actMessage = checkValidationMessage(expMessage);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		
		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);
			
		if(actMessage.equalsIgnoreCase(expMessage) )
		{	
			return  true;
		}
		else
		{
			return  false;
		}
		
		
		
	}
			
	@FindBy(xpath = "(//*[@id='536870916']/span)[1]")
	private static WebElement triggerWarehouseExpansionBtn;		
			
			
	  
	public boolean checkSaveTriggerwithDiscountasMapFieldinPVVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
							 	    	
		Thread.sleep(6000);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();
					
		Thread.sleep(6000);
		
		checkDeletingTriggerInPurchaseVoucherVAT();
		//Thread.sleep(15000);
		
		new WebDriverWait(getDriver(), 250).until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		
		//getAction().moveToElement(editLayoutTriggerName).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Map Trigger");
		Thread.sleep(4000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(8000);
		
		triggersConjuctionDrpdwn.click();
        Thread.sleep(2000);		
		triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);			
		triggersConjuctionDrpdwn.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
		triggersSelectFieldTxt.click();
					
		Thread.sleep(3000);
		
		getAction().moveToElement(triggerVendorACExpansionBtn).build().perform();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorACExpansionBtn));
		triggerVendorACExpansionBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorAcName));
		triggerVendorAcName.click();
		
		Thread.sleep(2000);
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
		triggersSelectOperdrpdwn.click();
		triggersSelectOperdrpdwn.sendKeys("Equal");
		
		
		Thread.sleep(5000);
		triggersSelectOperdrpdwn.sendKeys(Keys.TAB);
		
		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
		triggersCompareWithdrpdwn.click();
		triggersCompareWithdrpdwn.sendKeys("value");
		//triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
		Thread.sleep(2000);
		triggersCompareWithdrpdwn.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		
		triggersValueTxt.sendKeys("Vendor B");
		Thread.sleep(1500);
		triggersValueTxt.sendKeys(Keys.TAB);

		
		Thread.sleep(4000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
		if(editLayoutTriggerRaiseDoc.isSelected()==false)
		{
			editLayoutTriggerRaiseDocSelected.click();
		}
		Thread.sleep(1000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
		triggerTabAlertBtn.click();
		
		Thread.sleep(2000);*/
		
		ScrollIntoView(editLayoutTriggersAllVouchers);
		Thread.sleep(1000);
							
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
		Select s=new Select(editLayoutTriggersAllVouchers);
		s.selectByVisibleText("Job orders");
		Thread.sleep(3000);
		
		
		getAction().moveToElement(trigger1stRow1stCol).build().perform();
		
		
		ClickUsingJs(trigger1stRow1stCol);
		Thread.sleep(2000);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectFieldDropdown));
	
		Thread.sleep(1000);
		Select s1=new Select(triggerSelectFieldDropdown);
		s1.selectByVisibleText("Discount");
		Thread.sleep(3000);
		triggerSelectFieldDropdown.sendKeys(Keys.TAB);
		
		
		ClickUsingJs(trigger1stRow2ndCol);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectMapDropdown));
		Select s2=new Select(triggerSelectMapDropdown);
		s2.selectByVisibleText("Discount");
		Thread.sleep(2000);
		
		
		getAction().moveToElement(editLayoutTriggersSaveBtn).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		Thread.sleep(2000);
		
		String expSaveMessage = "Data saved successfully";
		String actSaveMessage = checkValidationMessage(expSaveMessage);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		
		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);
			
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Settings_closeBtn));
		Settings_closeBtn.click();*/
			
		if( actSaveMessage.equalsIgnoreCase(expSaveMessage))
		{
			System.out.println("Test Pass : Trigger  Added Successfully");
			
			return true;
		}
		else
		{
			System.out.println("Test Fail : Trigger NOT Added Successfully");
			
			return false;
		}

		
		
	}
	
	public static String checkValidationMessage(String ExpMessage) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	        try
	        {
	                getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
	                String actErrorMessage=errorMessage.getText();
	                String expErrorMessage=ExpMessage;

	                try
	                {
	                
		                getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
		                errorMessageCloseBtn.click();

		                System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
		                
		                return actErrorMessage;
	                }
	                catch(Exception ee)
	                {
	                	
	                	System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
	                	
	                	return actErrorMessage;
	                }
	        }
	        catch(Exception e)
	        {
	                System.err.println("Error Message NOT Found or NOT Clickable");
	                System.err.println(e.getMessage());
	                
	                String Exception=e.getMessage();
	            
	                return Exception;
	        }
	}
	
	
	
	@FindBy(xpath = "//input[@id='id_body_33554527']")
	public static WebElement enter_Discount;
	
	
	
	
	
	public boolean checkSavingPurchaseVoucherVATVoucher1withDiscountasMapField() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();

		Thread.sleep(6000);
		new WebDriverWait(getDriver(), 300).until(ExpectedConditions.visibilityOf(newBtn));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();

		vendorAccountTxt.sendKeys(Keys.SPACE);	

		ArrayList<String>  actVendorAccountList= new ArrayList<String>(); 

		int count=vendorAccountListCount.size();

		System.err.println(count);

		for(int i=0 ; i < count ;i++)
		{
			String data=vendorAccountListCount.get(i).getText();

			if (data.equalsIgnoreCase("Vendor B")) 
			{
				vendorAccountListCount.get(i).click();
				break;
			}
		}

		vendorAccountTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(Keys.SPACE);

		int currencycount=currencyListCount.size();

		System.err.println(currencycount);

		for(int i=0 ; i < currencycount ;i++)
		{
			String data=currencyListCount.get(i).getText();

			if (data.equalsIgnoreCase("INR")) 
			{
				currencyListCount.get(i).click();
				break;
			}
		}

		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);	

		ArrayList<String> actDepartmentList= new ArrayList<String>(); 

		int departmentCount=departmentListCount.size();

		System.err.println(departmentCount);

		for(int i=0 ; i < departmentCount ;i++)
		{
			String data=departmentListCount.get(i).getText();

			if (data.equalsIgnoreCase("INDIA")) 
			{
				departmentListCount.get(i).click();
				break;
			}
		}

		departmentTxt.sendKeys(Keys.TAB);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.clear();
		placeOFSupplyTxt.sendKeys("Abu Dhabi");
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys("DUBAI");
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();

		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount=warehouseBodyComboList.size();

		for(int i=0 ; i < warehousecount ;i++)
		{
			String data=warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("Mumbai")) 
			{
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount=pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount; i++) 
		{
			String Item=pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("BR COGS ITEM")) 
			{
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("2");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));
		select1stRow_13thColumn.click();*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
		enter_Discount.click();
		enter_Discount.clear();
		enter_Discount.sendKeys("10");
		enter_Discount.sendKeys(Keys.TAB);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.sendKeys(Keys.TAB);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys("PVATR1#V1");
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 2); 

		String next2DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "+next2DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next2DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		Thread.sleep(2000);

				
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": SU/IND/TEXT1";
		
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) )
		{
			return true;
		} 
		else 
		{
			return false;
		}
	
	}
	
	
	public static  ArrayList<String> checkNetValueofVoucher1inPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(1000);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("SU/IND/TEXT1");
		Thread.sleep(4000);
		documentNumberTxt.sendKeys(Keys.TAB);
		Thread.sleep(8000);

		checkValidationMessage("Voucher loaded successfully");

		ArrayList<String> DiscNetArry = new ArrayList<String>();
		String actDiscount = select1stRow_13thColumn.getText();
		
		
		String actNetValue= voucherFooterNetValue.getText();
		
		DiscNetArry.add(actDiscount);
		DiscNetArry.add(actNetValue);
		
		
		return DiscNetArry;
		
	}
	
	
	
	
	
	
	
	@FindBy(xpath="//label[@id='id_transactionentry_footer_panel_summary_value_net']")
	public static WebElement voucherFooterNetValue;
	
	
	@FindBy(xpath="//*[@id='id_transaction_homescreen_edit']")
	public static WebElement editBtn;
	
	public boolean checkRaisedVoucherAfterSavingTriggerVoucher1() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		ArrayList<String> DiscNetArry = new ArrayList<String>();
		DiscNetArry=checkNetValueofVoucher1inPurchaseVoucherVAT();
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Job Order");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();

		Thread.sleep(6000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data ="1";
			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(6000);
		
		
		
		String actDiscountValue=select1stRow_6thColumn.getText();
		String expDiscountValue=DiscNetArry.get(0);
		
		String actNetValue=voucherFooterNetValue.getText();
		
		String expNetValue=DiscNetArry.get(1);
		
		System.out.println("Actual Discount    : "+actDiscountValue);
		System.out.println("Expected Discount  : "+expDiscountValue);
		
		System.out.println("Actual Net   : "+actNetValue);
		System.out.println("ExpectedNet  : "+expNetValue);
		
		
		
		if(actDiscountValue.equalsIgnoreCase(expDiscountValue)&&actNetValue.equalsIgnoreCase(expNetValue))
		{
			return true;
		}
		else
		{
			return false;
		}
					
	}
	
	@FindBy (xpath="//*[@id='6']/span")
    public static WebElement securityMenu;
	
	@FindBy(xpath="//span[contains(text(),'Create User')]")
	public static WebElement createUserMenu;
	
	@FindBy(xpath="//*[@id='createUserBtn']")
	public static WebElement createUserIcon;
	
	@FindBy(xpath="//input[@id='loginName']")
	public static WebElement createUserLoginNameTxt;
	
	@FindBy(xpath="//select[@id='ERPRoles']")
	public static WebElement createUserERPRoleDropDown;

	@FindBy(xpath="//select[@id='passwordPolicy']")
	public static WebElement createUserPasswordPolicyDropdown;

	@FindBy(xpath="//input[@id='password']")
	public static WebElement createUserPasswordTxt;

	@FindBy(xpath="//input[@id='confirmpassword']")
	public static WebElement createUserConfirmPasswordTxt;

	@FindBy(xpath="//input[@id='Username']")
	public static WebElement createUserNameTxt;

	@FindBy(xpath="//input[@id='lginAbbr']")
	public static WebElement createUserLoginAbbrivationTxt;

	@FindBy(xpath="//select[@id='SecurityQues']")
	public static WebElement createUserSecurityQuestionDropDown;

	@FindBy(xpath="//input[@id='SecurityAns']")
	public static WebElement createUserSecurityAnswerTxt;

	@FindBy(xpath="//select[@id='language']")
	public static WebElement createUserLanguageTxt;

	@FindBy(xpath="//select[@id='altLanguage']")
	public static WebElement createUserAlternateLanguageTxt;

	@FindBy(xpath="//i[@class='icon-unlock-user icon-font6']")
	public static WebElement createUserUnlockUserIcon;

	@FindBy(xpath="//i[@class='icon-move-user icon-font6']")
	public static WebElement createUserMoveUserIcon;

	@FindBy(xpath="//i[@class='icon-reset icon-font6']")
	public static WebElement createUserResetIcon;

	@FindBy(xpath="//span[@id='btnLoadFrom']")
	public static WebElement createUserLoadFromIcon;	
	
	
	
	
	//*[@id='btnSave']
	
	@FindBy(xpath="//*[@id='btnSave']//i[@class='icon-save hiconright2']")
	public static WebElement createUserSaveIcon;

	@FindBy(xpath="//*[@id='CancelCreateUser']")
	public static WebElement createUserCloseIcon;
	
	@FindBy(xpath="//select[@id='ERPRoles']/option")
	public static List<WebElement> createUserERPRolesDropdownOptions;
	
	@FindBy(xpath="//select[@id='language']/option")
	 public static List<WebElement> createUserLanguageDropdownList;
	
	 @FindBy(xpath="//select[@id='language']")
	 public static WebElement getLanguage;

	 
	 @FindBy(xpath="//select[@id='altLanguage']")
	 public static WebElement getAltLanguage;

	 
	 @FindBy(xpath="//select[@id='altLanguage']/option")
	 public static List<WebElement> createUserAltLanguageDropdownList;

	 @FindBy(xpath="//select[@id='typeOfUser']")
		public static WebElement createUserUserTypeDrpDwn;
		
		@FindBy(xpath="//input[@id='emailId']")
		public static WebElement createUserEmailTxt;
		
		
		@FindBy(xpath="//a[contains(text(),'Additional Info')]")
		public static WebElement createUserAdditionalInfoTab;
		
	public static boolean checkCreateUserforTriggers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		boolean flag=checkLogoutAndLoginWithSU();
		Thread.sleep(5000);
		
		/*getDriver().navigate().refresh();
		Thread.sleep(3000);*/
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		 homeMenu.click();
						
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		 securityMenu.click();

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		 createUserMenu.click();
		
		 Thread.sleep(2000);
	    
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
		 createUserIcon.click();
		 
		 Thread.sleep(10000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
		 createUserLoginNameTxt.click();
		 createUserLoginNameTxt.sendKeys("TriggerUser");

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		 createUserERPRoleDropDown.click();
		 createUserERPRoleDropDown.sendKeys("Role ST");
		 
	     int count = createUserERPRolesDropdownOptions.size()-1;
		 System.out.println("Count "+ count);
	    
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		 createUserPasswordTxt.sendKeys("12345");

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		 createUserConfirmPasswordTxt.sendKeys("12345");
		 	
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		 createUserNameTxt.sendKeys("UserCreation");
		 	
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
		 createUserLoginAbbrivationTxt.sendKeys("UC");
		 
		 createUserSecurityQuestionDropDown.click();
		 Select s = new Select(createUserSecurityQuestionDropDown);
		 s.selectByVisibleText("In which country were you born?");

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));	
		 createUserSecurityAnswerTxt.sendKeys("India");
		 	 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
			 getLanguage.click();
			 int counnt = createUserLanguageDropdownList.size()-1;
			 String actLanguageDropdownCount = Integer.toString(counnt);
			 String expLanguageDropdownCount = "2";
			 
			 Set LanguageDropdownList = new HashSet<String>();
			 
			 for (int i = 1; i < counnt; i++) 
			 {
				 String data = createUserLanguageDropdownList.get(i).getText();
				 LanguageDropdownList.add(data);
			 }
			 
			 String actLanguageDropdownList = LanguageDropdownList.toString();
			 String expLanguageDropdownList = "[English, Arabic]";
			 
			 System.out.println("LanguageDropdownCount  : "+actLanguageDropdownCount  +"  value expected  "+expLanguageDropdownCount);
			 System.out.println("LanguageDropdownList   : "+actLanguageDropdownList   +"  value expected  "+expLanguageDropdownList);
			 
			 getLanguage.sendKeys(Keys.TAB);
			 
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
				 getAltLanguage.click();
				 int cont = createUserAltLanguageDropdownList.size()-1;
				 String actAltLanguageDropdownCount = Integer.toString(count);
				 String expAltLanguageDropdownCount ="2";
				 
				 Set AltLanguageDropdownList = new HashSet<String>();
				 
				 for (int i = 1; i < cont; i++) 
				 {
					 String data = createUserAltLanguageDropdownList.get(i).getText();
					 AltLanguageDropdownList.add(data);
				 }
				 
				 String actAltLanguageDropdownList = AltLanguageDropdownList.toString();
				 String expAltLanguageDropdownList = "[English, Arabic]";
				 
				 getAltLanguage.sendKeys(Keys.TAB);
				 
				 Thread.sleep(2000);
		 
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
				 createUserAdditionalInfoTab.click();

				 Thread.sleep(2000);
		 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserUserTypeDrpDwn));
			 Select user = new Select(createUserUserTypeDrpDwn);
			 user.selectByVisibleText("Customer");
		 
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailTxt));
			createUserEmailTxt.click();
			createUserEmailTxt.clear();
			createUserEmailTxt.sendKeys("emailvalidationone@gmail.com");
			Thread.sleep(1500);	
			createUserEmailTxt.sendKeys(Keys.TAB);
			Thread.sleep(1500);

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSaveIcon));
		 createUserSaveIcon.click();
		// ClickUsingJs(createUserSaveIcon);
		 Thread.sleep(2000);
		 
		 String expMessage = "new user created successfully";
		 String actMessage = checkValidationMessage(expMessage);
		 
		 System.out.println(actMessage);
		 System.out.println(expMessage);
		 
		 
	     if (actMessage.equalsIgnoreCase(expMessage)) 
	     {
	    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
	    	 createUserCloseIcon.click();
	    	 return true;
	 	 } 
	 	 else 
	 	 {
	 		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
	    	 createUserCloseIcon.click();
	    	 
			return false;
	 	 }
	}
	
	
	
	
	
	@FindBy(xpath="//input[@id='triggers_chksendAlert']")
	public static WebElement sendAlertChkBox;
	
	
	
	@FindBy(xpath="(//*[@id='panelsStayOpen-collapseFour']//div[1]//div[1]//label/b)[1]")
	public static WebElement cretaNewAlertButton;
	
	@FindBy(xpath="//input[@id='smsTemplateOptCtrl']")
	private static WebElement emailTemplates;
	
	@FindBy(xpath="//i[@class='icon-open icon-font7']")
	private static WebElement companyMenu;
	
	@FindBy(xpath="//span[contains(text(),'Company Name')]")
	private static WebElement companyNameSource;
	
	@FindBy(xpath="//textarea[@id='txtTemplate']")
	private static WebElement companyNameTemplateArea;
	
	@FindBy(xpath="//span[@id='btnSaveSmsTemplate']")
	private static WebElement saveBtnInEmailTemplateHeader;
	
	@FindBy(xpath="//span[@id='btnNewSmsTemplate']")
	private static WebElement newButtonInCreateTemplate;
	
	@FindBy(xpath="//span[@id='btnCloseSMSModal']")
	private static WebElement closeButtonInCreateTemplate;
	
	
	@FindBy(xpath="//select[@id='triggers_sendAlert']")
	private static WebElement selectSendAlert;
	
	
	@FindBy(xpath="//input[@id='triggers_chksendAlert']/following-sibling::span")
	private static WebElement SendAlertChkBoxSelected;
	
	@FindBy(xpath="//input[@id='triggers_chksendAlert']")
	private static WebElement SendAlertChkBox;
	
	
	
	public static boolean checkCreatingAlertinTriggersPurchaseVoucherVAT() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Configure Transactions");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();	  
		
		Thread.sleep(2000);
		
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
		purchaseVouchersVATBtn.click();
						 	    	
		Thread.sleep(2000);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();
			
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys("Suspend Trigger");
		Thread.sleep(4000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersDeleteBtn));
		editLayoutTriggersDeleteBtn.click();
		Thread.sleep(2000);
		getWaitForAlert();
		Thread.sleep(2000);
		getAlert().accept();
		Thread.sleep(1000);
		
		checkValidationMessage("");
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Map Trigger");
		Thread.sleep(4000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(8000);
		
		
		
		if(SendAlertChkBox.isSelected()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SendAlertChkBoxSelected));
		SendAlertChkBoxSelected.click();
		}
		Thread.sleep(2000);
		
				
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cretaNewAlertButton));
		cretaNewAlertButton.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newButtonInCreateTemplate));
		newButtonInCreateTemplate.click();
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplates));
		emailTemplates.sendKeys("MapAlertTriggerTemplate");
		emailTemplates.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
		companyNameTemplateArea.click();
		Thread.sleep(1500);
		

		//Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\TriggerAlerts.exe");
		
		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailHeaderTemplateNew.exe");
		
		Thread.sleep(175000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
		saveBtnInEmailTemplateHeader.click();
		
		Thread.sleep(2000);
		System.err.println("Template save button clicked");
		
		String expEmailTemplateSaveMsg = "Template saved successfully";
		String actEmailTemplateSaveMsg =  checkValidationMessage(expEmailTemplateSaveMsg);
		
		ClickUsingJs(closeButtonInCreateTemplate);
		
	
		Thread.sleep(2000);
		if(actEmailTemplateSaveMsg.equalsIgnoreCase(expEmailTemplateSaveMsg))
		{
			return true;
		}
		else
		{
			return false;
		}
		
			 
	}
	
	
	public static boolean checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		
	     	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
	     	logoutOption.click();

	     	Thread.sleep(3000);
	     	getDriver().navigate().refresh();
			Thread.sleep(2000);
		LoginPage lp=new LoginPage(getDriver()); 
		
	    String unamelt="su";
	  
	    String pawslt="su";
	      
	    lp.enterUserName(unamelt);
	    
	    lp.enterPassword(pawslt);
	    
	    Thread.sleep(2000);

        String compname = "Automation Company : 08/10/2020";

        Select oSelect = new Select(companyDropDownList);

        List<WebElement> elementCount = oSelect.getOptions();

        int cqSize = elementCount.size();

        System.out.println("CompanyDropdownList Count :" + cqSize);

        int i;

        for (i = 0; i < elementCount.size(); i++) {

                elementCount.get(i).getText();

                String optionName = elementCount.get(i).getText();
                if (optionName.toUpperCase().startsWith(compname.toUpperCase())) {
                        System.out.println("q" + elementCount.get(i).getText());
                        elementCount.get(i).click();
                }

        }

        Thread.sleep(2000);

	    
	    lp.clickOnSignInBtn();
	    
	    //checkRefershPopOnlogin();
	    
	    //checkPopUpWindow();
	  
		Thread.sleep(6000);

		String actUserInfo=userNameDisplay.getText();
		
		System.out.println("User Info  : "+actUserInfo);
		
		System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());
		
		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		//companyLogo.click();
		
		String getCompanyTxt=Company_Name.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		//companyLogo.click();
		
		
		String expUserInfo           ="su";
		String expLoginCompanyName   ="Automation";
		
		boolean actDashboard = dashboard.isDisplayed();
		boolean expDashboard = true;
		
		System.out.println("**********************************************checkLogoutAndLoginWithSU*********************************************");
		
		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard==expDashboard) 
		{
			return true;
		}
		else
		{
			return false;
		}
	}
  	
	
	
	
	
	public static boolean checkCreatedAlertandUserSelectinPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Configure Transactions");
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();	  
		
		Thread.sleep(10000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
		purchaseVouchersVATBtn.click();
						 	    	
		Thread.sleep(8000);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Map Trigger");
		Thread.sleep(4000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(8000);
		
		/*if(SendAlertChkBox.isDisplayed()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
		triggerTabAlertBtn.click();
		}
		Thread.sleep(2000);*/
		
		new WebDriverWait(getDriver(), 100).until(ExpectedConditions.visibilityOf(SendAlertChkBoxSelected));
		
		if(sendAlertChkBox.isSelected()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SendAlertChkBoxSelected));
		SendAlertChkBoxSelected.click();
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectSendAlert));
		Select s=new Select(selectSendAlert);
		s.selectByVisibleText("MapAlertTriggerTemplate");
		Thread.sleep(2000);
		String actMsg=s.getFirstSelectedOption().getText();
		String expMsg="MapAlertTriggerTemplate";
		
		System.out.println("Actual Message: 	" 		+ actMsg);
		System.out.println("Expected Message:  	"	+expMsg);
		
		getAction().moveToElement(loginChkBoxSelected).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginChkBoxSelected));
		if(loginChkBox.isSelected()==false)
		{
			loginChkBoxSelected.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectUsersList));
		Select s1=new Select(selectUsersList);
		s1.selectByVisibleText("UserAllOptionsST");
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		Thread.sleep(3000);
		
		String expSaveMessage = "Data saved successfully";
		String actSaveMessage = checkValidationMessage(expSaveMessage);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		
		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();
		
		
		if(actMsg.equalsIgnoreCase(expMsg) && actSaveMessage.equalsIgnoreCase(expSaveMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	public static boolean checkSavingVoucher2inPVVATforAlerts() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

		Thread.sleep(4000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();

		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PVVAT_SettingsBtn));
		PVVAT_SettingsBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(showDependency1));
		showDependency1.click();
		Thread.sleep(2000);
		
		getAction().moveToElement(updateBtn).build().perform();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Settings_closeBtn));
		Settings_closeBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();

		vendorAccountTxt.sendKeys(Keys.SPACE);	

		ArrayList<String>  actVendorAccountList= new ArrayList<String>(); 

		int count=vendorAccountListCount.size();

		System.err.println(count);

		for(int i=0 ; i < count ;i++)
		{
			String data=vendorAccountListCount.get(i).getText();

			if (data.equalsIgnoreCase("Vendor B")) 
			{
				vendorAccountListCount.get(i).click();
				break;
			}
		}

		vendorAccountTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(Keys.SPACE);

		int currencycount=currencyListCount.size();

		System.err.println(currencycount);

		for(int i=0 ; i < currencycount ;i++)
		{
			String data=currencyListCount.get(i).getText();

			if (data.equalsIgnoreCase("INR")) 
			{
				currencyListCount.get(i).click();
				break;
			}
		}

		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);	

		ArrayList<String> actDepartmentList= new ArrayList<String>(); 

		int departmentCount=departmentListCount.size();

		System.err.println(departmentCount);

		for(int i=0 ; i < departmentCount ;i++)
		{
			String data=departmentListCount.get(i).getText();

			if (data.equalsIgnoreCase("INDIA")) 
			{
				departmentListCount.get(i).click();
				break;
			}
		}

		departmentTxt.sendKeys(Keys.TAB);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.clear();
		placeOFSupplyTxt.sendKeys("Abu Dhabi");
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys("DUBAI");
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();

		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount=warehouseBodyComboList.size();

		for(int i=0 ; i < warehousecount ;i++)
		{
			String data=warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("HYDERABAD")) 
			{
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount=pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount; i++) 
		{
			String Item=pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("BR COGS ITEM")) 
			{
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("2");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));
		select1stRow_13thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
		enter_Discount.click();
		enter_Discount.clear();
		enter_Discount.sendKeys("10");
		enter_Discount.sendKeys(Keys.TAB);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.sendKeys(Keys.TAB);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys("PVATR1#V1");
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 2); 

		String next2DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "+next2DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next2DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		Thread.sleep(2000);

				
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": SU/IND/TEXT6";
		
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) )
		{
			return true;
		} 
		else 
		{
			return false;
		}
	
	
	}
	
	
	
		public boolean checkLogoutandLoginwithUserAllOptionsSTforAlertMessage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
    	{
			getDriver().navigate().refresh();
			Thread.sleep(2000);
    		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
    		userNameDisplayLogo.click();
   		
   	     	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
   	     	logoutOption.click();

   	     	Thread.sleep(3000);
   	     	
    		LoginPage lp=new LoginPage(getDriver()); 
    		
    	    String unamelt="UserAllOptionsST";
    	  
    	    String pawslt="12345";
    	      
    	    lp.enterUserName(unamelt);
    	    
    	    lp.enterPassword(pawslt);
    	    
    	    Thread.sleep(2000);

            String compname = "Automation Company : 08/10/2020";

            Select oSelect = new Select(companyDropDownList);

            List<WebElement> elementCount = oSelect.getOptions();

            int cqSize = elementCount.size();

            System.out.println("CompanyDropdownList Count :" + cqSize);

            int i;

            for (i = 0; i < elementCount.size(); i++) {

                    elementCount.get(i).getText();

                    String optionName = elementCount.get(i).getText();
                    if (optionName.toUpperCase().startsWith(compname.toUpperCase())) {
                            System.out.println("q" + elementCount.get(i).getText());
                            elementCount.get(i).click();
                    }

            }

            Thread.sleep(2000);

    	    
    	    lp.clickOnSignInBtn();
    	    
    	    //checkRefershPopOnlogin();
    	    
    	    //checkPopUpWindow();
    	  
    		Thread.sleep(6000);

    		String actUserInfo=userNameUserDisplay.getText();
    		
    		System.out.println("User Info  : "+actUserInfo);
    		
    		System.out.println("User Info Capture Text  :  "+userNameUserDisplay.getText());
    		
    		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
    		//companyLogo.click();
    		
    		String getCompanyTxt=Company_Name.getText();
    		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
    		System.out.println("company name  :  "+ getLoginCompanyName);
    		//companyLogo.click();
    		
    		
    		String expUserInfo           ="UserAllOptionsST";
    		String expLoginCompanyName   ="Automation";
    		
    		boolean actDashboard = dashboard.isDisplayed();
    		boolean expDashboard = true;
    		
    		System.out.println("**********************************************checkLogoutAndLoginWithUserAllOptions*********************************************");
    		
    		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
    		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
    		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

    		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard==expDashboard) 
    		{
				return true;
			}
    		else
    		{
    			return false;
			}
    	
	}
	
		 @FindBy(xpath="//*[@id='noofalerts']")
		  private static WebElement noOfAlertsDisplay;
		 
		 
		 @FindBy(xpath="//*[@class='icon-alerts-1 hiconright recentdropdown']")
		 public static WebElement alertsTab;
		 
		 @FindBy(xpath="//*[@id='id_mainlayoutmenu']/ul[2]/li[4]/a/i")
		  private static WebElement noOfAlertsDisplaySelcted;
		 

		  @FindBy(xpath="//li[@class='header']")
		  private static WebElement alertsMainHeader;

		  @FindBy(xpath="//ul[@id='alertsUL']/li[2]/ul/li/a")
			 private static List<WebElement> alertTransaction;
		
		  @FindBy(xpath="(//ul[@id='alertsUL']/li[2]/ul/li/ul/li/a)[1]")
			 private static List<WebElement> transactionAuthorization;
		  
		  
		  @FindBy(xpath="//ul[@id='alertsUL']/li[2]/ul/li/ul/li/ul/li/a/span")
			 private static List<WebElement> transactionAuthorizationText;
		  
		  @FindBy(xpath="(//ul[@id='alertsUL']/li[2]/ul/li/ul/li/a)[2]")
			 private static List<WebElement> transactionAuthorization2;
		  
		  
		  @FindBy(xpath="(//ul[@id='alertsUL']/li[2]/ul/li/ul/li/ul/li/a/span)[2]")
			 private static List<WebElement> transactionAuthorizationText2;
		  
		  
		
		public static boolean checkTriggerAlertMessageinUserAllOptionsST() throws InterruptedException
	{
			
			getDriver().navigate().refresh();
			Thread.sleep(2000);
			
		getWebDriverWait().until(ExpectedConditions.visibilityOf(noOfAlertsDisplay));
		noOfAlertsDisplaySelcted.click();
		
		// Actual
		String  actAlertsMainHeader=alertsMainHeader.getText();
		
		boolean actNoOfAlertsDisplay=alertsMainHeader.isDisplayed();
		boolean expNoOfAlertsDisplay=true;
		
		System.out.println("Alerts is display : "+ actAlertsMainHeader);
		
		
		
		int count = alertTransaction.size();
		
		String alertTest=null;
		
		for(int i=0; i<count; i++)
		{
			String data = alertTransaction.get(i).getText();
			
			if(data.equalsIgnoreCase("  Transactions (1)"))
			{
				alertTransaction.get(i).click();
				Thread.sleep(2000);
				transactionAuthorization.get(i).click();
				Thread.sleep(2000);
				alertTest=transactionAuthorizationText.get(i).getText();
				
				
			}
		}
		
		
		String actAlertMsg=alertTest;
		String expAlertMsg="Automation Company : 08/10/2020 5:43 PMSUBR COGS ITEMHYDERABAD21020.00NDT57:SU/IND/TEXT6";
		
		System.out.println("Actual Alert Text: 		" 	+	alertTest);
		System.out.println("Expected Alert Text: 	" 	+	expAlertMsg);
		
		if(actNoOfAlertsDisplay==expNoOfAlertsDisplay && actAlertMsg.equalsIgnoreCase(expAlertMsg))
		{
		return true;
		}
		else
		{
			return false;
		}
	}
	
	
	
	@FindBy(xpath="//input[@id='triggers_chksendEmail']")
	public static WebElement sendEmailChkBox;
	
	@FindBy(xpath="//input[@id='triggers_chksendEmail']/following-sibling::span")
	public static WebElement sendEmailChkBoxSelected;
	
	
	
	@FindBy(xpath="//input[@id='Triggers_LayoutID']")
	public static WebElement sendEmailTxtBox;
	
	@FindBy(xpath="//*[@id='panelsStayOpen-collapseFour']//div[7]//label/b")
	public static WebElement createNewEmailButton;
	
	
	@FindBy(xpath="//*[@title='New']")
	public static WebElement emailNewBtn;
	

	@FindBy(xpath="(//*[@id='SliderMenu']/a[1])[1]")
	public static WebElement emailNewLayoutBtn;
	
		
	@FindBy(xpath="//*[@id='drpdownSlider']//*[@class='icon-save hiconright2']")
	public static WebElement emailSaveBtn;
	
	
	@FindBy(xpath="(//*[@id='SliderMenu']/a[1])[3]")
	public static WebElement emailSaveOptionBtn;
	
	@FindBy(xpath="//h4[@id='popup_title']")
	public static WebElement emailSavePopupTitle;
	
	@FindBy(xpath="//input[@id='id_PopuplayoutName']")
	public static WebElement emailSaveTxt;
	
	
	@FindBy(xpath="//*[@id='btn_PopUpOk']")
	public static WebElement emailSaveLayoutBtn;
	
	
	
	
	@FindBy(xpath="//*[@id='DesignMenu']")
	public static WebElement emailLayoutDesignBtn;
	
	@FindBy(xpath="//*[@id='li_InvoicePreview']/a")
	public static WebElement emailLayoutPreviewBtn;
	
	
	@FindBy(xpath="//*[@title='Exit']")
	public static WebElement emailLayoutExitBtn;
	
	@FindBy(xpath="//input[@id='triggers_chksendSMS']")
	public static WebElement sendSMSChkBox;
	
	@FindBy(xpath="//input[@id='triggers_chksendSMS']/following-sibling::span")
	public static WebElement sendSMSChkBoxSelected;
	
	
	@FindBy(xpath="//*[@id='panelsStayOpen-collapseFour']//div[3]/div[2]//label/b")
	public static WebElement createNewSMSBtn;
	
	
	//input[@id='triggers_chksendTypeLogin']
	
	@FindBy(xpath="//input[@id='triggers_chksendTypeLogin']")
	public static WebElement loginChkBox;
	
	@FindBy(xpath="//input[@id='triggers_chksendTypeLogin']/following-sibling::span")
	public static WebElement loginChkBoxSelected;
	
	@FindBy(xpath="//input[@id='triggers_chksendTypeCustomeVendor']")
	public static WebElement customerChkBox;
	
	
	@FindBy(xpath="//input[@id='triggers_chksendTypeCustomeVendor']/following-sibling::span")
	public static WebElement customerChkBoxSelected;
	
	@FindBy(xpath="//select[@id='triggers_sendTypeloginUsers']")
	public static WebElement selectUsersList;
	
	
	
	
	public static boolean checkEmailTemplateCreationinPurchaseVoucherVATTriggers() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Configure Transactions");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();	  
		
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
		purchaseVouchersVATBtn.click();
						 	    	
		Thread.sleep(2000);
	
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Map Trigger");
		Thread.sleep(4000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(8000);
	
		if(SendAlertChkBox.isDisplayed()==true)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SendAlertChkBoxSelected));
		SendAlertChkBoxSelected.click();
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
		if(sendEmailChkBox.isSelected()==false)
		{
		sendEmailChkBoxSelected.click();
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createNewEmailButton));
		createNewEmailButton.click();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailNewBtn));
		emailNewBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailNewLayoutBtn));
		emailNewLayoutBtn.click();
		Thread.sleep(3000);
		
		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailTriggerInvoiceLayout.exe");
		
		Thread.sleep(210000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailLayoutPreviewBtn));
		emailLayoutPreviewBtn.click();
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailLayoutDesignBtn));
		emailLayoutDesignBtn.click();
		Thread.sleep(3000);*/
		
		/*getAction().moveToElement(emailSaveBtn).build().perform();
		Thread.sleep(2000);*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveBtn));
		emailSaveBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveOptionBtn));
		emailSaveOptionBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSavePopupTitle));
		boolean actTitleDisplay=emailSavePopupTitle.isDisplayed();
		boolean expTitleDisplay=true;
		
		
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveTxt));
		emailSaveTxt.click();
		emailSaveTxt.sendKeys("TriggersEmailLayout");
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveLayoutBtn));
		emailSaveLayoutBtn.click();
		Thread.sleep(2000);
		String expSaveLayotMsg="Layout saved successfully";
		String actSaveLayoutMsg=checkValidationMessage(expSaveLayotMsg);
		
		
		System.out.println("Actual Message: 	"	+actSaveLayoutMsg);
		System.out.println("Expected Message:	"	+expSaveLayotMsg);
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailLayoutExitBtn));
		emailLayoutExitBtn.click();
		Thread.sleep(2000);
		
		if(actTitleDisplay==expTitleDisplay && actSaveLayoutMsg.equalsIgnoreCase(expSaveLayotMsg))
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	
	@FindBy(xpath="//*[@id='id_trigger_email_subject']")
	public static WebElement emailSubjectTxt;
	
	@FindBy(xpath="//*[@id='id_trigger_btn_emailheadertemplate']")
	public static WebElement emailHeaderTemplateBtn;
	
	@FindBy(xpath="//*[@id='id_trigger_emailheadertemplate']")
	public static WebElement emailHeaderTemplateTxt;
	
	public static boolean checkCreatedEmailTemplateinPVVATView() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		/*Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(6000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Configure Transactions");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();	  
		
		Thread.sleep(6000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
		purchaseVouchersVATBtn.click();
						 	    	
		Thread.sleep(8000);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Map Trigger");
		Thread.sleep(4000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(8000);*/
		
		
		
		/*if(SendAlertChkBox.isDisplayed()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
		triggerTabAlertBtn.click();
		}
		Thread.sleep(2000);*/
		
		new WebDriverWait(getDriver(), 300).until(ExpectedConditions.visibilityOf(SendAlertChkBoxSelected));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SendAlertChkBoxSelected));
		if(sendAlertChkBox.isSelected()==true)
		{
			SendAlertChkBoxSelected.click();
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
		if(sendEmailChkBox.isSelected()==false)
		{
		sendEmailChkBoxSelected.click();
		}
		Thread.sleep(4000);
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailTxtBox));
		sendEmailTxtBox.click();
		sendEmailTxtBox.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
		sendEmailTxtBox.sendKeys("TriggersEmailLayout");
		Thread.sleep(3000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSubjectTxt));
		emailSubjectTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		emailSubjectTxt.sendKeys("Email Trigger Voucher");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateBtn));
		emailHeaderTemplateBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newButtonInCreateTemplate));
		newButtonInCreateTemplate.click();
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplates));
		emailTemplates.sendKeys("EmailHeaderTriggerTemplate");
		emailTemplates.sendKeys(Keys.TAB);

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
		companyNameTemplateArea.click();
		Thread.sleep(1500);*/

		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailTemplate.exe");
		Thread.sleep(175000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
		saveBtnInEmailTemplateHeader.click();

		Thread.sleep(2000);
		System.err.println("Template save button clicked");

		String expEmailTemplateSaveMsg = "Template saved successfully";
		String actEmailTemplateSaveMsg = checkValidationMessage(expEmailTemplateSaveMsg);
		ClickUsingJs(closeButtonInCreateTemplate);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
		closeButtonInCreateTemplate.click();*/
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateTxt));
		emailHeaderTemplateTxt.click();
		emailHeaderTemplateTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		emailHeaderTemplateTxt.sendKeys("EmailHeaderTriggerTemplate");
		Thread.sleep(2000);

		getAction().moveToElement(loginChkBoxSelected).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginChkBoxSelected));
		if(loginChkBox.isSelected()==false)
		{
			loginChkBoxSelected.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectUsersList));
		Select s1=new Select(selectUsersList);
		s1.selectByVisibleText("TriggerUser");
		Thread.sleep(2000);
		/*s1.selectByVisibleText("UserAllOptionsST");
		Thread.sleep(2000);*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		Thread.sleep(2500);
		
		String expSaveMessage = "Data saved successfully";
		String actSaveMessage = checkValidationMessage(expSaveMessage);
		Thread.sleep(2000);
		
		getAction().moveToElement(updateBtn).build().perform();;
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		
		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();
		
		
		if(actSaveMessage.equalsIgnoreCase(expSaveMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
	
		
	}
	
	
	@FindBy(xpath="//select[@id='triggers_sendSMS']")
	public static WebElement selectSMSTemplateAlert;
	
	
	
	public static boolean checkCreatingSMSTemplateforTriggersinPVVATVoucher() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Configure Transactions");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();	  
		
		Thread.sleep(8000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
		purchaseVouchersVATBtn.click();
						 	    	
		Thread.sleep(8000);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Map Trigger");
		Thread.sleep(4000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(8000);
		
		
		if(SendAlertChkBox.isSelected()==true)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SendAlertChkBoxSelected));
		SendAlertChkBoxSelected.click();
		}
		Thread.sleep(2000);
		
		if(sendEmailChkBox.isSelected()==true)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
		sendEmailChkBoxSelected.click();
		}
		Thread.sleep(2000);
		
		
		getAction().moveToElement(sendSMSChkBox).build().perform();
		Thread.sleep(2000);
		
		if(sendSMSChkBox.isSelected()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendSMSChkBoxSelected));
		sendSMSChkBoxSelected.click();
		}
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createNewSMSBtn));
		createNewSMSBtn.click();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newButtonInCreateTemplate));
		newButtonInCreateTemplate.click();
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplates));
		emailTemplates.sendKeys("MapSMSTriggerTemplate");
		emailTemplates.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
		companyNameTemplateArea.click();
		Thread.sleep(1500);
		

		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailHeaderTemplateNew.exe");
		
		Thread.sleep(175000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
		saveBtnInEmailTemplateHeader.click();
		
		Thread.sleep(2000);
		System.err.println("Template save button clicked");
		
		String expEmailTemplateSaveMsg = "Template saved successfully";
		String actEmailTemplateSaveMsg =  checkValidationMessage(expEmailTemplateSaveMsg);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
		closeButtonInCreateTemplate.click();*/
		ClickUsingJs(closeButtonInCreateTemplate);
		Thread.sleep(2000);
		if(actEmailTemplateSaveMsg.equalsIgnoreCase(expEmailTemplateSaveMsg))
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	
	public boolean checkCreatedSMSAlertinPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(6000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Configure Transactions");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();	  
		
		Thread.sleep(8000);
		
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
		purchaseVouchersVATBtn.click();
						 	    	
		Thread.sleep(8000);
	
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Map Trigger");
		Thread.sleep(4000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SendAlertChkBoxSelected));
		if(sendAlertChkBox.isSelected()==true)
		{
			SendAlertChkBoxSelected.click();
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
		if(sendEmailChkBox.isSelected()==true)
		{
		sendEmailChkBoxSelected.click();
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendSMSChkBoxSelected));
		if(sendSMSChkBox.isSelected()==false)
		{
			sendSMSChkBoxSelected.click();
		}
			
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectSMSTemplateAlert));
		Select s=new Select(selectSMSTemplateAlert);
		s.selectByVisibleText("MapSMSTriggerTemplate");
		Thread.sleep(2000);
		String actMsg=s.getFirstSelectedOption().getText();
		String expMsg="MapSMSTriggerTemplate";
		
		System.out.println("Actual Message: 	" 		+ actMsg);
		System.out.println("Expected Message:  	"	+expMsg);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		Thread.sleep(2000);
		
		String expSaveMessage = "Data saved successfully";
		String actSaveMessage = checkValidationMessage(expSaveMessage);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		
		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);
		
		Thread.sleep(10000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();
		
		Thread.sleep(8000);
		
		
		if(actMsg.equalsIgnoreCase(expMsg))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	public static boolean checkSavingVoucher3inPVVATforEmailValidation() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

		Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();

		vendorAccountTxt.sendKeys(Keys.SPACE);	

		ArrayList<String>  actVendorAccountList= new ArrayList<String>(); 

		int count=vendorAccountListCount.size();

		System.err.println(count);

		for(int i=0 ; i < count ;i++)
		{
			String data=vendorAccountListCount.get(i).getText();

			if (data.equalsIgnoreCase("Vendor B")) 
			{
				vendorAccountListCount.get(i).click();
				break;
			}
		}

		vendorAccountTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(Keys.SPACE);

		int currencycount=currencyListCount.size();

		System.err.println(currencycount);

		for(int i=0 ; i < currencycount ;i++)
		{
			String data=currencyListCount.get(i).getText();

			if (data.equalsIgnoreCase("INR")) 
			{
				currencyListCount.get(i).click();
				break;
			}
		}

		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);	

		ArrayList<String> actDepartmentList= new ArrayList<String>(); 

		int departmentCount=departmentListCount.size();

		System.err.println(departmentCount);

		for(int i=0 ; i < departmentCount ;i++)
		{
			String data=departmentListCount.get(i).getText();

			if (data.equalsIgnoreCase("INDIA")) 
			{
				departmentListCount.get(i).click();
				break;
			}
		}

		departmentTxt.sendKeys(Keys.TAB);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.clear();
		placeOFSupplyTxt.sendKeys("Abu Dhabi");
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys("DUBAI");
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();

		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount=warehouseBodyComboList.size();

		for(int i=0 ; i < warehousecount ;i++)
		{
			String data=warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("HYDERABAD")) 
			{
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount=pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount; i++) 
		{
			String Item=pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) 
			{
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("2");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));
		select1stRow_13thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
		enter_Discount.click();
		enter_Discount.clear();
		enter_Discount.sendKeys("10");
		enter_Discount.sendKeys(Keys.TAB);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.sendKeys(Keys.TAB);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys("PVATR1#V1");
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 2); 

		String next2DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "+next2DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next2DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		Thread.sleep(2000);

				
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": SU/IND/TEXT7";
		
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) )
		{
			return true;
		} 
		else 
		{
			return false;
		}
	
	
	}
	
	
	   @FindBy(xpath="//span[contains(text(),'Next')]")
		private static WebElement  NextBtn;
	
	   @FindBy(xpath="//table[@id=':22']/tbody/tr/td[4]/div[2]/span[1]/span")
	    private static List<WebElement> row1mailFromList;
	
	@FindBy(xpath="//input[@name='Passwd']")
	private static WebElement  PasswordTxt;

	@FindBy(xpath="//div[@id=':2i']")
	private static WebElement  row1;

	@FindBy(xpath="//input[@id='HeaderChkBox']")
	private static WebElement voucherHomeHeaderCheckbox;

	/*@FindBy(xpath="//span[contains(text(),'Next')]")
	private static WebElement NextBtn;
*/

	//@FindBy(xpath="//*[@id='yDmH0d']/div[5]/div/div[2]/div[3]/div[1]/span/span")
	@FindBy(xpath="//*[contains(text(),'Yes, remove')]")
	private static WebElement yesRemoveBtn;

/*	@FindBy(xpath="//table[@id=':23']/tbody/tr/td[5]/div[1]/div/div/span/span")
	private static List<WebElement> row1mailFromList;

		@FindBy(xpath="//input[@id='identifierId']")
	private static WebElement  userNameTxt;
*/


		//@FindBy(xpath="//div[text()='Attachments area']/parent::div/div[4]//span/div/div[1]")
	@FindBy(xpath="(//*[@data-tooltip='Download'])[1]")
		private static WebElement mailAttachmentDownloadBtn;
	    
	@FindBy(xpath = "(//div[@class='VYBDae-JX-ano'])[1]")
	private static WebElement mailAttachmentDownloadBtn1;
		
	/*@FindBy(xpath="//*[@id=':4']/div[2]/div[1]/div/div[2]/div[3]")
	private static WebElement mailDeleteButton;
*/
	@FindBy(xpath="//*[@id=':4']/div[2]/div[1]/div/div[1]/div/div")
	private static WebElement backToIndexBtn;



	/*@FindBy(xpath="//*[@id='gb']/div[2]/div[3]/div[1]/div[2]/div/a/img")
	private static WebElement gmailUserBtn;

	@FindBy(xpath="//div[@id='yDmH0d']//*[text()='Sign out']")
	private static WebElement gmailSignOutBtn;
*/
	
	//(//*[@id='gb']//iframe)[2]
	
	@FindBy(xpath="//iframe[@name='account']")
	private static WebElement SignOutFrame;




	@FindBy(xpath="//*[contains(text(),'Remove an account')]")
	private static WebElement removeAccountBtn;

	@FindBy(xpath="(//*[@jsname='MBVUVe']//div)[8]")
	private static WebElement removeDeleteBtn;


	/*@FindBy(xpath="//*[@id='yDmH0d']/div[5]/div/div[2]/div[3]/div[1]/span/span")
	private static WebElement yesRemoveBtn;
	*/

	@FindBy(xpath="//*[@id=':2d']")
	private static WebElement emailBody;
	
	@FindBy(xpath="//*[@id=':2c']")
	private static WebElement emailBody1;
	
	//tbody/tr/td[5]/div[1]/div/div/span/span
	@FindBy(xpath="//tbody/tr/td[5]/div[1]/div/div/span/span")
	private static List<WebElement> emailSubjectList;

	@FindBy(xpath="//tbody/tr[1]//td[5]//div[1]/div/div")
	private static WebElement emailRow1Subject;

	//@FindBy(xpath="//*[@id=':26']/div[1]/span")
	@FindBy(xpath="(//div[@class='G-Ni J-J5-Ji'])[1]//span")
	private static WebElement selectAllMailsChckbx;

	@FindBy(xpath="//*[@id=':4']/div/div[1]/div[1]/div/div/div[2]/div[3]")
	private static WebElement selectAllMailsDltBtn;

	//Email

	@FindBy(xpath="//*[@id='tomailId']")
	public static WebElement ToAddressInMail;

	@FindBy(xpath="//*[@id='subjectMail']")
	public static WebElement subjectInMail;


	@FindBy(xpath="//*[@id='emailSettings']//button[2]")
	public static WebElement okBtnInEmailOptions;



	@FindBy(xpath="//*[@id='attachpdf']")
	public static WebElement attachPDFChkBox;

	//Template header

	@FindBy(xpath="//*[@id='btnEmailheaer']")
	public static WebElement emailTemplateHeaderInEmailOptions;

	@FindBy(xpath="//*[@id='smsTemplateOptCtrl']")
	public static WebElement templateName;

	@FindBy(xpath="//*[@id='btnSaveSmsTemplate']")
	public static WebElement templateSaveBtn;


	@FindBy(xpath="//*[@id='btnCloseSMSModal']")
	public static WebElement templateCloseBtn;


	@FindBy(xpath="//input[@id='identifierId']")
	private static WebElement  userNameTxt;
	
	
	
	  @FindBy(xpath="//*[@id=':4']/div[2]/div[1]/div/div[2]/div[3]")
	    private static WebElement mailDeleteButton;
	    
	    @FindBy(xpath="//*[@id='gb']/div[2]/div[3]/div[1]/div[2]/div/a/img")
	    private static WebElement gmailUserBtn;
	
	
	   // @FindBy(xpath="//div[@id='yDmH0d']//*[text()='Sign out']")
	    @FindBy(xpath="//div[text()='Sign out']")
	    private static WebElement gmailSignOutBtn;
	    
	    @FindBy(xpath="//div[text()='Manage your Google Account']")
	    private static WebElement text;
	    
	    
	  
		@FindBy(xpath = "//*[@id=':28']")
		private static WebElement emailBody2;
	
	
	
	
	
public static boolean checkEmailValidationforLogininCreatedTriggerTemplate() throws InterruptedException, AWTException, IOException
{
	Thread.sleep(3000);
	
	File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\TriggersEmailLogin.pdf");
	
	if(Efile1.exists())
	{
		Efile1.delete();
	}
	
		Robot robot = new Robot();   
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);
		
		Thread.sleep(6000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

	 	System.out.println("openTabs"+openTabs);

	 	getDriver().switchTo().window(openTabs.get(0));
	 	Thread.sleep(2000);
	 	getDriver().switchTo().window(openTabs.get(1));
	 	Thread.sleep(2000);
	
	 	getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");
	 	
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
	 	userNameTxt.click();
	 	
	 	userNameTxt.sendKeys("emailvalidationone@gmail.com");
	 	
	 	Thread.sleep(2000);
	 	
	 	System.out.println("User enter text");
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
	 	NextBtn.click();
	 	
	 	Thread.sleep(2000);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
	 	PasswordTxt.click();
	 	
	 	PasswordTxt.sendKeys("validationone");
	 	
	 	Thread.sleep(2000);
	 	
	 	NextBtn.click();
	 	
	 	Thread.sleep(10000);

	 	int count1 = row1mailFromList.size();
	 //	System.out.println("No.of mails in the list"+count1);
	 	
	 //	System.err.println("Email Body *"+emailBody2.getText()+"*");
	 	
	 	if (emailBody2.getText().isEmpty()==false) 
	 	{
	 		int count = emailSubjectList.size();
	 		System.err.println("No.of mails in the list"+count);
		 	
		 	for (int i = 0; i < count; i++) 
	 	{
	 		String mailFromTxt = emailSubjectList.get(i).getText();
	 		
	 		if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher")) 
	 		{
	 			emailSubjectList.get(i).click();
	 			break;
			}
		}
		 	
		 	
		 	Thread.sleep(3000);
 	
 	Thread.sleep(3000);
 	
 	getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
 	
 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
 	mailAttachmentDownloadBtn1.click();
 	
 	Thread.sleep(6000);
 	
 	 
 	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_CONTROL);
	
	
	Thread.sleep(5000);
	
	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);
	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);
	
	robot.keyPress(KeyEvent.VK_ENTER);
	robot.keyRelease(KeyEvent.VK_ENTER);
	
	Thread.sleep(3000);
	
	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_CONTROL);
 	
	Thread.sleep(2000);
	
			
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\TriggersEmailLogin.exe");
			
			Thread.sleep(6000);
			
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(2000);
			
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
			int actOpenWindowsCount = getDriver().getWindowHandles().size();
			int expOpenWindowsCount = 4;
			
			System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
			
			getDriver().switchTo().window(newTabs.get(3)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(2)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(1));
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 	mailDeleteButton.click();
		 	
		 	Thread.sleep(6000);
		 	

		 	
		 	
		 /*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailUserBtn));
		 	gmailUserBtn.click();*/
		 	ClickUsingJs(gmailUserBtn);
		 	
		 	Thread.sleep(5000);
		 	getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(5000);
			
			
		 	//getAction().moveToElement(gmailSignOutBtn).click().build().perform();
			//ScrollIntoView(gmailSignOutBtn);
			
		
			//getAction().moveToElement(gmailSignOutBtn).build().perform();
			ClickUsingJs(gmailSignOutBtn);
		
			Thread.sleep(2000);
			
			getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
			
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 	removeAccountBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 	removeDeleteBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 	yesRemoveBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
				 
		 	getDriver().switchTo().window(newTabs.get(1)).close();
		 	Thread.sleep(2000);
		 	getDriver().switchTo().window(newTabs.get(0));
		 	Thread.sleep(2000);
		 
		 	
		 	String actAccount1PDF = getBaseDir()+"\\autoIt\\ExportFiles\\TriggersEmailLogin.pdf";
			String expAccount1PDF = getBaseDir()+"\\autoIt\\ImportFiles\\TriggersEmailLogin.pdf";
			
			PDFUtil pdfutil = new PDFUtil();
			
			
			
			boolean result = pdfutil.compare(actAccount1PDF, expAccount1PDF);
			Calendar cal=Calendar.getInstance();
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = df.format(cal.getTime());
			
			String actAccount1Data = pdfutil.getText(actAccount1PDF);
			String expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("28/10/2024", currentDate);
					
			System.err.println(actAccount1Data);
			System.err.println(expAccount1Data);
			
		
		 	
			
			if (actAccount1Data.equalsIgnoreCase(expAccount1Data)) 
			{
			
				System.err.println("PDF file is as Expected");
				return true;
			}
			else
			{
									
				System.err.println("PDF file is not as Expected");
				return false;
			}
			
			
		} 
		 	else 
		{	
		 		
		 		
		 		//System.err.println("Subject : "+emailRow1Subject.getText());
		 		
		 		Thread.sleep(2000);
		 		
		 		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
		 	ClickUsingJs(gmailUserBtn);
		 	
		 	Thread.sleep(3000);
		 getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(1000);
		 	getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 	Thread.sleep(6000);
		 	
		 	getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 	removeAccountBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 	removeDeleteBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 	yesRemoveBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
			getDriver().switchTo().window(newTabs.get(1)).close();
		 	Thread.sleep(2000);
		 	getDriver().switchTo().window(newTabs.get(0));
		 	Thread.sleep(2000);
		 		return false;
		}
	 	
	}

@FindBy(xpath="//*[@id='divBtnGroup2']")
public static WebElement detailsTab;


@FindBy(xpath="//*[@id='sEMail']")
public static WebElement emailTxt;

@FindBy(xpath="(//*[@class='icon-save hiconright2'])[1]")
public static WebElement saveBtn;



public static boolean checkSavingPurchaseVoucherVATforVendorEmail() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
{
	Thread.sleep(2000);
	
	checkEmailtoVendorAcc();
	Thread.sleep(4000);
	
	getDriver().navigate().refresh();
	Thread.sleep(4000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
	purchaseVouchersVat.click();

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();

	checkValidationMessage("Screen opened");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
	vendorAccountTxt.click();

	vendorAccountTxt.sendKeys(Keys.SPACE);	

	ArrayList<String>  actVendorAccountList= new ArrayList<String>(); 

	int count=vendorAccountListCount.size();

	System.err.println(count);

	for(int i=0 ; i < count ;i++)
	{
		String data=vendorAccountListCount.get(i).getText();

		if (data.equalsIgnoreCase("Vendor B")) 
		{
			vendorAccountListCount.get(i).click();
			break;
		}
	}
	
	Thread.sleep(2000);

	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
	vendorAccountTxt.click();
	Thread.sleep(2000);
	vendorAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	Thread.sleep(2000);;
	

	Robot robot = new Robot();
	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_E);
	robot.keyRelease(KeyEvent.VK_CONTROL);
	robot.keyRelease(KeyEvent.VK_E);

	Thread.sleep(4000);
	
	robot.keyPress(KeyEvent.VK_ESCAPE);
	robot.keyRelease(KeyEvent.VK_ESCAPE);
	robot.keyPress(KeyEvent.VK_ESCAPE);
	robot.keyRelease(KeyEvent.VK_ESCAPE);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
	vendorAccountTxt.click();

	vendorAccountTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	vendorAccountTxt.sendKeys("Vendor B");
	Thread.sleep(1000);
	
	
	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_E);
	robot.keyRelease(KeyEvent.VK_CONTROL);
	robot.keyRelease(KeyEvent.VK_E);

	Thread.sleep(4000);
	
	robot.keyPress(KeyEvent.VK_ESCAPE);
	robot.keyRelease(KeyEvent.VK_ESCAPE);
	robot.keyPress(KeyEvent.VK_ESCAPE);
	robot.keyRelease(KeyEvent.VK_ESCAPE);
	robot.keyPress(KeyEvent.VK_ESCAPE);
	robot.keyRelease(KeyEvent.VK_ESCAPE);
	Thread.sleep(2000);
	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_E);
	robot.keyRelease(KeyEvent.VK_CONTROL);
	robot.keyRelease(KeyEvent.VK_E);

	Thread.sleep(10000);
	
	//vendorAccountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(5000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(detailsTab));
	detailsTab.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTxt));
	emailTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	emailTxt.sendKeys("emailvalidationone@gmail.com");
	Thread.sleep(1000);
	emailTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	Thread.sleep(2000);*/
	
	vendorAccountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(3000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
	voucherHeaderDueDate.click();
	voucherHeaderDueDate.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	voucherHeaderCurrency.sendKeys(Keys.SPACE);

	int currencycount=currencyListCount.size();

	System.err.println(currencycount);

	for(int i=0 ; i < currencycount ;i++)
	{
		String data=currencyListCount.get(i).getText();

		if (data.equalsIgnoreCase("INR")) 
		{
			currencyListCount.get(i).click();
			break;
		}
	}

	voucherHeaderCurrency.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	departmentTxt.click();
	departmentTxt.sendKeys(Keys.SPACE);	

	ArrayList<String> actDepartmentList= new ArrayList<String>(); 

	int departmentCount=departmentListCount.size();

	System.err.println(departmentCount);

	for(int i=0 ; i < departmentCount ;i++)
	{
		String data=departmentListCount.get(i).getText();

		if (data.equalsIgnoreCase("INDIA")) 
		{
			departmentListCount.get(i).click();
			break;
		}
	}

	departmentTxt.sendKeys(Keys.TAB);
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
	placeOFSupplyTxt.click();
	placeOFSupplyTxt.clear();
	placeOFSupplyTxt.sendKeys("Abu Dhabi");
	Thread.sleep(2000);
	placeOFSupplyTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
	jurisdictionTxt.click();
	jurisdictionTxt.sendKeys("DUBAI");
	Thread.sleep(2000);
	jurisdictionTxt.sendKeys(Keys.TAB);
	
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
	enter_WarehouseTxt.click();

	enter_WarehouseTxt.sendKeys(Keys.SPACE);

	int warehousecount=warehouseBodyComboList.size();

	for(int i=0 ; i < warehousecount ;i++)
	{
		String data=warehouseBodyComboList.get(i).getText();

		if (data.equalsIgnoreCase("HYDERABAD")) 
		{
			warehouseBodyComboList.get(i).click();
			break;
		}
	}

	enter_WarehouseTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.click();
	enter_ItemTxt.sendKeys(Keys.SPACE);
	int pvvGridItemListCount=pvvGridItemList.size();
	for (int i = 0; i < pvvGridItemListCount; i++) 
	{
		String Item=pvvGridItemList.get(i).getText();
		if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) 
		{
			pvvGridItemList.get(i).click();
			break;
		}
	}
	enter_ItemTxt.sendKeys(Keys.TAB);	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
	enter_TaxCode.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
	enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
	select1stRow_9thColumn.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	enter_Quantity.click();
	enter_Quantity.clear();
	enter_Quantity.sendKeys("2");


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	select1stRow_11thColumn.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.click();
	enter_Rate.clear();
	enter_Rate.sendKeys("10");
	enter_Rate.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
	enter_Gross.click();
	enter_Gross.sendKeys(Keys.TAB);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));
	select1stRow_13thColumn.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
	enter_Discount.click();
	enter_Discount.clear();
	enter_Discount.sendKeys("10");
	enter_Discount.sendKeys(Keys.TAB);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	select1stRow_14thColumn.click();

	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
	enter_PvVat.sendKeys(Keys.TAB);*/

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
	enter_PvTaxable.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
	enter_Batch.click();
	enter_Batch.sendKeys("PVATR1#V1");
	enter_Batch.sendKeys(Keys.TAB);

	DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

	Calendar cal=Calendar.getInstance();
	SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
	cal.add(Calendar.DATE, 2); 

	String next2DaysDate=df.format(cal.getTime());

	System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "+next2DaysDate);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
	enter_Expirydate.click();
	enter_Expirydate.sendKeys(Keys.HOME);
	enter_Expirydate.sendKeys(next2DaysDate);
	enter_Expirydate.sendKeys(Keys.TAB);

	Thread.sleep(2000);

	*/

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
	voucherSaveBtn.click();
	
	Thread.sleep(3000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
	billRefNewReferenceTxt.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
	pickBtn.click();
	
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
	Bill_OkBtn.click();
	
	Thread.sleep(2000);

			
	String expMessage1 = "Voucher saved successfully";
	
	String actMessage = checkValidationMessage(expMessage1);
	String expMessage2 = ": SU/IND/TEXT8";
	
	
	System.out.println("Actual Message    : "+actMessage);
	System.out.println("Expected Message  : "+expMessage1);

	if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) )
	{
		return true;
	} 
	else 
	{
		return false;
	}

}


public static boolean checkCreateEmailTemplateforVendor() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	Thread.sleep(2000);
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	System.out.println("Setting buton is enabled");
	serachMenuTextHomePage.click();
	serachMenuTextHomePage.sendKeys("Configure Transactions");
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	searchMenuTextClick.click();	  
	
	Thread.sleep(6000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
	purchaseVouchersVATBtn.click();
					 	    	
	Thread.sleep(8000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
	triggersBtn.click();
	Thread.sleep(4000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	editLayoutTriggerName.click();
	editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	Thread.sleep(1000);
	editLayoutTriggerName.sendKeys("Map Trigger");
	Thread.sleep(4000);
	editLayoutTriggerName.sendKeys(Keys.TAB);
	Thread.sleep(8000);
	

	getAction().moveToElement(sendEmailChkBox).build().perform();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));

	if(sendEmailChkBox.isSelected()==false)
	{
	sendEmailChkBoxSelected.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailTxtBox));
	sendEmailTxtBox.click();
	sendEmailTxtBox.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	sendEmailTxtBox.sendKeys("TriggersEmailLayout");
	Thread.sleep(3000);
	}
	
	
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginChkBoxSelected));
	if(loginChkBox.isSelected()==true)
	{
		loginChkBoxSelected.click();
	}

	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerChkBoxSelected));
	if(customerChkBox.isSelected()==false)
	{
		customerChkBoxSelected.click();
	}
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
	editLayoutTriggersSaveBtn.click();
	Thread.sleep(2500);
	
	String expSaveMessage = "Data saved successfully";
	String actSaveMessage = checkValidationMessage(expSaveMessage);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
	updateBtn.click();
	
	
	String expUpdateMessage = "Data saved successfully";
	String actUpdateMessage = checkValidationMessage(expUpdateMessage);
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
	settings_closeBtn.click();
	
	
	if( actSaveMessage.equalsIgnoreCase(expSaveMessage))
	{
		return true;
	}
	else
	{
		return false;
	}



}
	
public static boolean checkEmailValidationforVendorinCreatedTriggerTemplate() throws InterruptedException, AWTException, IOException
{
	Thread.sleep(4000);
	
	File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\TriggersEmailVendor.pdf");
	
	if(Efile1.exists())
	{
		Efile1.delete();
	}
	
		Robot robot = new Robot();   
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);
		
		Thread.sleep(4000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

	 	System.out.println("openTabs"+openTabs);

	 	getDriver().switchTo().window(openTabs.get(0));
	 	Thread.sleep(2000);
	 	
	 	getDriver().switchTo().window(openTabs.get(1));
	 	Thread.sleep(2000);
	
	 	getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");
	 	
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
	 	userNameTxt.click();
	 	
	 	userNameTxt.sendKeys("emailvalidationone@gmail.com");
	 	
	 	Thread.sleep(2000);
	 	
	 	System.out.println("User enter text");
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
	 	NextBtn.click();
	 	
	 	Thread.sleep(2000);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
	 	PasswordTxt.click();
	 	
	 	PasswordTxt.sendKeys("validationone");
	 	
	 	Thread.sleep(2000);
	 	
	 	NextBtn.click();
	 	
	 	Thread.sleep(15000);

	 	int count1 = row1mailFromList.size();
	 	//System.out.println("No.of mails in the list"+count1);
	 	
	 	//System.err.println("Email Body *"+emailBody1.getText()+"*");
	 	
	 	if (emailBody2.getText().isEmpty()==false) 
	 	{
	 		int count = emailSubjectList.size();
	 		System.err.println("No.of mails in the list"+count);
		 	
		 	for (int i = 0; i < count; i++) 
	 	{
	 		String mailFromTxt = emailSubjectList.get(i).getText();
	 		
	 		if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher")) 
	 		{
	 			emailSubjectList.get(i).click();
	 			break;
			}
		}
		 	
		 	
		 	Thread.sleep(3000);
 	
 	Thread.sleep(3000);
 	
 	getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
 	
 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
 	mailAttachmentDownloadBtn1.click();
 	
 	Thread.sleep(3000);
 	
 	 
 	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_CONTROL);
	
	
	Thread.sleep(5000);
	
	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);
	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);
	
	robot.keyPress(KeyEvent.VK_ENTER);
	robot.keyRelease(KeyEvent.VK_ENTER);
	
	Thread.sleep(3000);
	
	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_CONTROL);
 	
	Thread.sleep(2000);
	
			
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\TriggersEmailVendor.exe");
			
			Thread.sleep(4000);
			
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(2000);
			
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
			int actOpenWindowsCount = getDriver().getWindowHandles().size();
			int expOpenWindowsCount = 4;
			
			System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
			
			getDriver().switchTo().window(newTabs.get(3)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(2)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(1));
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 	mailDeleteButton.click();
		 	
		 	Thread.sleep(6000);
		 	

		 	
		 	
		 /*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailUserBtn));
		 	gmailUserBtn.click();*/
		 	ClickUsingJs(gmailUserBtn);
		 	
		 	Thread.sleep(5000);
		 	getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(8000);
			
			
		 	//getAction().moveToElement(gmailSignOutBtn).click().build().perform();
			//ScrollIntoView(gmailSignOutBtn);
			
		
			//getAction().moveToElement(gmailSignOutBtn).build().perform();
			ClickUsingJs(gmailSignOutBtn);
		
			Thread.sleep(2000);
			
			getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
			
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 	removeAccountBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 	removeDeleteBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 	yesRemoveBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
				 
		 	getDriver().switchTo().window(newTabs.get(1)).close();
		 	Thread.sleep(2000);
		 	getDriver().switchTo().window(newTabs.get(0));
		 	Thread.sleep(2000);
		 
		 	
		 	String actAccount1PDF = getBaseDir()+"\\autoIt\\ExportFiles\\TriggersEmailVendor.pdf";
			String expAccount1PDF = getBaseDir()+"\\autoIt\\ImportFiles\\TriggersEmailVendor.pdf";
			
			PDFUtil pdfutil = new PDFUtil();
			
			
			
			boolean result = pdfutil.compare(actAccount1PDF, expAccount1PDF);
			Calendar cal=Calendar.getInstance();
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = df.format(cal.getTime());
			
			String actAccount1Data = pdfutil.getText(actAccount1PDF);
			String expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("17/12/2024", currentDate);
					
			System.err.println(actAccount1Data);
			System.err.println(expAccount1Data);
			
		
		 	
			
			if (actAccount1Data.equalsIgnoreCase(expAccount1Data)) 
			{
			
				System.err.println("PDF file is as Expected");
				return true;
			}
			else
			{
									
				System.err.println("PDF file is not as Expected");
				return false;
			}
			
			
		} 
		 	else 
		{	
		 		
		 		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
		 		
		 		
		 	Thread.sleep(2000);
	
		 	ClickUsingJs(gmailUserBtn);
		 	
		 	Thread.sleep(3000);
		 getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(1000);
		 	getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 	Thread.sleep(2000);
		 	
		 	getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 	removeAccountBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 	removeDeleteBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 	yesRemoveBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
			getDriver().switchTo().window(newTabs.get(1)).close();
		 	Thread.sleep(2000);
		 	getDriver().switchTo().window(newTabs.get(0));
		 	Thread.sleep(2000);
		 	
		 		return false;
		}
	 	
	}


@FindBy(xpath="//*[@id='trigger_Controlbuttons']/a[1]/i")
private static WebElement  editLayoutTriggerDeleteBtn;



public void checkDeletingTriggerInPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(5000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	editLayoutTriggerName.click();
	editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);;
	editLayoutTriggerName.sendKeys("Test Trigger");
	Thread.sleep(4000);
	editLayoutTriggerName.sendKeys(Keys.TAB);
	
	Thread.sleep(8000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSuspendedChkboxsSelected));
	boolean acttriggersSuspendedChkboxs = triggersSuspendedChkboxs.isSelected();
	boolean exptriggersSuspendedChkboxs = true;
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerDeleteBtn));
	editLayoutTriggerDeleteBtn.click();
	Thread.sleep(2000);
	
	if (getIsAlertPresent())
	{
		getWaitForAlert();
		
		System.out.println("Deleted");
		
		getAlert().accept();
		Thread.sleep(4000);
	}
	else
	{
		editLayoutTriggerDeleteBtn.click();
		
		getWaitForAlert();
		System.out.println("Deleted");
		
		getAlert().accept();
		
		Thread.sleep(4000);
	}
	
		
}


@FindBy(xpath="(//a[@id='3333'])[2]")
public static WebElement salesInvoiceVATBtn;

	
@FindBy (xpath="//table[@id='editScreen_bodyTable']/tbody/tr[7]/td/span[2]")
private static WebElement editScreenBodyRow7Delete;

@FindBy(xpath="//i[@class='icon-pick icon-font5']")
private static WebElement batchPickOnFIFOIcon;

@FindBy(xpath="//label[@id='div_Ok']")
private static WebElement batchOkIcon;




public boolean checkDeleteVATFieldinEditScreenofSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	System.out.println("Setting buton is enabled");
	serachMenuTextHomePage.click();
	serachMenuTextHomePage.sendKeys("Configure Transactions");
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	searchMenuTextClick.click();	  
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATBtn));
	salesInvoiceVATBtn.click();
					 	    	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
	editScreenTab.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyTab));
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyRow7Delete));
	editScreenBodyRow7Delete.click();
	
	getWaitForAlert();
	Thread.sleep(2000);
	getAlert().accept();
	
	System.out.println("*******************************************checkDeleteFieldInEditScreenBodyInTrialPurchaseUpdate*******************************************");
		
	String expMessage = "Data deleted successfully";
	String actMessage = checkValidationMessage(expMessage);
	if(actMessage.equalsIgnoreCase(expMessage) )
	{	
		return  true;
	}
	else
	{
		return  false;
	}
	
		
}
		
public boolean checkSaveTriggerwithDiscountasMapFieldinSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
						 	    	
	Thread.sleep(2000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
	triggersBtn.click();
	
			
	Thread.sleep(4000);
	
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	editLayoutTriggerName.click();
	editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
	Thread.sleep(1000);
	editLayoutTriggerName.sendKeys("Sales Map Trigger");
	Thread.sleep(2000);
	editLayoutTriggerName.sendKeys(Keys.TAB);
	Thread.sleep(5000);
	
	triggersConjuctionDrpdwn.click();
    Thread.sleep(2000);		
	triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);			
	triggersConjuctionDrpdwn.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
	triggersSelectFieldTxt.click();
				
	Thread.sleep(3000);
	
	getAction().moveToElement(triggerVendorACExpansionBtn).build().perform();
	Thread.sleep(1000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorACExpansionBtn));
	triggerVendorACExpansionBtn.click();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerCustomerAcName));
	triggerCustomerAcName.click();
	
	Thread.sleep(2000);

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
	triggersSelectOperdrpdwn.click();
	triggersSelectOperdrpdwn.sendKeys("Equal");
	
	
	Thread.sleep(5000);
	triggersSelectOperdrpdwn.sendKeys(Keys.TAB);
	
	Thread.sleep(5000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
	triggersCompareWithdrpdwn.click();
	triggersCompareWithdrpdwn.sendKeys("value");
	//triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
	Thread.sleep(2000);
	triggersCompareWithdrpdwn.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	
	triggersValueTxt.sendKeys(Keys.SPACE);
	Thread.sleep(3000);
	triggersValueTxt.sendKeys("Customer A");
	Thread.sleep(1500);
	triggersValueTxt.sendKeys(Keys.TAB);

	
	Thread.sleep(1500);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
	if(editLayoutTriggerRaiseDoc.isSelected()==false)
	{
		editLayoutTriggerRaiseDocSelected.click();
	}
	
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
	triggerTabAlertBtn.click();
	
	Thread.sleep(2000);*/
	
	ScrollIntoView(editLayoutTriggersAllVouchers);
	Thread.sleep(1000);
	
						
	getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
	Select s=new Select(editLayoutTriggersAllVouchers);
	s.selectByVisibleText("Cash Sales");
	Thread.sleep(3000);
	
	
	getAction().moveToElement(trigger1stRow1stCol).build().perform();
	
	
	ClickUsingJs(trigger1stRow1stCol);
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectFieldDropdown));

	Thread.sleep(1000);
	Select s1=new Select(triggerSelectFieldDropdown);
	s1.selectByVisibleText("Discount");
	Thread.sleep(3000);
	triggerSelectFieldDropdown.sendKeys(Keys.TAB);
	
	
	ClickUsingJs(trigger1stRow2ndCol);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectMapDropdown));
	Select s2=new Select(triggerSelectMapDropdown);
	s2.selectByVisibleText("Discount");
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
	editLayoutTriggersSaveBtn.click();
	Thread.sleep(2500);
	
	String expSaveMessage = "Data saved successfully";
	String actSaveMessage = checkValidationMessage(expSaveMessage);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
	updateBtn.click();
	
	
	String expUpdateMessage = "Data saved successfully";
	String actUpdateMessage = checkValidationMessage(expUpdateMessage);
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
	closeBtn.click();
		
	if( actSaveMessage.equalsIgnoreCase(expSaveMessage))
	{
		System.out.println("Test Pass : Trigger  Added Successfully");
		
		return true;
	}
	else
	{
		System.out.println("Test Fail : Trigger NOT Added Successfully");
		
		return false;
	}

		
}

@FindBy(xpath = "//input[@id='id_body_33554476']")
public static WebElement enter_AQTxt;

@FindBy(xpath = "//input[@id='id_body_33554477']")
public static WebElement enter_FQTxt;


@FindBy(xpath = "//tbody[@id='id_header_4_table_body']/tr/td[2]")
public static List<WebElement> customerAccountListCount;


@FindBy(xpath = "//tbody[@id='id_header_268435470_table_body']/tr/td[2]")
public static List<WebElement> placeoFSupplyList;

@FindBy(xpath = "//tbody[@id='id_header_268435470_table_body']/tr/td[2]")
public static List<WebElement> placeOFSupplyList;


@FindBy(xpath = "//tbody[@id='id_body_23_table_body']/tr/td[2]")
public static List<WebElement> itemListCount;


@FindBy(xpath = "//tbody[@id='id_body_536870916_table_body']/tr/td[2]")
public static List<WebElement> pvwareHouseListCount;

	
public static boolean checkSavingSalesInvoiceVATVoucher1withDiscountasMapField() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
	financialTransactionSalesMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
	salesInvoiceVATVoucher.click();

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();

	checkValidationMessage("Screen opened");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	customerAccountTxt.click();
	customerAccountTxt.sendKeys(Keys.END);
	customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	customerAccountTxt.sendKeys(Keys.SPACE);

	int customercount=customerAccountListCount.size();

	System.err.println(customercount);

	for(int i=0 ; i < customercount ;i++)
	{
		String data=customerAccountListCount.get(i).getText();

		if(data.equalsIgnoreCase("Customer A"))
		{
			customerAccountListCount.get(i).click();

			break;
		}
	}

	customerAccountTxt.sendKeys(Keys.TAB);


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	departmentTxt.click();
	departmentTxt.sendKeys(Keys.SPACE);

	int departmentcount=departmentListCount.size();

	System.err.println(departmentcount);

	for(int i=0 ; i < departmentcount ;i++)
	{
		String data=departmentListCount.get(i).getText();

		if(data.equalsIgnoreCase("DUBAI"))
		{
			departmentListCount.get(i).click();

			break;
		}
	}

	departmentTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
	salesInvoiceVATPlaceOFSupply.click();
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SPACE);


	int placeOFSupplyListCount=placeOFSupplyList.size();

	System.err.println("placeOFSupplyListCount   : "+placeOFSupplyListCount);

	for(int i=0 ; i < placeOFSupplyListCount ;i++)
	{
		String data=placeOFSupplyList.get(i).getText();

		if(data.equalsIgnoreCase("Abu Dhabi"))
		{
			placeOFSupplyList.get(i).click();

			break;
		}
	}

	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
	jurisdictionTxt.click();
	jurisdictionTxt.sendKeys(Keys.END);
	jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	jurisdictionTxt.sendKeys("DUBAI");
	Thread.sleep(2000);
	jurisdictionTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
	pvWareHouseTxt.click();
	pvWareHouseTxt.sendKeys(Keys.SPACE);

	int warehousecount=pvwareHouseListCount.size();

	System.err.println(warehousecount);

	for(int i=0 ; i < warehousecount ;i++)
	{
		String data=pvwareHouseListCount.get(i).getText();

		if(data.equalsIgnoreCase("SECUNDERABAD"))
		{
			pvwareHouseListCount.get(i).click();

			break;
		}	
	}

	pvWareHouseTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.sendKeys(Keys.END);
	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
	enter_ItemTxt.sendKeys(Keys.SPACE);

	int itemcount1=itemListCount.size();

	System.err.println(itemcount1);

	for(int i=0 ; i < itemcount1 ;i++)
	{
		String data=itemListCount.get(i).getText();

		if(data.equalsIgnoreCase("STD RATE COGS ITEM"))
		{
			itemListCount.get(i).click();

			break;
		}
	}

	enter_ItemTxt.sendKeys(Keys.TAB);

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	select1stRow_5thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
	enter_AQTxt.sendKeys("6");
	enter_AQTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
	enter_FQTxt.sendKeys("6");
	enter_FQTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	select1stRow_11thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	select1stRow_14thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.click();
	enter_Rate.clear();
	enter_Rate.sendKeys("10");
	enter_Rate.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
	enter_Gross.click();
	enter_Gross.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_16thColumn));
	select1stRow_16thColumn.click();	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SIVAT_enter_Discount));
	SIVAT_enter_Discount.click();
	SIVAT_enter_Discount.clear();
	SIVAT_enter_Discount.sendKeys("10");
	SIVAT_enter_Discount.sendKeys(Keys.TAB);
	Thread.sleep(2000);

	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
	batchPickOnFIFOIcon.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
	batchOkIcon.click();
*/
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String docno=documentNumberTxt.getAttribute("value");
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
	voucherSaveBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
	billRefNewReferenceTxt.click();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
	pickBtn.click();
	Thread.sleep(2000);
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
	Bill_OkBtn.click();
	checkValidationMessage("");


	Thread.sleep(3000);
	/*String expMsg="This Transaction will make the Stock Negative";
	
	String actMsg=checkValidationMessage(expMsg);
	Thread.sleep(2000);*/
	checkValidationMessage("");
	
	String expMessage1 = "Voucher saved successfully";

	String actMessage = checkValidationMessage(expMessage1);
	String expMessage2 = ": "+docno;

			
	System.out.println("Actual Message    : "+actMessage);
	System.out.println("Expected Message  : "+expMessage1);

	if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) )

	
	{
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		Thread.sleep(2000);*/
		return true;
	}	
	else
	{
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		Thread.sleep(2000);*/
		return false;
	}

	
	
	
}


@FindBy(xpath = "//*[@id='id_body_33554473']")
public static WebElement SIVAT_enter_Discount;

public static  ArrayList<String> checkNetValueofVoucher1inSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(1000);
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	documentNumberTxt.click();
	documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	documentNumberTxt.sendKeys("1");
	Thread.sleep(2000);
	documentNumberTxt.sendKeys(Keys.TAB);

	checkValidationMessage("Voucher loaded successfully");

	ArrayList<String> DiscNetArry = new ArrayList<String>();
	String actDiscount = select1stRow_16thColumn.getText();
	
	
	String actNetValue= voucherFooterNetValue.getText();
	
	DiscNetArry.add(actDiscount);
	DiscNetArry.add(actNetValue);
	
	
	return DiscNetArry;
	
}


public boolean checkCashSalesVoucherafterTriggerOptionsSavewithMapField() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	ArrayList<String> DiscNetArry = new ArrayList<String>();
	DiscNetArry=checkNetValueofVoucher1inSalesInvoiceVAT();
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	System.out.println("Setting buton is enabled");
	serachMenuTextHomePage.click();
	serachMenuTextHomePage.sendKeys("Cash Sales");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	searchMenuTextClick.click();

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	
	int voucherGridDocNoCount = voucherGridDocNo.size();

	for(int i=0;i<voucherGridDocNoCount;i++)
	{
		String data ="1";
		if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
		{
			voucherGridIndexChkBox.get(i).click();
			break;
		}
	}
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
	editBtn.click();
	Thread.sleep(2000);
	
	
	
	String actDiscountValue=select1stRow_6thColumn.getText();
	String expDiscountValue=DiscNetArry.get(0);
	
	String actNetValue=voucherFooterNetValue.getText();
	
	String expNetValue=DiscNetArry.get(1);
	
	System.out.println("Actual Discount    : "+actDiscountValue);
	System.out.println("Expected Discount  : "+expDiscountValue);
	
	System.out.println("Actual Net   : "+actNetValue);
	System.out.println("ExpectedNet  : "+expNetValue);
	
	
	
	if(actDiscountValue.equalsIgnoreCase(expDiscountValue)&&actNetValue.equalsIgnoreCase(expNetValue))
	{
		return true;
	}
	else
	{
		return false;
	}
				
}


public static boolean checkCreatingAlertinTriggersSalesInvoiceVAT() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	
	serachMenuTextHomePage.click();
	serachMenuTextHomePage.sendKeys("Configure Transactions");
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	searchMenuTextClick.click();	  
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATBtn));
	salesInvoiceVATBtn.click();
					 	    	
	Thread.sleep(2000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
	triggersBtn.click();
		
	Thread.sleep(4000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	editLayoutTriggerName.click();
	editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
	Thread.sleep(1000);
	editLayoutTriggerName.sendKeys("Sales Map Trigger");
	Thread.sleep(2000);
	editLayoutTriggerName.sendKeys(Keys.TAB);
	Thread.sleep(5000);
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendAlertChkBox));
	sendAlertChkBox.click();*/
	
	/*if(SendAlertChkBox.isDisplayed()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
	triggerTabAlertBtn.click();
	}
	Thread.sleep(2000);*/
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cretaNewAlertButton));
	cretaNewAlertButton.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newButtonInCreateTemplate));
	newButtonInCreateTemplate.click();
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplates));
	emailTemplates.sendKeys("SalesMapAlertTriggerTemplate");
	emailTemplates.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
	companyNameTemplateArea.click();
	Thread.sleep(1500);
	

	Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\salesTriggerAlert.exe");
	
	Thread.sleep(80000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
	saveBtnInEmailTemplateHeader.click();
	
	Thread.sleep(2000);
	System.err.println("Template save button clicked");
	
	String expEmailTemplateSaveMsg = "Template saved successfully";
	String actEmailTemplateSaveMsg =  checkValidationMessage(expEmailTemplateSaveMsg);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
	closeButtonInCreateTemplate.click();
	Thread.sleep(2000);
	if(actEmailTemplateSaveMsg.equalsIgnoreCase(expEmailTemplateSaveMsg))
	{
		return true;
	}
	else
	{
		return false;
	}
	
		 
}


public static boolean checkCreatedAlertandUserSelectinSalesInvoiceVATVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	System.out.println("Setting buton is enabled");
	serachMenuTextHomePage.click();
	serachMenuTextHomePage.sendKeys("Configure Transactions");
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	searchMenuTextClick.click();	  
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATBtn));
	salesInvoiceVATBtn.click();
					 	    	
	Thread.sleep(2000);
	
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
	triggersBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	editLayoutTriggerName.click();
	editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
	Thread.sleep(1000);
	editLayoutTriggerName.sendKeys("Sales Map Trigger");
	Thread.sleep(2000);
	editLayoutTriggerName.sendKeys(Keys.TAB);
	Thread.sleep(5000);
	/*
	if(SendAlertChkBox.isDisplayed()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
	triggerTabAlertBtn.click();
	}
	Thread.sleep(2000);*/
	
	
	
	
	if(sendAlertChkBox.isSelected()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SendAlertChkBoxSelected));
	SendAlertChkBoxSelected.click();
	}
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectSendAlert));
	Select s=new Select(selectSendAlert);
	s.selectByVisibleText("SalesMapAlertTriggerTemplate");
	Thread.sleep(2000);
	String actMsg=s.getFirstSelectedOption().getText();
	String expMsg="SalesMapAlertTriggerTemplate";
	
	System.out.println("Actual Message: 	" 		+ actMsg);
	System.out.println("Expected Message:  	"	+expMsg);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginChkBoxSelected));
	if(loginChkBox.isSelected()==false)
	{
		loginChkBoxSelected.click();
	}
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectUsersList));
	Select s1=new Select(selectUsersList);
	s1.selectByVisibleText("UserAllOptionsST");
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
	editLayoutTriggersSaveBtn.click();
	Thread.sleep(2500);
	
	String expSaveMessage = "Data saved successfully";
	String actSaveMessage = checkValidationMessage(expSaveMessage);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
	updateBtn.click();
	
	
	String expUpdateMessage = "Data saved successfully";
	String actUpdateMessage = checkValidationMessage(expUpdateMessage);
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
	closeBtn.click();
	
	
	if(actMsg.equalsIgnoreCase(expMsg) && actSaveMessage.equalsIgnoreCase(expSaveMessage))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}


public static boolean checkSavingVoucher2inSalesInvoiceVATforAlertMessage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
	financialTransactionSalesMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
	salesInvoiceVATVoucher.click();

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();

	checkValidationMessage("Screen opened");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	customerAccountTxt.click();
	customerAccountTxt.sendKeys(Keys.END);
	customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	customerAccountTxt.sendKeys(Keys.SPACE);

	int customercount=customerAccountListCount.size();

	System.err.println(customercount);

	for(int i=0 ; i < customercount ;i++)
	{
		String data=customerAccountListCount.get(i).getText();

		if(data.equalsIgnoreCase("Customer A"))
		{
			customerAccountListCount.get(i).click();

			break;
		}
	}

	customerAccountTxt.sendKeys(Keys.TAB);


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	departmentTxt.click();
	departmentTxt.sendKeys(Keys.SPACE);

	int departmentcount=departmentListCount.size();

	System.err.println(departmentcount);

	for(int i=0 ; i < departmentcount ;i++)
	{
		String data=departmentListCount.get(i).getText();

		if(data.equalsIgnoreCase("DUBAI"))
		{
			departmentListCount.get(i).click();

			break;
		}
	}

	departmentTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
	salesInvoiceVATPlaceOFSupply.click();
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SPACE);


	int placeOFSupplyListCount=placeOFSupplyList.size();

	System.err.println("placeOFSupplyListCount   : "+placeOFSupplyListCount);

	for(int i=0 ; i < placeOFSupplyListCount ;i++)
	{
		String data=placeOFSupplyList.get(i).getText();

		if(data.equalsIgnoreCase("Abu Dhabi"))
		{
			placeOFSupplyList.get(i).click();

			break;
		}
	}

	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
	jurisdictionTxt.click();
	jurisdictionTxt.sendKeys(Keys.END);
	jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	jurisdictionTxt.sendKeys("DUBAI");
	Thread.sleep(2000);
	jurisdictionTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
	pvWareHouseTxt.click();
	pvWareHouseTxt.sendKeys(Keys.SPACE);

	int warehousecount=pvwareHouseListCount.size();

	System.err.println(warehousecount);

	for(int i=0 ; i < warehousecount ;i++)
	{
		String data=pvwareHouseListCount.get(i).getText();

		if(data.equalsIgnoreCase("SECUNDERABAD"))
		{
			pvwareHouseListCount.get(i).click();

			break;
		}	
	}

	pvWareHouseTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.sendKeys(Keys.END);
	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
	enter_ItemTxt.sendKeys(Keys.SPACE);

	int itemcount1=itemListCount.size();

	System.err.println(itemcount1);

	for(int i=0 ; i < itemcount1 ;i++)
	{
		String data=itemListCount.get(i).getText();

		if(data.equalsIgnoreCase("STD RATE COGS ITEM"))
		{
			itemListCount.get(i).click();

			break;
		}
	}

	enter_ItemTxt.sendKeys(Keys.TAB);

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	select1stRow_5thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
	enter_AQTxt.sendKeys("6");
	enter_AQTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
	enter_FQTxt.sendKeys("6");
	enter_FQTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	select1stRow_11thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	select1stRow_14thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.click();
	enter_Rate.clear();
	enter_Rate.sendKeys("10");
	enter_Rate.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
	enter_Gross.click();
	enter_Gross.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_16thColumn));
	select1stRow_16thColumn.click();	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SIVAT_enter_Discount));
	SIVAT_enter_Discount.click();
	SIVAT_enter_Discount.clear();
	SIVAT_enter_Discount.sendKeys("5");
	SIVAT_enter_Discount.sendKeys(Keys.TAB);
	Thread.sleep(2000);

	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
	batchPickOnFIFOIcon.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
	batchOkIcon.click();
*/
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String docno=documentNumberTxt.getAttribute("value");
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
	voucherSaveBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
	billRefNewReferenceTxt.click();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
	pickBtn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
	Bill_OkBtn.click();

	Thread.sleep(3000);
	checkValidationMessage("");
	
	String expMessage1 = "Voucher saved successfully";

	String actMessage = checkValidationMessage(expMessage1);
	String expMessage2 = ": "+docno;

			
	System.out.println("Actual Message    : "+actMessage);
	System.out.println("Expected Message  : "+expMessage1);

	if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) )

	
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		Thread.sleep(2000);
		return true;
	}	
	else
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		Thread.sleep(2000);
		return false;
	}



}


public static boolean checkTriggerAlertMessageinUserAllOptionsSTofSalesInvoiceVAT() throws InterruptedException
{
	
	Thread.sleep(2000);
	getDriver().navigate().refresh();
	
	Thread.sleep(2000);
getWebDriverWait().until(ExpectedConditions.visibilityOf(noOfAlertsDisplay));
noOfAlertsDisplaySelcted.click();
Thread.sleep(2000);

// Actual
String  actAlertsMainHeader=alertsMainHeader.getText();

boolean actNoOfAlertsDisplay=alertsMainHeader.isDisplayed();
boolean expNoOfAlertsDisplay=true;

System.out.println("Alerts is display"+ actAlertsMainHeader);

Thread.sleep(2000);

int count = alertTransaction.size();

String alertTest=null;

for(int i=0; i<count; i++)
{
	String data = alertTransaction.get(i).getText();
	
	if(data.contains("  Transactions "))
	{
		alertTransaction.get(i).click();
		Thread.sleep(2000);
		transactionAuthorization2.get(i).click();
		Thread.sleep(2000);
		alertTest=transactionAuthorizationText2.get(i).getText();
		
		
	}
}


String actAlertMsg=alertTest;
String expAlertMsg="Automation Company : 08/10/2020 5:43 PMSUSECUNDERABAD-12";

System.out.println("Actual Alert Text: 		" 	+	alertTest);
System.out.println("Expected Alert Text: 	" 	+	expAlertMsg);

if(actNoOfAlertsDisplay==expNoOfAlertsDisplay && actAlertMsg.equalsIgnoreCase(expAlertMsg))
{
return true;
}
else
{
	return false;
}
}


public static boolean checkEmailTemplateCreationinSalesInvoiceVATTriggers() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	System.out.println("Setting buton is enabled");
	serachMenuTextHomePage.click();
	serachMenuTextHomePage.sendKeys("Configure Transactions");
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	searchMenuTextClick.click();	  
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATBtn));
	salesInvoiceVATBtn.click();
					 	    	
	Thread.sleep(2000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
	triggersBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	editLayoutTriggerName.click();
	editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
	Thread.sleep(1000);
	editLayoutTriggerName.sendKeys("Sales Map Trigger");
	Thread.sleep(2000);
	editLayoutTriggerName.sendKeys(Keys.TAB);
	Thread.sleep(5000);
	
	
	if(SendAlertChkBox.isSelected()==true)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SendAlertChkBoxSelected));
	SendAlertChkBoxSelected.click();
	}
	Thread.sleep(2000);
	
	
	
	
	if(sendEmailChkBox.isSelected()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
	sendEmailChkBoxSelected.click();
	}
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createNewEmailButton));
	createNewEmailButton.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailNewBtn));
	emailNewBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailNewLayoutBtn));
	emailNewLayoutBtn.click();
	Thread.sleep(2000);
	
	Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\SalesEmailTriggerNew2.exe");
	
	Thread.sleep(150000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveBtn));
	emailSaveBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveOptionBtn));
	emailSaveOptionBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSavePopupTitle));
	boolean actTitleDisplay=emailSavePopupTitle.isDisplayed();
	boolean expTitleDisplay=true;
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveTxt));
	emailSaveTxt.click();
	emailSaveTxt.sendKeys("SalesTriggersEmailLayout");
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveLayoutBtn));
	emailSaveLayoutBtn.click();
	Thread.sleep(2000);
	String expSaveLayotMsg="Layout saved successfully";
	String actSaveLayoutMsg=checkValidationMessage(expSaveLayotMsg);
	
	
	System.out.println("Actual Message: 	"	+actSaveLayoutMsg);
	System.out.println("Expected Message:	"	+expSaveLayotMsg);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailLayoutPreviewBtn));
	emailLayoutPreviewBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailLayoutDesignBtn));
	emailLayoutDesignBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailLayoutExitBtn));
	emailLayoutExitBtn.click();
	Thread.sleep(2000);
	
	if(actTitleDisplay==expTitleDisplay && actSaveLayoutMsg.equalsIgnoreCase(expSaveLayotMsg))
	{
		return true;
	}
	else
	{
		return false;
	}
	
	
}
	

public static boolean checkCreatedEmailTemplateinSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	System.out.println("Setting buton is enabled");
	serachMenuTextHomePage.click();
	serachMenuTextHomePage.sendKeys("Configure Transactions");
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	searchMenuTextClick.click();	  
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATBtn));
	salesInvoiceVATBtn.click();
					 	    	
	Thread.sleep(2000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
	triggersBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	editLayoutTriggerName.click();
	editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
	Thread.sleep(1000);
	editLayoutTriggerName.sendKeys("Sales Map Trigger");
	Thread.sleep(2000);
	editLayoutTriggerName.sendKeys(Keys.TAB);
	Thread.sleep(6000);
	
	/*if(SendAlertChkBox.isDisplayed()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
	triggerTabAlertBtn.click();
	}
	Thread.sleep(2000);*/
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SendAlertChkBoxSelected));
	if(sendAlertChkBox.isSelected()==true)
	{
		SendAlertChkBoxSelected.click();
	}
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
	if(sendEmailChkBox.isSelected()==false)
	{
	sendEmailChkBoxSelected.click();
	}
	Thread.sleep(2000);
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailTxtBox));
	sendEmailTxtBox.click();
	sendEmailTxtBox.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	sendEmailTxtBox.sendKeys("SalesTriggersEmailLayout");
	Thread.sleep(3000);
	sendEmailTxtBox.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginChkBoxSelected));
	if(loginChkBox.isSelected()==false)
	{
		loginChkBoxSelected.click();
	}
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectUsersList));
	Select s1=new Select(selectUsersList);
	s1.selectByVisibleText("TriggerUser");
	Thread.sleep(2000);
	s1.selectByVisibleText("UserAllOptionsST");
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
	editLayoutTriggersSaveBtn.click();
	Thread.sleep(2500);
	
	String expSaveMessage = "Data saved successfully";
	String actSaveMessage = checkValidationMessage(expSaveMessage);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
	updateBtn.click();
	
	
	String expUpdateMessage = "Data saved successfully";
	String actUpdateMessage = checkValidationMessage(expUpdateMessage);
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
	closeBtn.click();
	
	
	if(actSaveMessage.equalsIgnoreCase(expSaveMessage))
	{
		return true;
	}
	else
	{
		return false;
	}

	
}


public static boolean checkSavingVoucher3inSalesInvoiceVATforEmailValidation() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
	financialTransactionSalesMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
	salesInvoiceVATVoucher.click();

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();

	checkValidationMessage("Screen opened");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	customerAccountTxt.click();
	customerAccountTxt.sendKeys(Keys.END);
	customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	customerAccountTxt.sendKeys(Keys.SPACE);

	int customercount=customerAccountListCount.size();

	System.err.println(customercount);

	for(int i=0 ; i < customercount ;i++)
	{
		String data=customerAccountListCount.get(i).getText();

		if(data.equalsIgnoreCase("Customer A"))
		{
			customerAccountListCount.get(i).click();

			break;
		}
	}

	customerAccountTxt.sendKeys(Keys.TAB);


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	departmentTxt.click();
	departmentTxt.sendKeys(Keys.SPACE);

	int departmentcount=departmentListCount.size();

	System.err.println(departmentcount);

	for(int i=0 ; i < departmentcount ;i++)
	{
		String data=departmentListCount.get(i).getText();

		if(data.equalsIgnoreCase("DUBAI"))
		{
			departmentListCount.get(i).click();

			break;
		}
	}

	departmentTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
	salesInvoiceVATPlaceOFSupply.click();
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SPACE);


	int placeOFSupplyListCount=placeOFSupplyList.size();

	System.err.println("placeOFSupplyListCount   : "+placeOFSupplyListCount);

	for(int i=0 ; i < placeOFSupplyListCount ;i++)
	{
		String data=placeOFSupplyList.get(i).getText();

		if(data.equalsIgnoreCase("Abu Dhabi"))
		{
			placeOFSupplyList.get(i).click();

			break;
		}
	}

	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
	jurisdictionTxt.click();
	jurisdictionTxt.sendKeys(Keys.END);
	jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	jurisdictionTxt.sendKeys("DUBAI");
	Thread.sleep(2000);
	jurisdictionTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
	pvWareHouseTxt.click();
	pvWareHouseTxt.sendKeys(Keys.SPACE);

	int warehousecount=pvwareHouseListCount.size();

	System.err.println(warehousecount);

	for(int i=0 ; i < warehousecount ;i++)
	{
		String data=pvwareHouseListCount.get(i).getText();

		if(data.equalsIgnoreCase("SECUNDERABAD"))
		{
			pvwareHouseListCount.get(i).click();

			break;
		}	
	}

	pvWareHouseTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.sendKeys(Keys.END);
	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
	enter_ItemTxt.sendKeys(Keys.SPACE);

	int itemcount1=itemListCount.size();

	System.err.println(itemcount1);

	for(int i=0 ; i < itemcount1 ;i++)
	{
		String data=itemListCount.get(i).getText();

		if(data.equalsIgnoreCase("STD RATE COGS ITEM"))
		{
			itemListCount.get(i).click();

			break;
		}
	}

	enter_ItemTxt.sendKeys(Keys.TAB);

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	select1stRow_5thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
	enter_AQTxt.sendKeys("6");
	enter_AQTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
	enter_FQTxt.sendKeys("6");
	enter_FQTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	select1stRow_11thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	select1stRow_14thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.click();
	enter_Rate.clear();
	enter_Rate.sendKeys("10");
	enter_Rate.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
	enter_Gross.click();
	enter_Gross.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_16thColumn));
	select1stRow_16thColumn.click();	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SIVAT_enter_Discount));
	SIVAT_enter_Discount.click();
	SIVAT_enter_Discount.clear();
	SIVAT_enter_Discount.sendKeys("5");
	SIVAT_enter_Discount.sendKeys(Keys.TAB);
	Thread.sleep(2000);

	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
	batchPickOnFIFOIcon.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
	batchOkIcon.click();
*/
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String docno=documentNumberTxt.getAttribute("value");
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
	voucherSaveBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
	billRefNewReferenceTxt.click();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
	pickBtn.click();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
	Bill_OkBtn.click();

	Thread.sleep(3000);
	
	checkValidationMessage("");
	Thread.sleep(2000);
	
	String expMessage1 = "Voucher saved successfully";

	String actMessage = checkValidationMessage(expMessage1);
	String expMessage2 = ": "+docno;

			
	System.out.println("Actual Message    : "+actMessage);
	System.out.println("Expected Message  : "+expMessage1);

	if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) )

	
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		Thread.sleep(2000);
		return true;
	}	
	else
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		Thread.sleep(2000);
		return false;
	}



}

public static boolean checkEmailValidationUserforCreatedSalesTriggerTemplate() throws InterruptedException, AWTException, IOException
{
	Thread.sleep(2000);
	File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\SalesTriggersEmail.pdf");
	
	if(Efile1.exists())
	{
		Efile1.delete();
	}
	
		Robot robot = new Robot();   
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);
		
		Thread.sleep(4000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

	 	System.out.println("openTabs"+openTabs);

	 	getDriver().switchTo().window(openTabs.get(0));
	 	
	 	getDriver().switchTo().window(openTabs.get(1));
		
	
	 	getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");
	 	
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
	 	userNameTxt.click();
	 	
	 	userNameTxt.sendKeys("emailvalidationone@gmail.com");
	 	
	 	Thread.sleep(2000);
	 	
	 	System.out.println("User enter text");
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
	 	NextBtn.click();
	 	
	 	Thread.sleep(2000);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
	 	PasswordTxt.click();
	 	
	 	PasswordTxt.sendKeys("validationone");
	 	
	 	Thread.sleep(2000);
	 	
	 	NextBtn.click();
	 	
	 	Thread.sleep(8000);

	 	int count1 = row1mailFromList.size();
	 //	System.out.println("No.of mails in the list"+count1);
	 	
	 //	System.err.println("Email Body *"+emailBody.getText()+"*");
	 	
	 	if (emailBody2.getText().isEmpty()==false) 
	 	{
	 		int count = emailSubjectList.size();
	 		System.err.println("No.of mails in the list"+count);
		 	
		 	for (int i = 0; i < count; i++) 
	 	{
	 		String mailFromTxt = emailSubjectList.get(i).getText();
	 		
	 		if (mailFromTxt.equalsIgnoreCase("Focus Document")) 
	 		{
	 			emailSubjectList.get(i).click();
	 			break;
			}
		}
		 	
		 	
		 	Thread.sleep(3000);
 	
 	Thread.sleep(3000);
 	
 	getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
 	
 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
 	mailAttachmentDownloadBtn1.click();
 	
 	Thread.sleep(3000);
 	
 	 
 	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_CONTROL);
	
	
	Thread.sleep(5000);
	
	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);
	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);
	
	robot.keyPress(KeyEvent.VK_ENTER);
	robot.keyRelease(KeyEvent.VK_ENTER);
	
	Thread.sleep(3000);
	
	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_CONTROL);
 	
	Thread.sleep(2000);
	
			
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\SalesTriggersEmail.exe");
			
			Thread.sleep(4000);
			
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(2000);
			
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
			int actOpenWindowsCount = getDriver().getWindowHandles().size();
			int expOpenWindowsCount = 4;
			
			System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
			
			getDriver().switchTo().window(newTabs.get(3)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(2)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(1));
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 	mailDeleteButton.click();
		 	
		 	Thread.sleep(2000);
		 	

		 	Thread.sleep(2000);
		 	
		 /*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailUserBtn));
		 	gmailUserBtn.click();*/
		 	ClickUsingJs(gmailUserBtn);
		 	
		 	Thread.sleep(3000);
		 	getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(1000);
		 	getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 	Thread.sleep(6000);
		 	getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 	removeAccountBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 	removeDeleteBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 	yesRemoveBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
				 
		 	getDriver().switchTo().window(newTabs.get(1)).close();
		 	Thread.sleep(2000);
		 	getDriver().switchTo().window(newTabs.get(0));
		 	Thread.sleep(2000);
		 
		 	
		 	String actAccount1PDF = getBaseDir()+"\\autoIt\\ExportFiles\\SalesTriggersEmail.pdf";
			String expAccount1PDF = getBaseDir()+"\\autoIt\\ImportFiles\\SalesTriggersEmail.pdf";
			
			PDFUtil pdfutil = new PDFUtil();
			
			
			
			boolean result = pdfutil.compare(actAccount1PDF, expAccount1PDF);
			Calendar cal=Calendar.getInstance();
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = df.format(cal.getTime());
			
			String actAccount1Data = pdfutil.getText(actAccount1PDF);
			String expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("26/12/2022", currentDate);
					
			System.err.println(actAccount1Data);
			System.err.println(expAccount1Data);
			
		
		 	
			
			if (actAccount1Data.equalsIgnoreCase(expAccount1Data)) 
			{
			
				System.err.println("PDF file is as Expected");
				return true;
			}
			else
			{
									
				System.err.println("PDF file is not as Expected");
				return false;
			}
			
			
		} 
		 	else 
		{	
		 		
		 		
		 		System.err.println("Subject : "+emailRow1Subject.getText());
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsChckbx));
		 		selectAllMailsChckbx.click();
		 		
		 		Thread.sleep(2000);
		 	
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsDltBtn));
		 		selectAllMailsDltBtn.click();
		 		
		 		Thread.sleep(4000);

		 	Thread.sleep(2000);
		 	
		 /*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailUserBtn));
		 	gmailUserBtn.click();*/
		 	ClickUsingJs(gmailUserBtn);
		 	
		 	Thread.sleep(3000);
		 getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(1000);
		 	getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 	Thread.sleep(6000);
		 	
		 	getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 	removeAccountBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 	removeDeleteBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 	yesRemoveBtn.click();
		 	
		 	Thread.sleep(2000);
		 		return false;
		}
	 	
	}

public static boolean checkCreateEmailTemplateforCustomer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	Thread.sleep(2000);
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	System.out.println("Setting buton is enabled");
	serachMenuTextHomePage.click();
	serachMenuTextHomePage.sendKeys("Configure Transactions");
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	searchMenuTextClick.click();	  
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATBtn));
	salesInvoiceVATBtn.click();
					 	    	
	Thread.sleep(2000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
	triggersBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	editLayoutTriggerName.click();
	editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
	Thread.sleep(1000);
	editLayoutTriggerName.sendKeys("SalesMap Trigger");
	Thread.sleep(2000);
	editLayoutTriggerName.sendKeys(Keys.TAB);
	Thread.sleep(5000);
	
/*	if(SendAlertChkBox.isDisplayed()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
	triggerTabAlertBtn.click();
	}
	Thread.sleep(2000);*/
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
	if(sendEmailChkBox.isSelected()==true)
	{
	sendEmailChkBoxSelected.click();
	}
	Thread.sleep(2000);
	
	if(sendEmailChkBox.isSelected()==false)
	{
	sendEmailChkBoxSelected.click();
	}
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailTxtBox));
	sendEmailTxtBox.click();
	sendEmailTxtBox.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
	sendEmailTxtBox.sendKeys("SalesTriggersEmailLayout");
	Thread.sleep(3000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerChkBoxSelected));
	if(customerChkBox.isSelected()==false)
	{
		customerChkBoxSelected.click();
	}
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
	editLayoutTriggersSaveBtn.click();
	Thread.sleep(2500);
	
	String expSaveMessage = "Data saved successfully";
	String actSaveMessage = checkValidationMessage(expSaveMessage);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
	updateBtn.click();
	
	
	String expUpdateMessage = "Data saved successfully";
	String actUpdateMessage = checkValidationMessage(expUpdateMessage);
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
	closeBtn.click();
	
	
	if( actSaveMessage.equalsIgnoreCase(expSaveMessage))
	{
		return true;
	}
	else
	{
		return false;
	}



}


public static boolean checkSavingVoucher4inSalesInvoiceVATforCustomerEmailValidation() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
{

	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialTransactionSalesMenu));
	financialTransactionSalesMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATVoucher));
	salesInvoiceVATVoucher.click();

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();

	checkValidationMessage("Screen opened");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	customerAccountTxt.click();
	customerAccountTxt.sendKeys(Keys.END);
	customerAccountTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	customerAccountTxt.sendKeys(Keys.SPACE);

	int customercount=customerAccountListCount.size();

	System.err.println(customercount);

	for(int i=0 ; i < customercount ;i++)
	{
		String data=customerAccountListCount.get(i).getText();

		if(data.equalsIgnoreCase("Customer A"))
		{
			customerAccountListCount.get(i).click();

			break;
		}
	}

	//customerAccountTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	customerAccountTxt.click();
	Thread.sleep(2000);
	
	Robot robot = new Robot();   
	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_E);
	robot.keyRelease(KeyEvent.VK_CONTROL);
	robot.keyRelease(KeyEvent.VK_E);
	
	customerAccountTxt.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(detailsTab));
	detailsTab.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTxt));
	emailTxt.sendKeys("emailvalidationone@gmail.com");
	emailTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	Thread.sleep(2000);
	
	vendorAccountTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(3000);
	
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	departmentTxt.click();
	departmentTxt.sendKeys(Keys.SPACE);

	int departmentcount=departmentListCount.size();

	System.err.println(departmentcount);

	for(int i=0 ; i < departmentcount ;i++)
	{
		String data=departmentListCount.get(i).getText();

		if(data.equalsIgnoreCase("DUBAI"))
		{
			departmentListCount.get(i).click();

			break;
		}
	}

	departmentTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATPlaceOFSupply));
	salesInvoiceVATPlaceOFSupply.click();
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.END);
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SHIFT,Keys.HOME);
	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.SPACE);


	int placeOFSupplyListCount=placeOFSupplyList.size();

	System.err.println("placeOFSupplyListCount   : "+placeOFSupplyListCount);

	for(int i=0 ; i < placeOFSupplyListCount ;i++)
	{
		String data=placeOFSupplyList.get(i).getText();

		if(data.equalsIgnoreCase("Abu Dhabi"))
		{
			placeOFSupplyList.get(i).click();

			break;
		}
	}

	salesInvoiceVATPlaceOFSupply.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
	jurisdictionTxt.click();
	jurisdictionTxt.sendKeys(Keys.END);
	jurisdictionTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	jurisdictionTxt.sendKeys("DUBAI");
	Thread.sleep(2000);
	jurisdictionTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvWareHouseTxt));
	pvWareHouseTxt.click();
	pvWareHouseTxt.sendKeys(Keys.SPACE);

	int warehousecount=pvwareHouseListCount.size();

	System.err.println(warehousecount);

	for(int i=0 ; i < warehousecount ;i++)
	{
		String data=pvwareHouseListCount.get(i).getText();

		if(data.equalsIgnoreCase("SECUNDERABAD"))
		{
			pvwareHouseListCount.get(i).click();

			break;
		}	
	}

	pvWareHouseTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.sendKeys(Keys.END);
	enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	enter_ItemTxt.sendKeys(Keys.BACK_SPACE);
	enter_ItemTxt.sendKeys(Keys.SPACE);

	int itemcount1=itemListCount.size();

	System.err.println(itemcount1);

	for(int i=0 ; i < itemcount1 ;i++)
	{
		String data=itemListCount.get(i).getText();

		if(data.equalsIgnoreCase("STD RATE COGS ITEM"))
		{
			itemListCount.get(i).click();

			break;
		}
	}

	enter_ItemTxt.sendKeys(Keys.TAB);

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	select1stRow_5thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	select1stRow_8thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
	enter_AQTxt.sendKeys("6");
	enter_AQTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
	enter_FQTxt.sendKeys("6");
	enter_FQTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	select1stRow_11thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	select1stRow_14thColumn.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.click();
	enter_Rate.clear();
	enter_Rate.sendKeys("10");
	enter_Rate.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
	enter_Gross.click();
	enter_Gross.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_16thColumn));
	select1stRow_16thColumn.click();	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SIVAT_enter_Discount));
	SIVAT_enter_Discount.click();
	SIVAT_enter_Discount.clear();
	SIVAT_enter_Discount.sendKeys("3");
	SIVAT_enter_Discount.sendKeys(Keys.TAB);
	Thread.sleep(2000);

	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchPickOnFIFOIcon));
	batchPickOnFIFOIcon.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchOkIcon));
	batchOkIcon.click();
*/
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String docno=documentNumberTxt.getAttribute("value");
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
	voucherSaveBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
	billRefNewReferenceTxt.click();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
	pickBtn.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
	Bill_OkBtn.click();

	Thread.sleep(3000);
	
	checkValidationMessage("");
	Thread.sleep(2000);
	
	
	String expMessage1 = "Voucher saved successfully";

	String actMessage = checkValidationMessage(expMessage1);
	String expMessage2 = ": "+docno;

			
	System.out.println("Actual Message    : "+actMessage);
	System.out.println("Expected Message  : "+expMessage1);

	if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) )

	
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		Thread.sleep(2000);
		return true;
	}	
	else
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		Thread.sleep(2000);
		return false;
	}



}


public static boolean checkCreatingSMSTemplateforTriggersinSalesInvoiceVATVoucher() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
{
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	System.out.println("Setting buton is enabled");
	serachMenuTextHomePage.click();
	serachMenuTextHomePage.sendKeys("Configure Transactions");
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	searchMenuTextClick.click();	  
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATBtn));
	salesInvoiceVATBtn.click();
					 	    	
	Thread.sleep(2000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
	triggersBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	editLayoutTriggerName.click();
	editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
	Thread.sleep(1000);
	editLayoutTriggerName.sendKeys("Sales Map Trigger");
	Thread.sleep(2000);
	editLayoutTriggerName.sendKeys(Keys.TAB);
	Thread.sleep(5000);
/*	
	if(SendAlertChkBox.isDisplayed()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
	triggerTabAlertBtn.click();
	}
	Thread.sleep(2000);*/
	
	
	if(sendSMSChkBox.isSelected()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendSMSChkBoxSelected));
	sendSMSChkBoxSelected.click();
	}
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createNewSMSBtn));
	createNewSMSBtn.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newButtonInCreateTemplate));
	newButtonInCreateTemplate.click();
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplates));
	emailTemplates.sendKeys("SalesMapSMSTriggerTemplate");
	emailTemplates.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
	companyNameTemplateArea.click();
	Thread.sleep(1500);
	

	Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\salesTriggerAlert.exe");
	
	Thread.sleep(150000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
	saveBtnInEmailTemplateHeader.click();
	
	Thread.sleep(2000);
	System.err.println("Template save button clicked");
	
	String expEmailTemplateSaveMsg = "Template saved successfully";
	String actEmailTemplateSaveMsg =  checkValidationMessage(expEmailTemplateSaveMsg);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
	closeButtonInCreateTemplate.click();
	Thread.sleep(2000);
	if(actEmailTemplateSaveMsg.equalsIgnoreCase(expEmailTemplateSaveMsg))
	{
		return true;
	}
	else
	{
		return false;
	}
	
	
}


public boolean checkCreatedSMSAlertinSalesInvoiceVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(2000);
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	System.out.println("Setting buton is enabled");
	serachMenuTextHomePage.click();
	serachMenuTextHomePage.sendKeys("Configure Transactions");
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	searchMenuTextClick.click();	  
	
	Thread.sleep(2000);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceVATBtn));
	salesInvoiceVATBtn.click();
					 	    	
	Thread.sleep(2000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
	triggersBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	editLayoutTriggerName.click();
	editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
	Thread.sleep(1000);
	editLayoutTriggerName.sendKeys("Sales Map Trigger");
	Thread.sleep(2000);
	editLayoutTriggerName.sendKeys(Keys.TAB);
	Thread.sleep(5000);
	
	
	
	/*if(SendAlertChkBox.isDisplayed()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
	triggerTabAlertBtn.click();
	}
	Thread.sleep(2000);*/
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SendAlertChkBoxSelected));
	if(sendAlertChkBox.isSelected()==true)
	{
		SendAlertChkBoxSelected.click();
	}
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
	if(sendEmailChkBox.isSelected()==true)
	{
	sendEmailChkBoxSelected.click();
	}
	Thread.sleep(2000);
	
	if(sendSMSChkBox.isSelected()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendSMSChkBoxSelected));
	sendSMSChkBoxSelected.click();
	}
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectSMSTemplateAlert));
	Select s=new Select(selectSMSTemplateAlert);
	s.selectByVisibleText("SalesMapSMSTriggerTemplate");
	Thread.sleep(2000);
	String actMsg=s.getFirstSelectedOption().getText();
	String expMsg="SalesMapSMSTriggerTemplate";
	
	System.out.println("Actual Message: 	" 		+ actMsg);
	System.out.println("Expected Message:  	"	+expMsg);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
	editLayoutTriggersSaveBtn.click();
	Thread.sleep(2500);
	
	String expSaveMessage = "Data saved successfully";
	String actSaveMessage = checkValidationMessage(expSaveMessage);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
	updateBtn.click();
	
	
	String expUpdateMessage = "Data saved successfully";
	String actUpdateMessage = checkValidationMessage(expUpdateMessage);
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
	closeBtn.click();
	
	
	
	
	if(actMsg.equalsIgnoreCase(expMsg) && expUpdateMessage.equalsIgnoreCase(expUpdateMessage))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}



public static boolean CheckLogin() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
        Thread.sleep(3000);

        getDriver().navigate().refresh();

        LoginPage lp=new LoginPage(getDriver()); 

        lp.checkLoginPageTitleByURLInputInBrowser();

        String unamelt="su";

        String pawslt="su";

        lp.enterUserName(unamelt);

        lp.enterPassword(pawslt);

        Thread.sleep(2000);

        String compname = "Automation Company : 08/10/2020";

        Select oSelect = new Select(companyDropDownList);

        List<WebElement> elementCount = oSelect.getOptions();

        int cqSize = elementCount.size();

        System.out.println("CompanyDropdownList Count :" + cqSize);

        int i;

        for (i = 0; i < elementCount.size(); i++) {

                elementCount.get(i).getText();

                String optionName = elementCount.get(i).getText();
                if (optionName.toUpperCase().startsWith(compname.toUpperCase())) {
                        System.out.println("q" + elementCount.get(i).getText());
                        elementCount.get(i).click();
                }

        }

        Thread.sleep(2000);

        lp.clickOnSignInBtn();

        Thread.sleep(2000);

        String actUserInfo1=userNameDisplay.getText();

        System.out.println("User Info  : "+actUserInfo1);

        System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

        //getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
        //companyLogo.click();

        String getCompanyTxt1=Company_Name.getText();
        String getLoginCompanyName1=getCompanyTxt1.substring(0, 31);
        System.out.println("company name  :  "+ getLoginCompanyName1);
       // companyLogo.click();

        String expUserInfo1           ="SU";
        String expLoginCompanyName1   ="Automation Company : 08/10/2020";

        System.out.println("UserInfo1             : "+actUserInfo1            +" Value Expected : "+expUserInfo1);
        System.out.println("LoginCompanyName1     : "+getLoginCompanyName1    +" Value Expected : "+expLoginCompanyName1);

        if(actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1))
        {
                return true;
        }
        else
        {
                return false;
        }

}


public boolean checkLogoutPurchaseandSalesInvoiceVoucherVATPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
{
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	 
	 try
		{
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		  userNameDisplayLogo.click();
		  Thread.sleep(2000);
		 
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		  logoutOption.click();
		  
		  Thread.sleep(2000);
		  
		  boolean actUserLoginPage              = username.isDisplayed() && username.isEnabled()
                                               && password.isDisplayed() && password.isEnabled();
                                      
		  boolean expUserLoginPage              = true;
		  
		  if(actUserLoginPage==expUserLoginPage)  
	      {
			System.out.println("***Test Pass: Login Successfull***");
			
			return true;
		  }
	      else
	      {
	  	 
			System.out.println("***Test Fail: Login Not Successfull***");
			
			return false;
		  }
		}
		catch (Exception e)
		{
		 	String exception = e.getMessage();
		 		
			return false;
		}
	}

@FindBy(xpath = "(//*[@id='5058'])[1]")
private static WebElement triggerWarehouseName;


public boolean checkSavingTriggeronWarehouseRaiseSameVoucher()throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys("Warehouse Trigger");
		Thread.sleep(4000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(8000);
		
		triggersConjuctionDrpdwn.click();
        Thread.sleep(2000);		
		triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);			
		triggersConjuctionDrpdwn.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
		triggersSelectFieldTxt.click();
					
		Thread.sleep(3000);
		
		getAction().moveToElement(triggerWarehouseExpansionBtn).build().perform();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerWarehouseExpansionBtn));
		triggerWarehouseExpansionBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerWarehouseName));
		triggerWarehouseName.click();
		
		Thread.sleep(2000);
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
		triggersSelectOperdrpdwn.click();
		triggersSelectOperdrpdwn.sendKeys("Equal");
		
		
		Thread.sleep(5000);
		triggersSelectOperdrpdwn.sendKeys(Keys.TAB);
		
		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
		triggersCompareWithdrpdwn.click();
		triggersCompareWithdrpdwn.sendKeys("value");
		//triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
		Thread.sleep(2000);
		triggersCompareWithdrpdwn.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		
		triggersValueTxt.sendKeys("Hyderabad");
		Thread.sleep(1500);
		triggersValueTxt.sendKeys(Keys.TAB);

		
		Thread.sleep(3000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
		if(editLayoutTriggerRaiseDoc.isSelected()==false)
		{
			editLayoutTriggerRaiseDocSelected.click();
		}
		
		Thread.sleep(1000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
		triggerTabAlertBtn.click();
		
		Thread.sleep(2000);*/
		
		ScrollIntoView(editLayoutTriggersAllVouchers);
		Thread.sleep(1000);
							
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
		Select s=new Select(editLayoutTriggersAllVouchers);
		s.selectByVisibleText("Job orders");
		Thread.sleep(3000);
		
		
		/*getAction().moveToElement(trigger1stRow1stCol).build().perform();
		
		
		ClickUsingJs(trigger1stRow1stCol);
		Thread.sleep(2000);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectFieldDropdown));
	
		Thread.sleep(1000);
		Select s1=new Select(triggerSelectFieldDropdown);
		s1.selectByVisibleText("Discount");
		Thread.sleep(3000);
		triggerSelectFieldDropdown.sendKeys(Keys.TAB);
		
		
		ClickUsingJs(trigger1stRow2ndCol);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectMapDropdown));
		Select s2=new Select(triggerSelectMapDropdown);
		s2.selectByVisibleText("Discount");
		Thread.sleep(2000);*/
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		Thread.sleep(2500);
		
		String expSaveMessage = "Data saved successfully";
		String actSaveMessage = checkValidationMessage(expSaveMessage);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		
		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Settings_closeBtn));
		Settings_closeBtn.click();
			
		if( actSaveMessage.equalsIgnoreCase(expSaveMessage))
		{
			System.out.println("Test Pass : Trigger  Added Successfully");
			
			return true;
		}
		else
		{
			System.out.println("Test Fail : Trigger NOT Added Successfully");
			
			return false;
		}
}


		public boolean checkSavingPurchaseVoucherVATforWarehouseTrigger() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException 
		{
			Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();
		
		vendorAccountTxt.sendKeys(Keys.SPACE);	
		
		ArrayList<String>  actVendorAccountList= new ArrayList<String>(); 
		
		int count=vendorAccountListCount.size();
		
		System.err.println(count);
		
		for(int i=0 ; i < count ;i++)
		{
			String data=vendorAccountListCount.get(i).getText();
		
			if (data.equalsIgnoreCase("Vendor A")) 
			{
				vendorAccountListCount.get(i).click();
				break;
			}
		}
		
		vendorAccountTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(Keys.SPACE);
		
		int currencycount=currencyListCount.size();
		
		System.err.println(currencycount);
		
		for(int i=0 ; i < currencycount ;i++)
		{
			String data=currencyListCount.get(i).getText();
		
			if (data.equalsIgnoreCase("INR")) 
			{
				currencyListCount.get(i).click();
				break;
			}
		}
		
		voucherHeaderCurrency.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);	
		
		ArrayList<String> actDepartmentList= new ArrayList<String>(); 
		
		int departmentCount=departmentListCount.size();
		
		System.err.println(departmentCount);
		
		for(int i=0 ; i < departmentCount ;i++)
		{
			String data=departmentListCount.get(i).getText();
		
			if (data.equalsIgnoreCase("INDIA")) 
			{
				departmentListCount.get(i).click();
				break;
			}
		}
		
		departmentTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.clear();
		placeOFSupplyTxt.sendKeys("Abu Dhabi");
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys("DUBAI");
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		
		enter_WarehouseTxt.sendKeys(Keys.SPACE);
		
		int warehousecount=warehouseBodyComboList.size();
		
		for(int i=0 ; i < warehousecount ;i++)
		{
			String data=warehouseBodyComboList.get(i).getText();
		
			if (data.equalsIgnoreCase("HYDERABAD")) 
			{
				warehouseBodyComboList.get(i).click();
				break;
			}
		}
		
		enter_WarehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount=pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount; i++) 
		{
			String Item=pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("BR COGS ITEM")) 
			{
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("2");
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));
		select1stRow_13thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
		enter_Discount.click();
		enter_Discount.clear();
		enter_Discount.sendKeys("10");
		enter_Discount.sendKeys(Keys.TAB);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvVat));
		enter_PvVat.sendKeys(Keys.TAB);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
		enter_PvTaxable.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys("PVATR1#V1");
		enter_Batch.sendKeys(Keys.TAB);
		
		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		
		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 2); 
		
		String next2DaysDate=df.format(cal.getTime());
		
		System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "+next2DaysDate);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next2DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		Thread.sleep(2000);
		
				
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": SU/IND/TEXT2";
		
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);
		
		if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) )
		{
			return true;
		} 
		else 
		{
			return false;
		}
	}

		
		public boolean checkRaisedVoucherAfterSavingTriggerVoucher2()throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			String actWarehouseTxt = checkWarehouseinPurchaseVoucherVAT();
			

			Thread.sleep(10000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Job Order");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(8000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

			int voucherGridDocNoCount = voucherGridDocNo.size();

			for (int i = 0; i < voucherGridDocNoCount; i++) {
				String data = "2";
				if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
					voucherGridIndexChkBox.get(i).click();
					break;
				}
			}

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
			editBtn.click();
			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
			String actWarehouse = actWarehouseTxt;
			String expWarehouse = warehouseTxt.getAttribute("value");
			

			System.out.println("Warehouse Text:" + actWarehouse + "Expected :" + expWarehouse);

			if (actWarehouse.equalsIgnoreCase(expWarehouse)) {
				return true;
			} else {
				return false;
			}

		}

	
		
		public static String checkWarehouseinPurchaseVoucherVAT()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();
			Thread.sleep(2000);

			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			documentNumberTxt.click();
			documentNumberTxt.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
			documentNumberTxt.sendKeys("SU/IND/TEXT2");
			Thread.sleep(4000);
			documentNumberTxt.sendKeys(Keys.TAB);
			Thread.sleep(8000);

			checkValidationMessage("Voucher loaded successfully");
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
						
			String warehouseTxt = select1stRow_1stColumn.getText();
			

			return warehouseTxt;

		}
		
		
		@FindBy(xpath = "//*[@class='icon-settings hiconright2']")
		public static WebElement PVVAT_SettingsBtn;

		
		@FindBy(xpath = "(//*[@id='268435459']/span)[1]")
		private static WebElement triggerDepartmentExpansionBtn;
		
		
		@FindBy(xpath = "(//*[@id='5042'])[1]")
		private static WebElement triggerDepartmentName;
		
		
		@FindBy(xpath = "(//*[@class='icon-pluse icon-font6'])[2]")
		public static WebElement plusSign;
		
		
		@FindBy(xpath = "//*[@id='17_1_AdvanceFilter_']/table/tbody/tr[2]/td[1]/select")
		private static WebElement triggersConjuctionDrpdwn2ndRow;
		
		@FindBy(xpath = "//div/table/tbody/tr[2]/td[2]/input")
		private static WebElement triggersSelectFieldTxt2ndRow;
		
		@FindBy(xpath = "//tbody/tr[2]/td[3]/select")
		private static WebElement triggersSelectOperdrpdwn2ndRow;
		
		@FindBy(xpath = "//tbody/tr[2]/td[4]/select")
		private static WebElement triggersCompareWithdrpdwn2ndRow;

		@FindBy(xpath = "(//*[@id='23']/span)[2]")
		private static WebElement triggerItemExpansionBtn;
		
		@FindBy(xpath = "(//*[@id='5021'])[1]")
		private static WebElement triggerItemName;
		
		@FindBy(xpath = "//*[@id='17_1_AdvanceFilter_']/table/tbody/tr[2]/td[5]")
		private static WebElement triggersValueToEnterTxt2ndRow;
		
		public boolean checkSavingMultipleConditionTriggerinPurchaseVoucherVAT() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
			purchaseVouchersVat.click();

			Thread.sleep(6000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PVVAT_SettingsBtn));
			PVVAT_SettingsBtn.click();
			Thread.sleep(6000);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();

			Thread.sleep(6000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
			Thread.sleep(1000);
			editLayoutTriggerName.sendKeys("Multiple Condition Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(8000);

			triggersConjuctionDrpdwn.click();
			Thread.sleep(2000);
			triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);
			triggersConjuctionDrpdwn.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
			triggersSelectFieldTxt.click();

			Thread.sleep(3000);

			
			  getAction().moveToElement(triggerDepartmentExpansionBtn).build().perform();
			  Thread.sleep(1000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerDepartmentExpansionBtn));
			triggerDepartmentExpansionBtn.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerDepartmentName));
			triggerDepartmentName.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
			triggersSelectOperdrpdwn.click();
			triggersSelectOperdrpdwn.sendKeys("Equal");

			Thread.sleep(5000);
			triggersSelectOperdrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(5000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
			triggersCompareWithdrpdwn.click();
			triggersCompareWithdrpdwn.sendKeys("value");
			// triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
			Thread.sleep(2000);
			triggersCompareWithdrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(2000);

			/*triggersValueToEnterTxt.click();
			Thread.sleep(1000);*/
			triggersValueTxt.sendKeys(Keys.SPACE);
			Thread.sleep(3000);
			triggersValueTxt.sendKeys("INDIA");
			Thread.sleep(1500);
			// triggersValueTxt.sendKeys(Keys.TAB);

			Thread.sleep(1500);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(plusSign));
			plusSign.click();
			Thread.sleep(2000);

			triggersConjuctionDrpdwn2ndRow.click();
			Thread.sleep(2000);
			// triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);
			triggersConjuctionDrpdwn2ndRow.sendKeys(Keys.TAB);
			triggersConjuctionDrpdwn2ndRow.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt2ndRow));
			triggersSelectFieldTxt2ndRow.click();

			Thread.sleep(3000);

			//getAction().moveToElement(triggerItemExpansionBtn).build().perform();
		//	Thread.sleep(3000);
			
			ScrollToElement(triggerItemExpansionBtn);
			Thread.sleep(3000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerItemExpansionBtn));
			triggerItemExpansionBtn.click();

			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerItemName));
			triggerItemName.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn2ndRow));
			triggersSelectOperdrpdwn2ndRow.click();
			triggersSelectOperdrpdwn2ndRow.sendKeys("Equal");

			Thread.sleep(5000);
			triggersSelectOperdrpdwn2ndRow.sendKeys(Keys.TAB);

			Thread.sleep(5000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn2ndRow));
			triggersCompareWithdrpdwn2ndRow.click();
			triggersCompareWithdrpdwn2ndRow.sendKeys("value");
			// triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
			Thread.sleep(2000);
			triggersCompareWithdrpdwn2ndRow.sendKeys(Keys.TAB);

			Thread.sleep(2000);

			triggersValueToEnterTxt2ndRow.click();
			Thread.sleep(1000);
			triggersValueTxt.sendKeys(Keys.SPACE);
			Thread.sleep(3000);
			triggersValueTxt.sendKeys("STD RATE COGS ITEM");
			Thread.sleep(1500);
			triggersValueTxt.sendKeys(Keys.TAB);

			Thread.sleep(1500);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
			if(editLayoutTriggerRaiseDoc.isSelected()==false)
			{
				editLayoutTriggerRaiseDocSelected.click();
			}
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
			Select s = new Select(editLayoutTriggersAllVouchers);
			s.selectByVisibleText("Purchases Vouchers");
			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
			editLayoutTriggersSaveBtn.click();
			Thread.sleep(2500);

			String expSaveMessage = "Data saved successfully";
			String actSaveMessage = checkValidationMessage(expSaveMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();

			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();

			if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
				System.out.println("Test Pass : Trigger  Added Successfully");

				return true;
			} else {
				System.out.println("Test Fail : Trigger NOT Added Successfully");

				return false;
			}

		}	
		
		
		public boolean checkSavingPurchaseVoucherVATVoucherforMultipleConditionTrigger() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException 
		{
			

			Thread.sleep(2000);
			
			  getDriver().navigate().refresh(); Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu)); 
			  financialsMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu)); 
			  financialsTransactionMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			  financialsTransactionsPurchaseMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat)); 
			  purchaseVouchersVat.click();
			  
			  Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn)); 
			  newBtn.click();
			  
			  checkValidationMessage("Screen opened");
			 

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.click();

			vendorAccountTxt.sendKeys(Keys.SPACE);

			ArrayList<String> actVendorAccountList = new ArrayList<String>();

			int count = vendorAccountListCount.size();

			System.err.println(count);

			for (int i = 0; i < count; i++) {
				String data = vendorAccountListCount.get(i).getText();

				if (data.equalsIgnoreCase("Vendor B")) {
					vendorAccountListCount.get(i).click();
					break;
				}
			}

			vendorAccountTxt.sendKeys(Keys.TAB);

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
			voucherHeaderDueDate.click();
			voucherHeaderDueDate.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			voucherHeaderCurrency.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
			voucherHeaderCurrency.sendKeys(Keys.SPACE);

			int currencycount = currencyListCount.size();

			System.err.println(currencycount);

			for (int i = 0; i < currencycount; i++) {
				String data = currencyListCount.get(i).getText();

				if (data.equalsIgnoreCase("INR")) {
					currencyListCount.get(i).click();
					break;
				}
			}

			voucherHeaderCurrency.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);

			ArrayList<String> actDepartmentList = new ArrayList<String>();

			int departmentCount = departmentListCount.size();

			System.err.println(departmentCount);

			for (int i = 0; i < departmentCount; i++) {
				String data = departmentListCount.get(i).getText();

				if (data.equalsIgnoreCase("INDIA")) {
					departmentListCount.get(i).click();
					break;
				}
			}

			departmentTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
			placeOFSupplyTxt.click();
			placeOFSupplyTxt.clear();
			placeOFSupplyTxt.sendKeys("Abu Dhabi");
			Thread.sleep(2000);
			placeOFSupplyTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
			jurisdictionTxt.click();
			jurisdictionTxt.sendKeys("DUBAI");
			Thread.sleep(2000);
			jurisdictionTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();

			enter_WarehouseTxt.sendKeys(Keys.SPACE);

			int warehousecount = warehouseBodyComboList.size();

			for (int i = 0; i < warehousecount; i++) {
				String data = warehouseBodyComboList.get(i).getText();

				if (data.equalsIgnoreCase("Mumbai")) {
					warehouseBodyComboList.get(i).click();
					break;
				}
			}

			enter_WarehouseTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
			int pvvGridItemListCount = pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) {
				String Item = pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
			enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.click();
			enter_Quantity.clear();
			enter_Quantity.sendKeys("2");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.click();
			enter_Rate.clear();
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.click();
			enter_Gross.sendKeys(Keys.TAB);

			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
			voucherSaveBtn.click();

			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
			billRefNewReferenceTxt.click();
			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
			pickBtn.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
			Bill_OkBtn.click();

			Thread.sleep(2000);

			String expMessage1 = "Voucher saved successfully";

			String actMessage = checkValidationMessage(expMessage1);
			String expMessage2 = ": SU/IND/TEXT3";

			System.out.println("Actual Message    : " + actMessage);
			System.out.println("Expected Message  : " + expMessage1);

			if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)) {
				return true;
			} else {
				return false;
			}
		}
		
		
		@FindBy(xpath = "//*[@id='2008']")
		public static WebElement purchaseVouchers;
		
		public boolean checkRaisedVoucherforMultipleConditionTrigger() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			  getDriver().navigate().refresh(); Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu)); 
			  financialsMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu)); 
			  financialsTransactionMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			  financialsTransactionsPurchaseMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchers)); 
			  purchaseVouchers.click();
			  
			  Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn)); 
			  

				int voucherGridDocNoCount = voucherGridDocNo.size();

				for (int i = 0; i < voucherGridDocNoCount; i++) {
					String data = "1";
					if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
						voucherGridIndexChkBox.get(i).click();
						break;
					}
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
				editBtn.click();
				Thread.sleep(4000);
			  checkValidationMessage("Screen opened");
			  
			  String actDeptTxt=departmentTxt.getAttribute("value");
			  String expDeptTxt="INDIA";
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			  String actItemTxt=select1stRow_1stColumn.getText();
			  String expItemTxt="STD RATE COGS";
			  
			  System.out.println("Department :"	+	actDeptTxt	+	"Expected :"	+expDeptTxt);
			  System.out.println("Item :"	+	actItemTxt	+	"Expected :"	+expItemTxt);
			  
			  if(actDeptTxt.equalsIgnoreCase(expDeptTxt) && actItemTxt.equalsIgnoreCase(actItemTxt))
			  {
				  return true;
			  }
			  else
			  {
				  return false;
			  }

		}
		
		
		
		public boolean checkSavingPurchaseVoucherVATVoucher2forMultipleConditionTrigger() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException 
		{
			

			Thread.sleep(2000);
			
			  getDriver().navigate().refresh(); Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu)); 
			  financialsMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu)); 
			  financialsTransactionMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			  financialsTransactionsPurchaseMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat)); 
			  purchaseVouchersVat.click();
			  
			  Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn)); 
			  newBtn.click();
			  
			  checkValidationMessage("Screen opened");
			 

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.click();

			vendorAccountTxt.sendKeys(Keys.SPACE);

			ArrayList<String> actVendorAccountList = new ArrayList<String>();

			int count = vendorAccountListCount.size();

			System.err.println(count);

			for (int i = 0; i < count; i++) {
				String data = vendorAccountListCount.get(i).getText();

				if (data.equalsIgnoreCase("Vendor B")) {
					vendorAccountListCount.get(i).click();
					break;
				}
			}

			vendorAccountTxt.sendKeys(Keys.TAB);

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
			voucherHeaderDueDate.click();
			voucherHeaderDueDate.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			voucherHeaderCurrency.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
			voucherHeaderCurrency.sendKeys(Keys.SPACE);

			int currencycount = currencyListCount.size();

			System.err.println(currencycount);

			for (int i = 0; i < currencycount; i++) {
				String data = currencyListCount.get(i).getText();

				if (data.equalsIgnoreCase("INR")) {
					currencyListCount.get(i).click();
					break;
				}
			}

			voucherHeaderCurrency.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			departmentTxt.sendKeys(Keys.SPACE);

			ArrayList<String> actDepartmentList = new ArrayList<String>();

			int departmentCount = departmentListCount.size();

			System.err.println(departmentCount);

			for (int i = 0; i < departmentCount; i++) {
				String data = departmentListCount.get(i).getText();

				if (data.equalsIgnoreCase("INDIA")) {
					departmentListCount.get(i).click();
					break;
				}
			}

			departmentTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
			placeOFSupplyTxt.click();
			placeOFSupplyTxt.clear();
			placeOFSupplyTxt.sendKeys("Abu Dhabi");
			Thread.sleep(2000);
			placeOFSupplyTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
			jurisdictionTxt.click();
			jurisdictionTxt.sendKeys("DUBAI");
			Thread.sleep(2000);
			jurisdictionTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
			enter_WarehouseTxt.click();

			enter_WarehouseTxt.sendKeys(Keys.SPACE);

			int warehousecount = warehouseBodyComboList.size();

			for (int i = 0; i < warehousecount; i++) {
				String data = warehouseBodyComboList.get(i).getText();

				if (data.equalsIgnoreCase("Mumbai")) {
					warehouseBodyComboList.get(i).click();
					break;
				}
			}

			enter_WarehouseTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);
			int pvvGridItemListCount = pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) {
				String Item = pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase("BR COGS ITEM")) {
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
			enter_TaxCode.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
			enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
			select1stRow_9thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.click();
			enter_Quantity.clear();
			enter_Quantity.sendKeys("2");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			select1stRow_11thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.click();
			enter_Rate.clear();
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
			enter_Gross.click();
			enter_Gross.sendKeys(Keys.TAB);

			
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn)); 
			  select1stRow_14thColumn.click();
			  
			
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable)); 
			  enter_PvTaxable.sendKeys(Keys.TAB);
			  
			  Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch)); 
			  enter_Batch.click(); enter_Batch.sendKeys("PVATR1#V5");
			  enter_Batch.sendKeys(Keys.TAB);
			  
			  DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			  
			  Calendar cal=Calendar.getInstance(); SimpleDateFormat currentDate = new
			  SimpleDateFormat("dd/MM/yyyy"); cal.add(Calendar.DATE, 2);
			  
			  String next2DaysDate=df.format(cal.getTime());
			  
			  System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "
			  +next2DaysDate);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			  enter_Expirydate)); enter_Expirydate.click();
			  enter_Expirydate.sendKeys(Keys.HOME);
			  enter_Expirydate.sendKeys(next2DaysDate);
			  enter_Expirydate.sendKeys(Keys.TAB);
			 

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
			voucherSaveBtn.click();

			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
			billRefNewReferenceTxt.click();
			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
			pickBtn.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
			Bill_OkBtn.click();

			Thread.sleep(2000);

			String expMessage1 = "Voucher saved successfully";

			String actMessage = checkValidationMessage(expMessage1);
			String expMessage2 = ": SU/IND/TEXT4";

			System.out.println("Actual Message    : " + actMessage);
			System.out.println("Expected Message  : " + expMessage1);

			if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)) {
				return true;
			} else {
				return false;
			}
		}
		
		
		
		public boolean checkRaisedVoucherforMultipleConditionTriggerisNotTriggered() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			  getDriver().navigate().refresh(); Thread.sleep(2000);
			  Thread.sleep(2000);
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu)); 
			  financialsMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu)); 
			  financialsTransactionMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			  financialsTransactionsPurchaseMenu.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchers)); 
			  purchaseVouchers.click();
			  
			  Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn)); 
			  

				int voucherGridDocNoCount = voucherGridDocNo.size();
				boolean flag=true;

				for (int i = 0; i < voucherGridDocNoCount; i++) {
					String data = "2";
					if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
						System.out.println("Voucher Trigged");
						flag=false;
						break;
						
					}
				}
					if(flag)
					{
						return true;
					}
					else
					{
						return false;
					}
				}

		
		@FindBy(xpath="(//*[@id='4'])[1]")
		public static WebElement miscellaneousTab;
		
		public boolean checkEnablingDoNotModifyDocumentIfTriggerisRaised() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Preferences");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTab));
			miscellaneousTab.click();
			Thread.sleep(1500);
			
			getAction().moveToElement(donotModifyDocChkBoxSelected).build().perform();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(donotModifyDocChkBoxSelected));
			if(donotModifyDocChkBox.isSelected()==false)
			{
				donotModifyDocChkBoxSelected.click();
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			getWaitForAlert();
			getAlert().accept();

			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();

			if(actUpdateMessage.equalsIgnoreCase(expUpdateMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
		}

	public boolean checkSavingPurchaseVoucherVATVoucherandEditingVoucherafterVoucherTriggered() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

		

		Thread.sleep(2000);
		
		  getDriver().navigate().refresh();
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu)); 
		  financialsMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu)); 
		  financialsTransactionMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		  financialsTransactionsPurchaseMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat)); 
		  purchaseVouchersVat.click();
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn)); 
		  newBtn.click();
		  
		  checkValidationMessage("Screen opened");
		 

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();

		vendorAccountTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actVendorAccountList = new ArrayList<String>();

		int count = vendorAccountListCount.size();

		System.err.println(count);

		for (int i = 0; i < count; i++) {
			String data = vendorAccountListCount.get(i).getText();

			if (data.equalsIgnoreCase("Vendor B")) {
				vendorAccountListCount.get(i).click();
				break;
			}
		}

		vendorAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		voucherHeaderCurrency.sendKeys(Keys.SPACE);

		int currencycount = currencyListCount.size();

		System.err.println(currencycount);

		for (int i = 0; i < currencycount; i++) {
			String data = currencyListCount.get(i).getText();

			if (data.equalsIgnoreCase("INR")) {
				currencyListCount.get(i).click();
				break;
			}
		}

		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actDepartmentList = new ArrayList<String>();

		int departmentCount = departmentListCount.size();

		System.err.println(departmentCount);

		for (int i = 0; i < departmentCount; i++) {
			String data = departmentListCount.get(i).getText();

			if (data.equalsIgnoreCase("INDIA")) {
				departmentListCount.get(i).click();
				break;
			}
		}

		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.clear();
		placeOFSupplyTxt.sendKeys("Abu Dhabi");
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys("DUBAI");
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();

		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount = warehouseBodyComboList.size();

		for (int i = 0; i < warehousecount; i++) {
			String data = warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("Mumbai")) {
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount = pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount; i++) {
			String Item = pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("BR COGS ITEM")) {
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("2");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn)); 
		  select1stRow_14thColumn.click();
		  
		
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable)); 
		  enter_PvTaxable.sendKeys(Keys.TAB);
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch)); 
		  enter_Batch.click(); enter_Batch.sendKeys("PVATR1#V5");
		  enter_Batch.sendKeys(Keys.TAB);
		  
		  DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		  
		  Calendar cal=Calendar.getInstance(); SimpleDateFormat currentDate = new
		  SimpleDateFormat("dd/MM/yyyy"); cal.add(Calendar.DATE, 2);
		  
		  String next2DaysDate=df.format(cal.getTime());
		  
		  System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "
		  +next2DaysDate);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		  enter_Expirydate)); enter_Expirydate.click();
		  enter_Expirydate.sendKeys(Keys.HOME);
		  enter_Expirydate.sendKeys(next2DaysDate);
		  enter_Expirydate.sendKeys(Keys.TAB);
		 

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		Thread.sleep(2000);

		String expMessage1 = "Voucher saved successfully";

		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": SU/IND/TEXT5";

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage1);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("SU/IND/TEXT5");
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		
		Thread.sleep(2000);
		checkValidationMessage("Voucher cannot be modified because trigger is raised from it");
		checkValidationMessage("Voucher loaded successfully");
		
		

		if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)) {
			return true;
		} else {
			return false;
		}

	}
	@FindBy(xpath="//input[@id='chkDonotModifyDoc']")
	public static WebElement donotModifyDocChkBox;

	@FindBy(xpath="//input[@id='chkDonotModifyDoc']/../span")
	public static WebElement donotModifyDocChkBoxSelected;
	
	public boolean checkDisablingDoNotModifyDocumentIfTriggerisRaised() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Preferences");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTab));
		miscellaneousTab.click();
		Thread.sleep(1500);
		
		getAction().moveToElement(donotModifyDocChkBoxSelected).build().perform();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(donotModifyDocChkBoxSelected));
		if(donotModifyDocChkBox.isSelected()==true)
		{
			donotModifyDocChkBoxSelected.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		getWaitForAlert();
		getAlert().accept();

		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();

		if(actUpdateMessage.equalsIgnoreCase(expUpdateMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
		
	
	public boolean checkSavingTriggerWithoutEnablingLineWiseCheckBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		    Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Configure Transactions");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(2000);

		

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
			purchaseVouchersVATBtn.click();

			Thread.sleep(6000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();

			Thread.sleep(10000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
			Thread.sleep(1000);
			editLayoutTriggerName.sendKeys("LineWise Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(8000);

			triggersConjuctionDrpdwn.click();
			Thread.sleep(2000);
			triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);
			triggersConjuctionDrpdwn.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
			triggersSelectFieldTxt.click();

			Thread.sleep(3000);

			getAction().moveToElement(triggerItemExpansionBtn).build().perform();
			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerItemExpansionBtn));
			triggerItemExpansionBtn.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerItemName));
			triggerItemName.click();

			Thread.sleep(2000);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
			triggersSelectOperdrpdwn.click();
			triggersSelectOperdrpdwn.sendKeys("Equal");

			Thread.sleep(5000);
			triggersSelectOperdrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(5000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
			triggersCompareWithdrpdwn.click();
			triggersCompareWithdrpdwn.sendKeys("value");
			// triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
			Thread.sleep(2000);
			triggersCompareWithdrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(2000);

			/*triggersValueToEnterTxt.click();
			Thread.sleep(1000);*/
			triggersValueTxt.sendKeys(Keys.SPACE);
			Thread.sleep(3000);
			triggersValueTxt.sendKeys("STD RATE COGS ITEM");
			Thread.sleep(1500);
			// triggersValueTxt.sendKeys(Keys.TAB);

			Thread.sleep(1500);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
			if(editLayoutTriggerRaiseDoc.isSelected()==false)
			{
				editLayoutTriggerRaiseDocSelected.click();
			}

			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
			Select s = new Select(editLayoutTriggersAllVouchers);
			s.selectByVisibleText("Job orders");
			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
			editLayoutTriggersSaveBtn.click();
			Thread.sleep(2500);

			String expSaveMessage = "Data saved successfully";
			String actSaveMessage = checkValidationMessage(expSaveMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();

			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();

			if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
				System.out.println("Test Pass : Trigger  Added Successfully");

				return true;
			} else {
				System.out.println("Test Fail : Trigger NOT Added Successfully");

				return false;
			}


			
	 }
	 
	 
	 public boolean checkSavingPurchaseVoucherVATforLinewiseCheckBoxinTrigger() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		 Thread.sleep(2000);
			
		  getDriver().navigate().refresh();
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu)); 
		  financialsMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu)); 
		  financialsTransactionMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		  financialsTransactionsPurchaseMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat)); 
		  purchaseVouchersVat.click();
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn)); 
		  newBtn.click();
		  
		  checkValidationMessage("Screen opened");
		 

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();

		vendorAccountTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actVendorAccountList = new ArrayList<String>();

		int count = vendorAccountListCount.size();

		System.err.println(count);

		for (int i = 0; i < count; i++) {
			String data = vendorAccountListCount.get(i).getText();

			if (data.equalsIgnoreCase("Vendor A")) {
				vendorAccountListCount.get(i).click();
				break;
			}
		}

		vendorAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		voucherHeaderCurrency.sendKeys(Keys.SPACE);

		int currencycount = currencyListCount.size();

		System.err.println(currencycount);

		for (int i = 0; i < currencycount; i++) {
			String data = currencyListCount.get(i).getText();

			if (data.equalsIgnoreCase("INR")) {
				currencyListCount.get(i).click();
				break;
			}
		}

		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actDepartmentList = new ArrayList<String>();

		int departmentCount = departmentListCount.size();

		System.err.println(departmentCount);

		for (int i = 0; i < departmentCount; i++) {
			String data = departmentListCount.get(i).getText();

			if (data.equalsIgnoreCase("DUBAI")) {
				departmentListCount.get(i).click();
				break;
			}
		}

		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.clear();
		placeOFSupplyTxt.sendKeys("Abu Dhabi");
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys("DUBAI");
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();

		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount = warehouseBodyComboList.size();

		for (int i = 0; i < warehousecount; i++) {
			String data = warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("MUMBAI")) {
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount = pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount; i++) {
			String Item = pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("BR COGS ITEM")) {
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("2");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn)); 
		  select1stRow_14thColumn.click();
		  
		
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable)); 
		  enter_PvTaxable.sendKeys(Keys.TAB);
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch)); 
		  enter_Batch.click(); enter_Batch.sendKeys("PVATR1#V5");
		  enter_Batch.sendKeys(Keys.TAB);
		  
		  DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		  
		  Calendar cal=Calendar.getInstance(); SimpleDateFormat currentDate = new
		  SimpleDateFormat("dd/MM/yyyy"); cal.add(Calendar.DATE, 2);
		  
		  String next2DaysDate=df.format(cal.getTime());
		  
		  System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "
		  +next2DaysDate);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		  enter_Expirydate)); enter_Expirydate.click();
		  enter_Expirydate.sendKeys(Keys.HOME);
		  enter_Expirydate.sendKeys(next2DaysDate);
		  enter_Expirydate.sendKeys(Keys.TAB);
		 

		Thread.sleep(2000);
		
		/*
		 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 * select2ndRow_1stColumn)); select2ndRow_1stColumn.click();
		 */
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount1 = warehouseBodyComboList.size();

		for (int i = 0; i < warehousecount1; i++) {
			String data = warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("MUMBAI")) {
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount1 = pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount1; i++) {
			String Item = pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		select2ndRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("5");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);
		Thread.sleep(2000);	
		
	Thread.sleep(2000);
		

	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn)); 
	  select3rdRow_1stColumn.click();
	 
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount2 = warehouseBodyComboList.size();

		for (int i = 0; i < warehousecount2; i++) {
			String data = warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("MUMBAI")) {
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount2 = pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount2; i++) {
			String Item = pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_9thColumn));
		select3rdRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("5");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_11thColumn));
		select3rdRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		Thread.sleep(3000);
		//checkValidationMessage("");

		String expMessage1 = "Voucher saved successfully";

		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": SU/DUB/TEXT1";

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage1);
		if(actMessage.startsWith(expMessage1) )
		{
			return true;
		}
		else
		{
			return false;
		}
	 }
	 
	 
	 public boolean checkRaisedVoucherafterSavingPurchaseVoucherVATWithoutLineWiseCheckBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		 
		 ArrayList<String> actArry = new ArrayList<String>();
		 actArry=checkPurchaseVoucherVAT("SU/DUB/TEXT1");
		 
			Thread.sleep(8000);
			
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Job Order");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(8000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

			int voucherGridDocNoCount = voucherGridDocNo.size();

			for (int i = 0; i < voucherGridDocNoCount; i++) {
				String data = "6";
				if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
					voucherGridIndexChkBox.get(i).click();
					break;
				}
			}

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
			editBtn.click();
			Thread.sleep(4000);
			
			String actR1Item = select1stRow_1stColumn.getText();
			
			String actR1Units = select1stRow_2ndColumn.getText();
			

			String actR1Quantity = select1stRow_3rdColumn.getText();
			

			String actR1Rate = select1stRow_4thColumn.getText();
			

			String actR1Gross = select1stRow_5thColumn.getText();
			

			

			String actR2Item = select2ndRow_1stColumn.getText();
					
			String actR2Units = select2ndRow_2ndColumn.getText();
			

			String actR2Quantity = select2ndRow_3rdColumn.getText();
			
			String actR2Rate = select2ndRow_4thColumn.getText();
			

			String actR2Gross = select2ndRow_5thColumn.getText();
			
			
			String actR3Item = select3rdRow_1stColumn.getText();
			
			String actR3Units = select3rdRow_2ndColumn.getText();
			

			String actR3Quantity = select3rdRow_3rdColumn.getText();
			
			String actR3Rate = select3rdRow_4thColumn.getText();
			

			String actR3Gross = select3rdRow_5thColumn.getText();
			
			ArrayList<String> expArray = new ArrayList<String>();
			expArray.add(actR1Item);
			expArray.add(actR1Units);
			expArray.add(actR1Quantity);
			expArray.add(actR1Rate);
			expArray.add(actR1Gross);
			
			expArray.add(actR2Item);
			expArray.add(actR2Units);
			expArray.add(actR2Quantity);
			expArray.add(actR2Rate);
			expArray.add(actR2Gross);
			
			expArray.add(actR3Item);
			expArray.add(actR3Units);
			expArray.add(actR3Quantity);
			expArray.add(actR3Rate);
			expArray.add(actR3Gross);
			
			
			String actRowsList=actArry.toString();
			String expRowList=expArray.toString();
			
			System.err.println(actRowsList);
			System.err.println(expRowList);
			
			
			if(actRowsList.equalsIgnoreCase(expRowList))
			{
				return true;
			}
			else
			{
				return false;
			}
			
			
			
			
	 }

	 public ArrayList<String> checkPurchaseVoucherVAT(String DocNo) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {

			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			documentNumberTxt.click();
			documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			documentNumberTxt.sendKeys(DocNo);
			Thread.sleep(2000);
			documentNumberTxt.sendKeys(Keys.TAB);

			checkValidationMessage("Voucher loaded Successfully");

			/*
			 * getWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * documentNumberTxt)); String actDocNo =
			 * documentNumberTxt.getAttribute("value"); String expDocNo = "SU/IND/TEXT2";
			 * 
			 * getWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * vendorAccountTxt)); String actVendor =
			 * vendorAccountTxt.getAttribute("value"); String expVendor = "Vendor B";
			 * 
			 * getWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * voucherHeaderCurrency)); String actCurrency =
			 * voucherHeaderCurrency.getAttribute("value"); String expCurrency = "INR";
			 * 
			 * getWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * departmentTxt)); String actDepartment = departmentTxt.getAttribute("value");
			 * String expDepartment = "INDIA";
			 * 
			 * getWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * placeOFSupplyTxt)); String actPlaceOfSupply =
			 * placeOFSupplyTxt.getAttribute("value"); String expPlaceOfSupply =
			 * "Abu Dhabi";
			 * 
			 * getWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * jurisdictionTxt)); String actJurisdiction =
			 * jurisdictionTxt.getAttribute("value"); String expJurisdiction = "DUBAI";
			 */
			/*
			 * getWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * select1stRow_1stColumn)); String actR1Warehouse =
			 * select1stRow_1stColumn.getText();
			 */
			

			String actR1Item = select1stRow_2ndColumn.getText();
				
			String actR1Units = select1stRow_5thColumn.getText();
			

			String actR1Quantity = select1stRow_9thColumn.getText();
			

			String actR1Rate = select1stRow_11thColumn.getText();
			

			String actR1Gross = select1stRow_12thColumn.getText();
			

			

			String actR2Item = select2ndRow_2ndColumn.getText();
					
			String actR2Units = select2ndRow_5thColumn.getText();
			

			String actR2Quantity = select2ndRow_9thColumn.getText();
			
			String actR2Rate = select2ndRow_11thColumn.getText();
			

			String actR2Gross = select2ndRow_12thColumn.getText();
			
			
			String actR3Item = select3rdRow_2ndColumn.getText();
			
			String actR3Units = select3rdRow_5thColumn.getText();
			

			String actR3Quantity = select3rdRow_9thColumn.getText();
			
			String actR3Rate = select3rdRow_11thColumn.getText();
			

			String actR3Gross = select3rdRow_12thColumn.getText();
			
			ArrayList<String> actArray = new ArrayList<String>();
			actArray.add(actR1Item);
			actArray.add(actR1Units);
			actArray.add(actR1Quantity);
			actArray.add(actR1Rate);
			actArray.add(actR1Gross);
			
			actArray.add(actR2Item);
			actArray.add(actR2Units);
			actArray.add(actR2Quantity);
			actArray.add(actR2Rate);
			actArray.add(actR2Gross);
			
			actArray.add(actR3Item);
			actArray.add(actR3Units);
			actArray.add(actR3Quantity);
			actArray.add(actR3Rate);
			actArray.add(actR3Gross);
			
			return actArray;
	
	 }
	 
	 
	 
	 @FindBy(xpath="//input[@id='id_triggerdef_linewisechecking']")
		private static WebElement linewiseCheckingCheckBox;
	 
	 @FindBy(xpath="//input[@id='id_triggerdef_linewisechecking']//..//span")
		private static WebElement linewiseCheckingCheckBoxSelected;
		
		
		
		@FindBy(xpath="//input[@id='triggers_chksuspend']")
		private static WebElement SuspendCheckBox;
		
		@FindBy(xpath="//input[@id='triggers_chksuspend']//..//span")
		private static WebElement SuspendCheckBoxSelected;
	 
	 public boolean checkEditingTriggertoEnableLinewiseCheckingCheckBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	 {

		    Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Configure Transactions");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(2000);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
			purchaseVouchersVATBtn.click();

			Thread.sleep(6000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();

			Thread.sleep(10000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
			Thread.sleep(1000);
			editLayoutTriggerName.sendKeys("LineWise Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(4000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(linewiseCheckingCheckBoxSelected));
			if(linewiseCheckingCheckBox.isSelected()==false)
			{
				linewiseCheckingCheckBoxSelected.click();
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
			editLayoutTriggersSaveBtn.click();
			Thread.sleep(2500);

			String expSaveMessage = "Data saved successfully";
			String actSaveMessage = checkValidationMessage(expSaveMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();

			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();

			if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
				System.out.println("Test Pass : Trigger  Added Successfully");

				return true;
			} else {
				System.out.println("Test Fail : Trigger NOT Added Successfully");

				return false;
			}

			
	 }
	 
	 
	 public boolean checkRaisedVoucherafterSavingPurchaseVoucherVATWithLineWiseCheckBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {

		 
		 ArrayList<String> actArry = new ArrayList<String>();
		 actArry=checkPurchaseVoucherVAT("SU/DUB/TEXT2");
		 
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Job Order");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

			int voucherGridDocNoCount = voucherGridDocNo.size();

			for (int i = 0; i < voucherGridDocNoCount; i++) {
				String data = "7";
				if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
					voucherGridIndexChkBox.get(i).click();
					break;
				}
			}

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
			editBtn.click();
			Thread.sleep(2000);
			
			String actR1Item = select1stRow_1stColumn.getText();
			
			String actR1Units = select1stRow_2ndColumn.getText();
			

			String actR1Quantity = select1stRow_3rdColumn.getText();
			

			String actR1Rate = select1stRow_4thColumn.getText();
			

			String actR1Gross = select1stRow_5thColumn.getText();
			

			

			String actR2Item = select2ndRow_1stColumn.getText();
					
			String actR2Units = select2ndRow_2ndColumn.getText();
			

			String actR2Quantity = select2ndRow_3rdColumn.getText();
			
			String actR2Rate = select2ndRow_4thColumn.getText();
			

			String actR2Gross = select2ndRow_5thColumn.getText();
			
			
			/*
			 * String actR3Item = select3rdRow_1stColumn.getText();
			 * 
			 * String actR3Units = select3rdRow_2ndColumn.getText();
			 * 
			 * 
			 * String actR3Quantity = select3rdRow_3rdColumn.getText();
			 * 
			 * String actR3Rate = select3rdRow_4thColumn.getText();
			 * 
			 * 
			 * String actR3Gross = select3rdRow_5thColumn.getText();
			 */
			
			ArrayList<String> expArray = new ArrayList<String>();
			expArray.add(actR1Item);
			expArray.add(actR1Units);
			expArray.add(actR1Quantity);
			expArray.add(actR1Rate);
			expArray.add(actR1Gross);
			
			expArray.add(actR2Item);
			expArray.add(actR2Units);
			expArray.add(actR2Quantity);
			expArray.add(actR2Rate);
			expArray.add(actR2Gross);
			
			/*
			 * expArray.add(actR3Item); expArray.add(actR3Units);
			 * expArray.add(actR3Quantity); expArray.add(actR3Rate);
			 * expArray.add(actR3Gross);
			 */
			
			
			
			
			System.err.println(actArry);
			System.err.println(expArray);
			
			boolean res=actArry.retainAll(expArray);
			
			
			if(res)
			{
				return true;
			}
			else
			{
				return false;
			}

	 }
	 
	 @FindBy(xpath="(//*[@class='icon-delete vouchericon'])[2]")
	 public static WebElement editLayoutTriggersDeleteBtn;
	 
	 
	 public boolean checkDeleteTrigger() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		 Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Configure Transactions");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(2000);

			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
			purchaseVouchersVATBtn.click();

			Thread.sleep(6000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();

			Thread.sleep(10000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
			Thread.sleep(1000);
			editLayoutTriggerName.sendKeys("Warehouse Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersDeleteBtn));
			editLayoutTriggersDeleteBtn.click();
			Thread.sleep(2000);
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
			Thread.sleep(2000);
			
			checkValidationMessage("");
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
			Thread.sleep(2000);
			editLayoutTriggerName.sendKeys("Multiple Condition Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersDeleteBtn));
			editLayoutTriggersDeleteBtn.click();
			Thread.sleep(2000);
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
			Thread.sleep(4000);
			
			checkValidationMessage("");
			Thread.sleep(4000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
			Thread.sleep(2000);
			editLayoutTriggerName.sendKeys("LineWise Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersDeleteBtn));
			editLayoutTriggersDeleteBtn.click();
			Thread.sleep(5000);
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
			Thread.sleep(4000);
			
					
			String expMsg="Trigger deleted successfully";
			String actMsg= checkValidationMessage(expMsg);
			
			
			Thread.sleep(2000);
			getAction().moveToElement(updateBtn).build().perform();
			Thread.sleep(2000);
			//ScrollIntoView(updateBtn);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			Thread.sleep(1000);
			String expMsg1="Data saved successfully";
			String actMsg1= checkValidationMessage(expMsg1);
		
			
			if(actMsg.equalsIgnoreCase(expMsg) && actMsg1.equalsIgnoreCase(expMsg1))
			{
				
				
				return true;
			}
			else
			{
			
				
				return false;
			}
	 }
	 
	 public boolean checkSavingTriggertoRaiseaDocumentOnSaveandSuspend() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
		  	Thread.sleep(2000);

			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Configure Transactions");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
			documentCustomization45Btn.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
			purchaseVouchersVATBtn.click();

			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toogleExpandBtn));
			toogleExpandBtn.click();

			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();*/

			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT, Keys.HOME);
			Thread.sleep(2000);
			editLayoutTriggerName.sendKeys("Suspend Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(4000);

			triggersConjuctionDrpdwn.click();
			Thread.sleep(2000);
			triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);
			triggersConjuctionDrpdwn.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
			triggersSelectFieldTxt.click();

			Thread.sleep(3000);

			getAction().moveToElement(triggerItemExpansionBtn).build().perform();
			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerItemExpansionBtn));
			triggerItemExpansionBtn.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerItemName));
			triggerItemName.click();

			Thread.sleep(2000);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
			triggersSelectOperdrpdwn.click();
			triggersSelectOperdrpdwn.sendKeys("Equal");

			Thread.sleep(5000);
			triggersSelectOperdrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(5000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
			triggersCompareWithdrpdwn.click();
			triggersCompareWithdrpdwn.sendKeys("value");
			// triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
			Thread.sleep(2000);
			triggersCompareWithdrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(2000);

		
			triggersValueTxt.sendKeys(Keys.SPACE);
			Thread.sleep(3000);
			triggersValueTxt.sendKeys("STD RATE COGS ITEM");
			Thread.sleep(1500);
			// triggersValueTxt.sendKeys(Keys.TAB);

			Thread.sleep(1500);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
			if(editLayoutTriggerRaiseDoc.isSelected()==false)
			{
				editLayoutTriggerRaiseDocSelected.click();
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SuspendCheckBoxSelected));
			if(SuspendCheckBox.isSelected()==false)
			{
				SuspendCheckBoxSelected.click();
			}

			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
			Select s = new Select(editLayoutTriggersAllVouchers);
			s.selectByVisibleText("Job orders");
			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
			editLayoutTriggersSaveBtn.click();
			Thread.sleep(3000);

			String expSaveMessage = "Data saved successfully";
			String actSaveMessage = checkValidationMessage(expSaveMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();

			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();

			if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
				System.out.println("Test Pass : Trigger  Added Successfully");

				return true;
			} else {
				System.out.println("Test Fail : Trigger NOT Added Successfully");

				return false;
			}


	 }
	 
	 
	 public boolean checkSavingPurchaseVoucherVATtoCheckSuspendStatus() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
		 Thread.sleep(2000);
			
		  getDriver().navigate().refresh();
		  Thread.sleep(4000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu)); 
		  financialsMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu)); 
		  financialsTransactionMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		  financialsTransactionsPurchaseMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat)); 
		  purchaseVouchersVat.click();
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn)); 
		  newBtn.click();
		  
		  checkValidationMessage("Screen opened");
		 

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();

		vendorAccountTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actVendorAccountList = new ArrayList<String>();

		int count = vendorAccountListCount.size();

		System.err.println(count);

		for (int i = 0; i < count; i++) {
			String data = vendorAccountListCount.get(i).getText();

			if (data.equalsIgnoreCase("Vendor A")) {
				vendorAccountListCount.get(i).click();
				break;
			}
		}

		vendorAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		voucherHeaderCurrency.sendKeys(Keys.SPACE);

		int currencycount = currencyListCount.size();

		System.err.println(currencycount);

		for (int i = 0; i < currencycount; i++) {
			String data = currencyListCount.get(i).getText();

			if (data.equalsIgnoreCase("INR")) {
				currencyListCount.get(i).click();
				break;
			}
		}

		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actDepartmentList = new ArrayList<String>();

		int departmentCount = departmentListCount.size();

		System.err.println(departmentCount);

		for (int i = 0; i < departmentCount; i++) {
			String data = departmentListCount.get(i).getText();

			if (data.equalsIgnoreCase("DUBAI")) {
				departmentListCount.get(i).click();
				break;
			}
		}

		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.clear();
		placeOFSupplyTxt.sendKeys("Abu Dhabi");
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys("DUBAI");
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();

		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount = warehouseBodyComboList.size();

		for (int i = 0; i < warehousecount; i++) {
			String data = warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("MUMBAI")) {
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount = pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount; i++) {
			String Item = pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("2");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		
		  
		  

		Thread.sleep(2000);
		

		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		Thread.sleep(3000);
		//checkValidationMessage("");

		String expMessage1 = "Voucher saved successfully";

		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": SU/DUB/TEXT3";

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Vouchr_closeBtn));
		Vouchr_closeBtn.click();
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();
		String actStatus="";
		for (int i = 0; i < voucherGridDocNoCount; i++) {
			String data = "SU/DUB/TEXT3";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				actStatus=voucherGridSuspendStatus.get(i).getText();
				break;
			}
		}
		String expStatus="True";
		
		System.err.println(actStatus);
		
		if(actMessage.startsWith(expMessage1) &&  actStatus.equalsIgnoreCase(expStatus))
		{
			return true;
		}
		else
		{
			return false;
		}
	 }
	 
	 @FindBy(xpath="//tbody[@id='tblBodyTransRender']/tr/td[12]")
		public static List<WebElement> voucherGridSuspendStatus;
	 
	 public boolean checkEditingTriggertoEnableLineWiseCheckBoxtoSuspend() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		    Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Configure Transactions");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(2000);

			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
			purchaseVouchersVATBtn.click();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();

			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT, Keys.HOME);
			Thread.sleep(1000);
			editLayoutTriggerName.sendKeys("Suspend Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(4000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(linewiseCheckingCheckBoxSelected));
			if(linewiseCheckingCheckBox.isSelected()==false)
			{
				linewiseCheckingCheckBoxSelected.click();
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
			editLayoutTriggersSaveBtn.click();
			Thread.sleep(3000);

			String expSaveMessage = "Data saved successfully";
			String actSaveMessage = checkValidationMessage(expSaveMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();

			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();

			if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
				System.out.println("Test Pass : Trigger  Added Successfully");

				return true;
			} else {
				System.out.println("Test Fail : Trigger NOT Added Successfully");

				return false;
			}

	 }
	 
	 public boolean checkSavingPuchaseVoucherVATtoCheckStausofSuspendwhenLinewiseChkBoxEnabled() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {

		 Thread.sleep(2000);
			
		  getDriver().navigate().refresh();
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu)); 
		  financialsMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu)); 
		  financialsTransactionMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		  financialsTransactionsPurchaseMenu.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat)); 
		  purchaseVouchersVat.click();
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn)); 
		  newBtn.click();
		  
		  checkValidationMessage("Screen opened");
		 

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();

		vendorAccountTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actVendorAccountList = new ArrayList<String>();

		int count = vendorAccountListCount.size();

		System.err.println(count);

		for (int i = 0; i < count; i++) {
			String data = vendorAccountListCount.get(i).getText();

			if (data.equalsIgnoreCase("Vendor A")) {
				vendorAccountListCount.get(i).click();
				break;
			}
		}

		vendorAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		voucherHeaderCurrency.sendKeys(Keys.SPACE);

		int currencycount = currencyListCount.size();

		System.err.println(currencycount);

		for (int i = 0; i < currencycount; i++) {
			String data = currencyListCount.get(i).getText();

			if (data.equalsIgnoreCase("INR")) {
				currencyListCount.get(i).click();
				break;
			}
		}

		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actDepartmentList = new ArrayList<String>();

		int departmentCount = departmentListCount.size();

		System.err.println(departmentCount);

		for (int i = 0; i < departmentCount; i++) {
			String data = departmentListCount.get(i).getText();

			if (data.equalsIgnoreCase("DUBAI")) {
				departmentListCount.get(i).click();
				break;
			}
		}

		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.clear();
		placeOFSupplyTxt.sendKeys("Abu Dhabi");
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys("DUBAI");
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();

		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount = warehouseBodyComboList.size();

		for (int i = 0; i < warehousecount; i++) {
			String data = warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("MUMBAI")) {
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount = pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount; i++) {
			String Item = pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("BR COGS ITEM")) {
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("2");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn)); 
		  select1stRow_14thColumn.click();
		  
		
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable)); 
		  enter_PvTaxable.sendKeys(Keys.TAB);
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch)); 
		  enter_Batch.click(); enter_Batch.sendKeys("PVATR1#V5");
		  enter_Batch.sendKeys(Keys.TAB);
		  
		  DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		  
		  Calendar cal=Calendar.getInstance(); SimpleDateFormat currentDate = new
		  SimpleDateFormat("dd/MM/yyyy"); cal.add(Calendar.DATE, 2);
		  
		  String next2DaysDate=df.format(cal.getTime());
		  
		  System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "
		  +next2DaysDate);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		  enter_Expirydate)); enter_Expirydate.click();
		  enter_Expirydate.sendKeys(Keys.HOME);
		  enter_Expirydate.sendKeys(next2DaysDate);
		  enter_Expirydate.sendKeys(Keys.TAB);
		 

		Thread.sleep(2000);
		
		/*
		 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 * select2ndRow_1stColumn)); select2ndRow_1stColumn.click();
		 */
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount1 = warehouseBodyComboList.size();

		for (int i = 0; i < warehousecount1; i++) {
			String data = warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("MUMBAI")) {
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount1 = pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount1; i++) {
			String Item = pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		select2ndRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("5");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);
		Thread.sleep(2000);	
		
	Thread.sleep(2000);
		

	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn)); 
	  select3rdRow_1stColumn.click();
	 
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.click();
		enter_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_WarehouseTxt.sendKeys(Keys.SPACE);

		int warehousecount2 = warehouseBodyComboList.size();

		for (int i = 0; i < warehousecount2; i++) {
			String data = warehouseBodyComboList.get(i).getText();

			if (data.equalsIgnoreCase("MUMBAI")) {
				warehouseBodyComboList.get(i).click();
				break;
			}
		}

		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		int pvvGridItemListCount2 = pvvGridItemList.size();
		for (int i = 0; i < pvvGridItemListCount2; i++) {
			String Item = pvvGridItemList.get(i).getText();
			if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
		enter_TaxCode.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_9thColumn));
		select3rdRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("5");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_11thColumn));
		select3rdRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		Thread.sleep(3000);
		//checkValidationMessage("");

		String expMessage1 = "Voucher saved successfully";

		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": SU/DUB/TEXT4";

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Vouchr_closeBtn));
		Vouchr_closeBtn.click();
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();
		String actStatus="";
		for (int i = 0; i < voucherGridDocNoCount; i++) {
			String data = "SU/DUB/TEXT4";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				actStatus=voucherGridSuspendStatus.get(i).getText();
				break;
			}
		}
		String expStatus="False";
		
		System.err.println(actStatus);
		
		if(actMessage.startsWith(expMessage1) &&  actStatus.equalsIgnoreCase(expStatus))
		{
			return true;
		}
		else
		{
			return false;
		}
	 
	 }
	 
	 @FindBy(xpath="//*[@class='icon-close d-none d-md-block hiconright2']")
	public static WebElement Vouchr_closeBtn;
	 
	 @FindBy(xpath = "//a[@id='2560']")
		private static WebElement purchaseOrdersBtn;
	 
	 public boolean checkSaveTriggerOptionasOnAuthorizationinPurchaseOrder()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Configure Transactions");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(6000);

			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseOrdersBtn));
			purchaseOrdersBtn.click();

			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();

			Thread.sleep(5000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
			Thread.sleep(1000);
			editLayoutTriggerName.sendKeys("Authorize Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(4000);

			triggersConjuctionDrpdwn.click();
			Thread.sleep(2000);
			triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);
			triggersConjuctionDrpdwn.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
			triggersSelectFieldTxt.click();

			Thread.sleep(3000);

			getAction().moveToElement(triggerVendorACExpansionBtn).build().perform();
			Thread.sleep(1000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorACExpansionBtn));
			triggerVendorACExpansionBtn.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorAcName));
			triggerVendorAcName.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
			triggersSelectOperdrpdwn.click();
			triggersSelectOperdrpdwn.sendKeys("Equal");

			Thread.sleep(5000);
			triggersSelectOperdrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(5000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
			triggersCompareWithdrpdwn.click();
			triggersCompareWithdrpdwn.sendKeys("value");
			// triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
			Thread.sleep(2000);
			triggersCompareWithdrpdwn.sendKeys(Keys.TAB);

			Thread.sleep(2000);

			
			triggersValueTxt.sendKeys(Keys.SPACE);
			Thread.sleep(3000);
			triggersValueTxt.sendKeys("Vendor B");
			Thread.sleep(1500);
			triggersValueTxt.sendKeys(Keys.TAB);

			Thread.sleep(1500);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
			if(editLayoutTriggerRaiseDoc.isSelected()==false)
			{
				editLayoutTriggerRaiseDocSelected.click();
			}

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersOnAuthorizatinRadioBtn));
			triggersOnAuthorizatinRadioBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
			Select s = new Select(editLayoutTriggersAllVouchers);
			s.selectByVisibleText("Purchases Vouchers");
			Thread.sleep(3000);

			
			getAction().moveToElement(editLayoutTriggersSaveBtn).build().perform();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
			editLayoutTriggersSaveBtn.click();
			Thread.sleep(3000);

			String expSaveMessage = "Data saved successfully";
			String actSaveMessage = checkValidationMessage(expSaveMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			

			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);

			
			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();

			if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
				System.out.println("Test Pass : Trigger  Added Successfully");

				return true;
			} else {
				System.out.println("Test Fail : Trigger NOT Added Successfully");

				return false;
			}

		}
	 
	 @FindBy(xpath = "//a[@id='137']//span[contains(text(),'Transactions')]")
		public static WebElement invTransactionsMenu;

		public boolean checkSavingPurchaseOrdersVouchertoTrigger() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
		{
			
			Thread.sleep(6000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
			invTransactionsMenu.click();
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
			inventoryTransactionsPurchasesMenu.click();
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
			purchasesOrdersVoucher.click();
			
			Thread.sleep(8000);
			
			new WebDriverWait(getDriver(), 300).until(ExpectedConditions.visibilityOf(newBtn));
			
			//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
			MRpurchaseAccountTxt.click();
			MRpurchaseAccountTxt.sendKeys("STD RATE COGS ACC INV");
			Thread.sleep(3000);
			MRpurchaseAccountTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.click();
			vendorAccountTxt.sendKeys("Vendor B");
			Thread.sleep(3000);
			vendorAccountTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValuetxt));
			departmentValuetxt.click();
			departmentValuetxt.sendKeys("INDIA");
			Thread.sleep(3000);
			departmentValuetxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);

			Thread.sleep(2000);

			int pvvGridItemListCount = pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) {
				String Item = pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("2");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
			select1stRow_4thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
			select2ndRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.click();

			enter_ItemTxt.sendKeys(Keys.END);
			enter_ItemTxt.sendKeys(Keys.SHIFT, Keys.HOME);
			enter_ItemTxt.sendKeys("STD RATE COGS ITEM");

			Thread.sleep(2000);

			enter_ItemTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
			select2ndRow_3rdColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("1");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
			select2ndRow_4thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);

			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
			voucherSaveBtn.click();

			String expMessage1 = "Voucher saved successfully";

			String actMessage = checkValidationMessage(expMessage1);
			String expMessage2 = ": 1";

			System.out.println("Actual Message    : " + actMessage);
			System.out.println("Expected Message  : " + expMessage1);

			if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2))

			{
				return true;
			} else {
				return false;
			}

		}
	 
		public boolean checkVerifyingTriggerOptioninPurchaseVoucherBeforeAuthorization() throws InterruptedException {

			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesExpandBtn));
			purchasesExpandBtn.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersBtn));
			purchaseVouchersBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

			Calendar cal = Calendar.getInstance();
			SimpleDateFormat todayDate = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = todayDate.format(cal.getTime());
			System.out.println("--------------Date Before Adding:: " + currentDate);

			int voucherGridDocNoCount = voucherGridDocNo.size();
			boolean flag=true;
			for (int i = 0; i < voucherGridDocNoCount; i++) {
				String data = "4";
				if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
					flag=false;
					break;
				}
			}

			//boolean result = voucherHomeBodyList.getText().isEmpty();
			if (flag) {
				return true;
			} else {
				return false;
			}

		}
		
		
		@FindBy(xpath = "//*[@id='id_transactionentry_authorize']/a/span")
		private static WebElement POauthorizeBtn;

		@FindBy(xpath = "//*[@id='id_transaction_homescreen_authorize']")
		private static WebElement authorizeBtn;

		@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr[1]/td[13]")
		private static WebElement openingStocksNewHomeRow1AuthorizationStatus;

		@FindBy(xpath = "(//*[@id='1'])[1]")
		private static WebElement alertTransactions;

		@FindBy(xpath = "//*[@id='2560']")
		private static WebElement transactionAuthorizations;

		@FindBy(xpath = "//*[@id='masterChild_2560']")
		private static WebElement transactionAuthorizationText1;

		public static boolean checkAuthorizationInPurchaseOrderVoucherWithUserAllOptionsSTFromEntryPage()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			Thread.sleep(2000);
			
			/*getWebDriverWait().until(ExpectedConditions.visibilityOf(alertsTab));
			alertsTab.click();
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.visibilityOf(alertTransactions));
			alertTransactions.click();
			Thread.sleep(2000);
			
			getWebDriverWait().until(ExpectedConditions.visibilityOf(transactionAuthorizations));
			transactionAuthorizations.click();
			Thread.sleep(2000);
			
			getAction().moveToElement(transactionAuthorizationText1).build().perform();
			getWebDriverWait().until(ExpectedConditions.visibilityOf(transactionAuthorizationText1));
			transactionAuthorizationText1.click();
			Thread.sleep(2000);
			
			checkValidationMessage("Voucher loaded Successfully");

			Thread.sleep(3000);*/
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
			inventoryTransactionsPurchasesMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
			purchasesOrdersVoucher.click();

			Thread.sleep(3000);
			new WebDriverWait(getDriver(), 300).until(ExpectedConditions.visibilityOf(newBtn));

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
			getAction().doubleClick(grid_ChkBox1).build().perform();

			checkValidationMessage("Voucher loaded Successfully");
			
			Thread.sleep(2000);

			
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(POauthorizeBtn));
			POauthorizeBtn.click();

			if (getIsAlertPresent()) {
				getAlert().dismiss();

				getDriver().navigate().refresh();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
				invTransactionsMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
				inventoryTransactionsPurchasesMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
				purchasesOrdersVoucher.click();

				Thread.sleep(3000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

				int vcount = grid_VoucherNoList.size();

				for (int i = 0; i < vcount; i++) {
					String VoucherNo = grid_VoucherNoList.get(i).getText();

					if (VoucherNo.equalsIgnoreCase("2")) {
						if (grid_CheckBoxList.get(i).isSelected() == false) {
							grid_CheckBoxList.get(i).click();
							break;
						}
					}
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
				authorizeBtn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
				newBtn.click();

				checkValidationMessage("Screen Opened");
			}
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
			new_CloseBtn.click();

			Thread.sleep(2000);

			getFluentWebDriverWait()
					.until(ExpectedConditions.elementToBeClickable(openingStocksNewHomeRow1AuthorizationStatus));

			String actStatus = openingStocksNewHomeRow1AuthorizationStatus.getText();
			String expStatus = "Authorized";

			System.out.println("Authorization Status In home Page 	 :  " + actStatus + " Value Expected : " + expStatus);

			if (actStatus.equalsIgnoreCase(expStatus)) {
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();

				return true;
			} else {
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
				invTransactionsMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
				inventoryTransactionsPurchasesMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
				purchasesOrdersVoucher.click();

				Thread.sleep(3000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

				int vcount = grid_VoucherNoList.size();

				for (int i = 0; i < vcount; i++) {
					String VoucherNo = grid_VoucherNoList.get(i).getText();

					if (VoucherNo.equalsIgnoreCase("1")) {
						if (grid_CheckBoxList.get(i).isSelected() == false) {
							getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
							break;
						}
					}
				}

				checkValidationMessage("Voucher loaded Successfully");

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(POauthorizeBtn));
				POauthorizeBtn.click();

				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
				new_CloseBtn.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();

				return false;
			}
		}

		
		@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr[1]/td")
		private static List<WebElement> voucherHomeBodyLists;

		public boolean checkVerifingDetailsOfSavedPurchaseVouchersInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesExpandBtn));
			purchasesExpandBtn.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersBtn));
			purchaseVouchersBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

			Calendar cal = Calendar.getInstance();
			SimpleDateFormat todayDate = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = todayDate.format(cal.getTime());
			System.out.println("--------------Date Before Adding:: " + currentDate);

			int count = voucherHomeBodyLists.size();

			ArrayList<String> actVouchersList = new ArrayList<String>();

			for (int i = 0; i < count; i++) {
				String data = voucherHomeBodyLists.get(i).getText();

				if (data.isEmpty() == false && i != 9 && i != 10) {
					System.out.println(data);

					actVouchersList.add(data);
				}
			}

			ArrayList<String> expVouchersList = new ArrayList<String>();

			expVouchersList.add("1");
			expVouchersList.add(currentDate);
			expVouchersList.add("2");
			expVouchersList.add("UserAllOptionsST");
			expVouchersList.add("UserAllOptionsST");
			expVouchersList.add(currentDate);
			expVouchersList.add(currentDate);
			expVouchersList.add("True");
			expVouchersList.add("Authorized");
		/*	expVouchersList.add(" ");
			expVouchersList.add("Total");*/

			System.out.println(
					"**********************checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage*********************");

			System.out.println(actVouchersList);
			System.out.println(expVouchersList);

			if (actVouchersList.equals(expVouchersList)) {
				return true;
			} else {
				return false;
			}
		}
		
		
		public boolean checkCreateEmailTemplateforPurchaseOrderVoucher()
				throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException {

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Configure Transactions");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(8000);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseOrdersBtn));
			purchaseOrdersBtn.click();

			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();

			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
			Thread.sleep(1000);
			editLayoutTriggerName.sendKeys("Authorize Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
			if(sendEmailChkBox.isSelected()==false)
			{
				sendEmailChkBoxSelected.click();
			}
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createNewEmailButton));
			createNewEmailButton.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailNewBtn));
			emailNewBtn.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailNewLayoutBtn));
			emailNewLayoutBtn.click();
			Thread.sleep(2000);

			/*Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\POTriggersEmailLayout.exe");

			Thread.sleep(135000);*/
			
			
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\POTriggerEmailInvoiceLayout.exe");
			
			Thread.sleep(150000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveBtn));
			emailSaveBtn.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveOptionBtn));
			emailSaveOptionBtn.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSavePopupTitle));
			boolean actTitleDisplay = emailSavePopupTitle.isDisplayed();
			boolean expTitleDisplay = true;

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveTxt));
			emailSaveTxt.click();
			emailSaveTxt.sendKeys("POTriggersEmailLayout");
			Thread.sleep(1500);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveLayoutBtn));
			emailSaveLayoutBtn.click();
			Thread.sleep(2000);
			String expSaveLayotMsg = "Layout saved successfully";
			String actSaveLayoutMsg = checkValidationMessage(expSaveLayotMsg);

			System.out.println("Actual Message: 	" + actSaveLayoutMsg);
			System.out.println("Expected Message:	" + expSaveLayotMsg);
/*
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailLayoutPreviewBtn));
			emailLayoutPreviewBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailLayoutDesignBtn));
			emailLayoutDesignBtn.click();
			Thread.sleep(2000);
*/
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailLayoutExitBtn));
			emailLayoutExitBtn.click();
			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Configure Transactions");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(8000);

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseOrdersBtn));
			purchaseOrdersBtn.click();

			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();

			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
			Thread.sleep(1000);
			editLayoutTriggerName.sendKeys("Authorize Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
			if(sendEmailChkBox.isSelected()==false)
			{
				sendEmailChkBoxSelected.click();
			}
			Thread.sleep(2000);

			
			
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailTxtBox));
			sendEmailTxtBox.click();
			sendEmailTxtBox.sendKeys(Keys.HOME, Keys.SHIFT, Keys.END);
			sendEmailTxtBox.sendKeys("POTriggersEmailLayout");
			Thread.sleep(3000);
			
			sendEmailTxtBox.sendKeys(Keys.TAB);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSubjectTxt));
			emailSubjectTxt.sendKeys("Email Trigger Voucher For Authorization");
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateBtn));
			emailHeaderTemplateBtn.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newButtonInCreateTemplate));
			newButtonInCreateTemplate.click();
			Thread.sleep(1500);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplates));
			emailTemplates.sendKeys("EmailHeaderTriggerTemplate For PO");
			emailTemplates.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
			companyNameTemplateArea.click();
			Thread.sleep(1500);

			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\POEmailHeaderTemplate.exe");
			Thread.sleep(150000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
			saveBtnInEmailTemplateHeader.click();

			Thread.sleep(2000);
			System.err.println("Template save button clicked");

			String expEmailTemplateSaveMsg = "Template saved successfully";
			String actEmailTemplateSaveMsg = checkValidationMessage(expEmailTemplateSaveMsg);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
			ClickUsingJs(closeButtonInCreateTemplate);
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateTxt));
			emailHeaderTemplateTxt.click();
			emailHeaderTemplateTxt.sendKeys("EmailHeaderTriggerTemplate For PO");
			Thread.sleep(4000);
			emailHeaderTemplateTxt.sendKeys(Keys.TAB);
			Thread.sleep(4000);
			
			getAction().moveToElement(customerChkBoxSelected).build().perform();
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerChkBoxSelected));
			if (customerChkBox.isSelected() == false) {
				customerChkBoxSelected.click();
			}

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
			editLayoutTriggersSaveBtn.click();
			Thread.sleep(2500);

			String expSaveMessage = "Data saved successfully";
			String actSaveMessage = checkValidationMessage(expSaveMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();

			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);

			
			Thread.sleep(10000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();
			Thread.sleep(5000);

			if (actTitleDisplay == expTitleDisplay && actSaveLayoutMsg.equalsIgnoreCase(expSaveLayotMsg)) {
				return true;
			} else {
				return false;
			}

		}
		
		
		public boolean checkSavingPurchaseOrderVouchertoAuthorize()
				throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{

			getDriver().navigate().refresh();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
			invTransactionsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
			inventoryTransactionsPurchasesMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
			purchasesOrdersVoucher.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
			MRpurchaseAccountTxt.click();
			MRpurchaseAccountTxt.sendKeys("STD RATE COGS ACC INV");
			Thread.sleep(3000);
			MRpurchaseAccountTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.click();
			vendorAccountTxt.sendKeys("Vendor B");
			Thread.sleep(3000);
			vendorAccountTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValuetxt));
			departmentValuetxt.click();
			departmentValuetxt.sendKeys("INDIA");
			Thread.sleep(3000);
			departmentValuetxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);

			Thread.sleep(2000);

			int pvvGridItemListCount = pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) {
				String Item = pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("2");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
			select1stRow_4thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
			select2ndRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.click();

			enter_ItemTxt.sendKeys(Keys.END);
			enter_ItemTxt.sendKeys(Keys.SHIFT, Keys.HOME);
			enter_ItemTxt.sendKeys("STD RATE COGS ITEM");

			Thread.sleep(2000);

			enter_ItemTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
			select2ndRow_3rdColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("1");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
			select2ndRow_4thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
			voucherSaveBtn.click();

			String expMessage1 = "Voucher saved successfully";

			String actMessage = checkValidationMessage(expMessage1);
			// String expMessage2 = ": 3";
			
			Thread.sleep(2000);

			System.out.println("Actual Message    : " + actMessage);
			System.out.println("Expected Message  : " + expMessage1);

			if (actMessage.startsWith(expMessage1)/* && actMessage.endsWith(expMessage2) */ )

			{
				return true;
			} else {
				return false;
			}

		}

		
		public boolean checkVerifingDetailsOfSavedPurchaseVouchers2InHomepage()
				throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesExpandBtn));
			purchasesExpandBtn.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersBtn));
			purchaseVouchersBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

			Calendar cal = Calendar.getInstance();
			SimpleDateFormat todayDate = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = todayDate.format(cal.getTime());
			System.out.println("--------------Date Before Adding:: " + currentDate);

			int count = voucherHomeBodyLists2nd.size();

			ArrayList<String> actVouchersList = new ArrayList<String>();

			for (int i = 0; i < count; i++) {
				String data = voucherHomeBodyLists2nd.get(i).getText();

				if (data.isEmpty() == false && i != 9 && i != 10) {
					System.out.println(data);

					actVouchersList.add(data);
				}
			}

			ArrayList<String> expVouchersList = new ArrayList<String>();

			expVouchersList.add("2");
			expVouchersList.add(currentDate);
			expVouchersList.add("2");
			expVouchersList.add("UserAllOptionsST");
			expVouchersList.add("UserAllOptionsST");
			expVouchersList.add(currentDate);
			expVouchersList.add(currentDate);
			expVouchersList.add("True");
			expVouchersList.add("Authorized");
		/*	expVouchersList.add(" ");
			expVouchersList.add("Total");*/

			System.out.println(
					"**********************checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage*********************");

			System.out.println(actVouchersList);
			System.out.println(expVouchersList);

			if (actVouchersList.equals(expVouchersList)) {
				return true;
			} else {
				return false;
			}
		}

		
		@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr[2]/td")
		private static List<WebElement> voucherHomeBodyLists2nd;
		
		
		public boolean checkEmailAfterAuthorizationofPurchaseOrderVoucherFromEntryPage()
				throws InterruptedException, AWTException, IOException {

			Thread.sleep(4000);
			File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\POTriggersEmail.pdf");

			if (Efile1.exists()) {
				Efile1.delete();
			}

			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_T);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			robot.keyRelease(KeyEvent.VK_T);

			Thread.sleep(6000);

			ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

			System.out.println("openTabs" + openTabs);

			getDriver().switchTo().window(openTabs.get(0));
			Thread.sleep(2000);
			getDriver().switchTo().window(openTabs.get(1));
			Thread.sleep(2000);
			getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
			userNameTxt.click();

			userNameTxt.sendKeys("emailvalidationone@gmail.com");

			Thread.sleep(2000);

			System.out.println("User enter text");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
			NextBtn.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
			PasswordTxt.click();

			PasswordTxt.sendKeys("validationone");

			Thread.sleep(2000);

			NextBtn.click();

			Thread.sleep(8000);

			int count1 = row1mailFromList.size();
			//System.out.println("No.of mails in the list" + count1);

			//System.err.println("Email Body *" + emailBody1.getText() + "*");

			if (emailBody2.getText().isEmpty() == false) {
				int count = emailSubjectList.size();
				System.err.println("No.of mails in the list" + count);

				for (int i = 0; i < count; i++) {
					String mailFromTxt = emailSubjectList.get(i).getText();

					if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) {
						emailSubjectList.get(i).click();
						break;
					}

				}

				Thread.sleep(3000);

				Thread.sleep(3000);

				getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
				mailAttachmentDownloadBtn1.click();

				Thread.sleep(3000);

				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_J);
				robot.keyRelease(KeyEvent.VK_J);
				robot.keyRelease(KeyEvent.VK_CONTROL);

				Thread.sleep(5000);

				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);

				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);

				Thread.sleep(3000);

				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_S);
				robot.keyRelease(KeyEvent.VK_S);
				robot.keyRelease(KeyEvent.VK_CONTROL);

				Thread.sleep(2000);

				Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\POTriggersEmail.exe");

				Thread.sleep(6000);

				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_J);
				robot.keyRelease(KeyEvent.VK_J);
				robot.keyRelease(KeyEvent.VK_CONTROL);

				Thread.sleep(2000);

				ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

				int actOpenWindowsCount = getDriver().getWindowHandles().size();
				int expOpenWindowsCount = 4;

				System.out.println(
						"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

				getDriver().switchTo().window(newTabs.get(3)).close();
				Thread.sleep(1000);
				getDriver().switchTo().window(newTabs.get(2)).close();
				Thread.sleep(1000);
				getDriver().switchTo().window(newTabs.get(1));
				Thread.sleep(1000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
				mailDeleteButton.click();

				Thread.sleep(2000);

				Thread.sleep(2000);

			/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailUserBtn));
				gmailUserBtn.click();*/
				
				ClickUsingJs(gmailUserBtn);

				Thread.sleep(3000);
				getDriver().switchTo().frame(SignOutFrame);
				Thread.sleep(1000);
				getAction().moveToElement(gmailSignOutBtn).click().build().perform();
				Thread.sleep(6000);
				
				getDriver().get(
						"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
				removeAccountBtn.click();

				Thread.sleep(1000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
				removeDeleteBtn.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
				yesRemoveBtn.click();

				Thread.sleep(2000);

				getDriver().switchTo().window(newTabs.get(1)).close();
				Thread.sleep(2000);
				getDriver().switchTo().window(newTabs.get(0));
				Thread.sleep(2000);

				String actAccount1PDF = getBaseDir() + "\\autoIt\\ExportFiles\\POTriggersEmail.pdf";
				String expAccount1PDF = getBaseDir() + "\\autoIt\\ImportFiles\\POTriggersEmail.pdf";

				PDFUtil pdfutil = new PDFUtil();

				boolean result = pdfutil.compare(actAccount1PDF, expAccount1PDF);
				Calendar cal = Calendar.getInstance();
				DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
				String currentDate = df.format(cal.getTime());

				String actAccount1Data = pdfutil.getText(actAccount1PDF);
				String expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("26/12/2022", currentDate);

				System.err.println(actAccount1Data);
				System.err.println(expAccount1Data);

				if (actAccount1Data.equalsIgnoreCase(expAccount1Data)) {

					System.err.println("PDF file is as Expected");
					return true;
				} else {

					System.err.println("PDF file is not as Expected");
					return false;
				}

			} else {

				System.err.println("Subject : " + emailRow1Subject.getText());

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsChckbx));
				selectAllMailsChckbx.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsDltBtn));
				selectAllMailsDltBtn.click();

				Thread.sleep(4000);

				Thread.sleep(2000);

			/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailUserBtn));
				gmailUserBtn.click();
*/
				ClickUsingJs(gmailUserBtn);
				Thread.sleep(3000);
				getDriver().switchTo().frame(SignOutFrame);
				Thread.sleep(1000);
				getAction().moveToElement(gmailSignOutBtn).click().build().perform();
				Thread.sleep(6000);
				
				getDriver().get(
						"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
				removeAccountBtn.click();

				Thread.sleep(1000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
				removeDeleteBtn.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
				yesRemoveBtn.click();

				Thread.sleep(2000);
				ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

				getDriver().switchTo().window(newTabs.get(1)).close();
				Thread.sleep(2000);
				getDriver().switchTo().window(newTabs.get(0));
				Thread.sleep(2000);

				return false;
			}

		}
		
		/*public boolean checkSavingPurchaseOrderVouchertoAuthorize()
				throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {

			getDriver().navigate().refresh();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
			invTransactionsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
			inventoryTransactionsPurchasesMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
			purchasesOrdersVoucher.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
			MRpurchaseAccountTxt.click();
			MRpurchaseAccountTxt.sendKeys("STD RATE COGS ACC INV");
			Thread.sleep(3000);
			MRpurchaseAccountTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
			vendorAccountTxt.click();
			vendorAccountTxt.sendKeys("Vendor B");
			Thread.sleep(3000);
			vendorAccountTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValuetxt));
			departmentValuetxt.click();
			departmentValuetxt.sendKeys("INDIA");
			Thread.sleep(3000);
			departmentValuetxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.click();
			enter_ItemTxt.sendKeys(Keys.SPACE);

			Thread.sleep(2000);

			int pvvGridItemListCount = pvvGridItemList.size();
			for (int i = 0; i < pvvGridItemListCount; i++) {
				String Item = pvvGridItemList.get(i).getText();
				if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
					pvvGridItemList.get(i).click();
					break;
				}
			}
			enter_ItemTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			select1stRow_3rdColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("2");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
			select1stRow_4thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
			select2ndRow_1stColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
			enter_ItemTxt.click();

			enter_ItemTxt.sendKeys(Keys.END);
			enter_ItemTxt.sendKeys(Keys.SHIFT, Keys.HOME);
			enter_ItemTxt.sendKeys("STD RATE COGS ITEM");

			Thread.sleep(2000);

			enter_ItemTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
			select2ndRow_3rdColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys("1");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
			select2ndRow_4thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);

			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
			voucherSaveBtn.click();

			String expMessage1 = "Voucher saved successfully";

			String actMessage = checkValidationMessage(expMessage1);
			// String expMessage2 = ": 3";

			System.out.println("Actual Message    : " + actMessage);
			System.out.println("Expected Message  : " + expMessage1);

			if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)  )

			{
				return true;
			} else {
				return false;
			}

		}
		*/
		
		@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr//td[13]")
		public static List<WebElement> authorizationStatusList;

		@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr//td[2]")
		public static WebElement secondColCheckBox;
		
		public boolean checkAuthorizationofPurchaseOrderinHomePageatUserALLOptions() throws InterruptedException {

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
			invTransactionsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
			inventoryTransactionsPurchasesMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
			purchasesOrdersVoucher.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

			int vcount = authorizationStatusList.size();

			for (int i = 0; i < vcount; i++) {
				String VoucherNo = authorizationStatusList.get(i).getText();

				if (VoucherNo.equalsIgnoreCase("Pending")) {
					if (grid_CheckBoxList.get(i).isSelected() == false) {
						grid_CheckBoxList.get(i).click();
						break;
					}
				}
			}

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
			authorizeBtn.click();
			Thread.sleep(5000);

			getFluentWebDriverWait()
					.until(ExpectedConditions.elementToBeClickable(openingStocksNewHomeRow1AuthorizationStatus));

			String actStatus = openingStocksNewHomeRow1AuthorizationStatus.getText();
			String expStatus = "Authorized";

			System.out.println("Authorization Status In home Page 	 :  " + actStatus + " Value Expected : " + expStatus);

			if (actStatus.equalsIgnoreCase(expStatus)) {
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
				voucherhomeCloseBtn.click();

				return true;
			}

			else {
				return false;
			}
		}
		
		@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr[1]/td")
		private static List<WebElement> voucherHomeBodyLists1st;
		
		@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr[3]/td")
		private static List<WebElement> voucherHomeBodyLists3rd;

		@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr[4]/td")
		private static List<WebElement> voucherHomeBodyLists4th;
		
		public boolean checkVerifingDetailsOfSavedPurchaseVouchers3InHomepage()
				throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesExpandBtn));
			purchasesExpandBtn.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersBtn));
			purchaseVouchersBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

			Calendar cal = Calendar.getInstance();
			SimpleDateFormat todayDate = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = todayDate.format(cal.getTime());
			System.out.println("--------------Date Before Adding:: " + currentDate);

			int count = voucherHomeBodyLists3rd.size();

			ArrayList<String> actVouchersList = new ArrayList<String>();

			for (int i = 0; i < count; i++) {
				String data = voucherHomeBodyLists3rd.get(i).getText();

				if (data.isEmpty() == false && i != 9 && i != 10) {
					System.out.println(data);

					actVouchersList.add(data);
				}
			}

			ArrayList<String> expVouchersList = new ArrayList<String>();

			expVouchersList.add("3");
			expVouchersList.add(currentDate);
			expVouchersList.add("2");
			expVouchersList.add("UserAllOptionsST");
			expVouchersList.add("UserAllOptionsST");
			expVouchersList.add(currentDate);
			expVouchersList.add(currentDate);
			expVouchersList.add("True");
			expVouchersList.add("Authorized");
		/*	expVouchersList.add(" ");
			expVouchersList.add("Total");*/

			System.out.println(
					"**********************checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage*********************");

			System.out.println(actVouchersList);
			System.out.println(expVouchersList);

			if (actVouchersList.equals(expVouchersList)) {
				return true;
			} else {
				return false;
			}
		}
		
		
		public boolean checkEmailAfterAuthorizationofPurchaseOrderVoucherfromHomePage()
				throws InterruptedException, AWTException, IOException {

			Thread.sleep(2000);
			File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\POTriggersEmailHome.pdf");

			if (Efile1.exists()) {
				Efile1.delete();
			}

			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_T);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			robot.keyRelease(KeyEvent.VK_T);

			Thread.sleep(6000);

			ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

			System.out.println("openTabs" + openTabs);

			getDriver().switchTo().window(openTabs.get(0));
			Thread.sleep(2000);
			getDriver().switchTo().window(openTabs.get(1));
			Thread.sleep(2000);
			getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
			userNameTxt.click();

			userNameTxt.sendKeys("emailvalidationone@gmail.com");

			Thread.sleep(2000);

			System.out.println("User enter text");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
			NextBtn.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
			PasswordTxt.click();

			PasswordTxt.sendKeys("validationone");

			Thread.sleep(2000);

			NextBtn.click();

			Thread.sleep(8000);

			int count1 = row1mailFromList.size();
			//System.out.println("No.of mails in the list" + count1);

		//	System.err.println("Email Body *" + emailBody.getText() + "*");

			if (emailBody2.getText().isEmpty() == false) {
				int count = emailSubjectList.size();
				System.err.println("No.of mails in the list" + count);

				for (int i = 0; i < count; i++) {
					String mailFromTxt = emailSubjectList.get(i).getText();

					if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) {
						emailSubjectList.get(i).click();
						break;
					}

				}

				Thread.sleep(3000);

				Thread.sleep(3000);

				getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
				mailAttachmentDownloadBtn1.click();

				Thread.sleep(3000);

				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_J);
				robot.keyRelease(KeyEvent.VK_J);
				robot.keyRelease(KeyEvent.VK_CONTROL);

				Thread.sleep(5000);

				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);
				robot.keyPress(KeyEvent.VK_TAB);
				robot.keyRelease(KeyEvent.VK_TAB);

				robot.keyPress(KeyEvent.VK_ENTER);
				robot.keyRelease(KeyEvent.VK_ENTER);

				Thread.sleep(3000);

				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_S);
				robot.keyRelease(KeyEvent.VK_S);
				robot.keyRelease(KeyEvent.VK_CONTROL);

				Thread.sleep(2000);

				Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\POTriggersEmailHome.exe");

				Thread.sleep(4000);

				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_J);
				robot.keyRelease(KeyEvent.VK_J);
				robot.keyRelease(KeyEvent.VK_CONTROL);

				Thread.sleep(2000);

				ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

				int actOpenWindowsCount = getDriver().getWindowHandles().size();
				int expOpenWindowsCount = 4;

				System.out.println(
						"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

				getDriver().switchTo().window(newTabs.get(3)).close();
				Thread.sleep(1000);
				getDriver().switchTo().window(newTabs.get(2)).close();
				Thread.sleep(1000);
				getDriver().switchTo().window(newTabs.get(1));
				Thread.sleep(1000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
				mailDeleteButton.click();

				Thread.sleep(2000);

				Thread.sleep(2000);

				/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailUserBtn));
				gmailUserBtn.click();
*/
				ClickUsingJs(gmailUserBtn);
				Thread.sleep(3000);
				getDriver().switchTo().frame(SignOutFrame);
				Thread.sleep(1000);
				getAction().moveToElement(gmailSignOutBtn).click().build().perform();
				Thread.sleep(6000);
				getDriver().get(
						"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
				removeAccountBtn.click();

				Thread.sleep(1000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
				removeDeleteBtn.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
				yesRemoveBtn.click();

				Thread.sleep(2000);

				getDriver().switchTo().window(newTabs.get(1)).close();
				Thread.sleep(2000);
				getDriver().switchTo().window(newTabs.get(0));
				Thread.sleep(2000);

				String actAccount1PDF = getBaseDir() + "\\autoIt\\ExportFiles\\POTriggersEmailHome.pdf";
				String expAccount1PDF = getBaseDir() + "\\autoIt\\ImportFiles\\POTriggersEmailHome.pdf";

				PDFUtil pdfutil = new PDFUtil();

				boolean result = pdfutil.compare(actAccount1PDF, expAccount1PDF);
				Calendar cal = Calendar.getInstance();
				DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
				String currentDate = df.format(cal.getTime());

				String actAccount1Data = pdfutil.getText(actAccount1PDF);
				String expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("26/12/2022", currentDate);

				System.err.println(actAccount1Data);
				System.err.println(expAccount1Data);

				if (actAccount1Data.equalsIgnoreCase(expAccount1Data)) {

					System.err.println("PDF file is as Expected");
					return true;
				} else {

					System.err.println("PDF file is not as Expected");
					return false;
				}

			} else {

				System.err.println("Subject : " + emailRow1Subject.getText());

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsChckbx));
				selectAllMailsChckbx.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsDltBtn));
				selectAllMailsDltBtn.click();

				Thread.sleep(4000);

				Thread.sleep(2000);

			/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gmailUserBtn));
				gmailUserBtn.click();*/
				
				ClickUsingJs(gmailUserBtn);

				Thread.sleep(3000);
				getDriver().switchTo().frame(SignOutFrame);
				Thread.sleep(1000);
				getAction().moveToElement(gmailSignOutBtn).click().build().perform();
				Thread.sleep(6000);
				
				getDriver().get(
						"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
				removeAccountBtn.click();

				Thread.sleep(1000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
				removeDeleteBtn.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
				yesRemoveBtn.click();

				Thread.sleep(2000);
				return false;
			}

		}

	
	
	@FindBy(xpath="(//a[text()='Approve'])[4]")
	private static WebElement  approveThroughMail;

	@FindBy(xpath="//a[text()='Approve']")
	private static List<WebElement>  approveMailList;


	@FindBy(xpath="//a[text()='Approve']")
	private static WebElement  approveMail;

	@FindBy(xpath="/html/body/h1")
	private static WebElement  approveMessage;
	
	@FindBy(xpath="//a[text()='Inbox']")
	private static WebElement  inbox;
	
	public static boolean isElementPresent(By locatorKey) {
	    try {
	        getDriver().findElement(locatorKey);
	        return true;
	    } catch (org.openqa.selenium.NoSuchElementException e) {
	        return false;
	    }
	}


	
	
	
	public boolean checkAuthorizingVoucherFromEmail() throws InterruptedException, AWTException
	{
	

		File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\ApproveTheDocumentFromEmail.pdf");

		if (Efile1.exists()) {
			Efile1.delete();
		}

		Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);

		Thread.sleep(4000);

		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

		System.out.println("openTabs" + openTabs);

		getDriver().switchTo().window(openTabs.get(0));

		getDriver().switchTo().window(openTabs.get(1));

		getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

		Thread.sleep(6000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
		userNameTxt.click();

		userNameTxt.sendKeys("emailvalidationtwo@gmail.com");

		Thread.sleep(2000);

		System.out.println("User enter text");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
		NextBtn.click();

		Thread.sleep(3500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
		PasswordTxt.click();

		PasswordTxt.sendKeys("validationtwo");

		Thread.sleep(2000);

		NextBtn.click();

		Thread.sleep(8000);

	
	

			if (emailBody2.isDisplayed()) {

				int count = emailSubjectList.size();
				System.err.println("No.of mails in the list" + count);

				for (int i = 0; i < count; i++) {
					String mailFromTxt = emailSubjectList.get(i).getText();

					System.out.println("mailFromTxt :" + mailFromTxt);

					if (mailFromTxt.equalsIgnoreCase("Authorization")) {
						emailSubjectList.get(i).click();
						break;
					}
				}
			} 
			
			Thread.sleep(3000);
			
			click(approveMail);
				
				
				Thread.sleep(18000);
	
				String parent=getDriver().getWindowHandle();

				Set<String>s=getDriver().getWindowHandles();
				String child_window="";
				// Now iterate using Iterator
				Iterator<String> I1= s.iterator();

				while(I1.hasNext())
				{

				 child_window=I1.next();


				if(!parent.equals(child_window))
				{
					getDriver().switchTo().window(child_window);

				System.out.println(getDriver().switchTo().window(child_window).getTitle());

				//getDriver().close();
				}
				}	
				
				
				
				
				String expMesaage  ="You have approved the requested document";
				String actMessage="";
				
				while(1>0)
				{
					
					boolean result= isElementPresent(By.xpath("/html/body/h1"));
					System.out.println(result);
				if(result)
				{
					actMessage  =approveMessage.getText();
					break;
				}
				
				else
				{
					getDriver().switchTo().window(child_window).close();
					getDriver().switchTo().window(parent);
					click(approveMail);
					
					Thread.sleep(10000);
					
					String parent1=getDriver().getWindowHandle();
					while(I1.hasNext())
					{

					 child_window=I1.next();


					if(!parent1.equals(child_window))
					{
						getDriver().switchTo().window(child_window);

					System.out.println(getDriver().switchTo().window(child_window).getTitle());

					//getDriver().close();
					}
					}	
					
					actMessage  =approveMessage.getText();
				}
				}
				
				Thread.sleep(1500);
				System.out.println("actMesaage   :"+actMessage);
				System.out.println("expMesaage   :"+expMesaage);
				
				Thread.sleep(3000);
				
				

				Thread.sleep(2000);

				ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

				int actOpenWindowsCount = getDriver().getWindowHandles().size();
				int expOpenWindowsCount = 3;

				System.out.println(
						"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);
			
			
				Thread.sleep(1000);
				getDriver().switchTo().window(newTabs.get(2)).close();
				Thread.sleep(2000);
				
				getDriver().switchTo().window(parent);
				
				Thread.sleep(4000);
				
				
				
				click(inbox);
				
				

			
				Thread.sleep(3000);

				ClickUsingJs(gmailUserBtn);
				Thread.sleep(1000);

				Thread.sleep(2000);
				getDriver().switchTo().frame(SignOutFrame);
				Thread.sleep(1000);
				getAction().moveToElement(gmailSignOutBtn).click().build().perform();
				
				Thread.sleep(1200);
                
                getDriver().switchTo().defaultContent();
                System.out.println("switch out");
                
                Thread.sleep(1200);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
				removeAccountBtn.click();

				Thread.sleep(1000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
				removeDeleteBtn.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
				yesRemoveBtn.click();

				Thread.sleep(1000);

				getDriver().switchTo().window(newTabs.get(1)).close();
				Thread.sleep(2000);
				getDriver().switchTo().window(newTabs.get(0));
				Thread.sleep(2000);
			

				if (actMessage.equalsIgnoreCase(expMesaage))
						{

				
					return true;
				} else {

					
					return false;
				}

			}

			
	


	public boolean checkEmailAfterAuthorizationofPurchaseOrderVoucherThroughEmail() throws InterruptedException, AWTException, IOException
	{
		

	File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\POAuthorizationFromEmail.pdf");
		
		if(Efile1.exists())
		{
			Efile1.delete();
			System.out.println("File deleted");
		}
		
		
		Robot robot = new Robot();   
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);
		
		Thread.sleep(4000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

	 	System.out.println("openTabs"+openTabs);

	 	getDriver().switchTo().window(openTabs.get(0));
	 	Thread.sleep(2000);
	 	getDriver().switchTo().window(openTabs.get(1));
		

	 	getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");
	 	
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
	 	userNameTxt.click();
	 	
	 	userNameTxt.sendKeys("emailvalidationone@gmail.com");
	 	
	 	Thread.sleep(2000);
	 	
	 	System.out.println("User enter text");
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
	 	NextBtn.click();
	 	
	 	Thread.sleep(2000);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
	 	PasswordTxt.click();
	 	
	 	PasswordTxt.sendKeys("validationone");
	 	
	 	Thread.sleep(2000);
	 	
	 	NextBtn.click();
	 	
	 	Thread.sleep(15000);
	 	
	 boolean result=false;
	 	 	
	 	
	 		int count = emailSubjectList.size();
	 		System.err.println("No.of mails in the list"+count);
			 
	 		
	 		System.out.println(emailBody1.getText().isEmpty());
	 		
	 		if (emailBody2.getText().isEmpty()==false) 
		 	 	{
	 		
		 	for (int i = 0; i < count; i++) 
	 	{
	 		String mailFromTxt = emailSubjectList.get(i).getText();
	 		System.out.println(mailFromTxt);
	 	 	
	 		if (mailFromTxt.equalsIgnoreCase("Document no 4 has been approved")) 
	 		{
	 			emailSubjectList.get(i).click();     
	 			break;
	 		}
	 	}
		 	
	 			 	
		 	Thread.sleep(6000);
		
	
		
		if(refreshBtn.isDisplayed())
		{
			
			System.out.println("Subject is not matching");
		 	
		}
		
	 		
	 	
		else
		{
			
			System.out.println("Subject is  matching");
			getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
		mailAttachmentDownloadBtn1.click();
		
		Thread.sleep(3000);
		
		 
		robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_CONTROL);


	Thread.sleep(5000);

	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);
	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);

	robot.keyPress(KeyEvent.VK_ENTER);
	robot.keyRelease(KeyEvent.VK_ENTER);

	Thread.sleep(3000);

	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_CONTROL);
		
	Thread.sleep(2000);

			
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\POAuthorizationFromEmail.exe");
			
			Thread.sleep(4000);
			
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(2000);
			
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
			int actOpenWindowsCount = getDriver().getWindowHandles().size();
			int expOpenWindowsCount = 4;
			
			System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
			
			getDriver().switchTo().window(newTabs.get(3)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(2)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(1));
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 	mailDeleteButton.click();
		 	
		 	Thread.sleep(2000);
		 	

		 	Thread.sleep(2000);
		 	
		 	String actAccount1PDF = getBaseDir()+"\\autoIt\\ExportFiles\\POAuthorizationFromEmail.pdf";
			String expAccount1PDF = getBaseDir()+"\\autoIt\\ImportFiles\\POAuthorizationFromEmail.pdf";
			
		
		 
			PDFUtil pdfutil = new PDFUtil();
		
			boolean result1 = pdfutil.compare(actAccount1PDF, expAccount1PDF);
			Calendar cal=Calendar.getInstance();
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = df.format(cal.getTime());
			
			String  actAccount1Data = pdfutil.getText(actAccount1PDF);
			String  expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("02/05/2024", currentDate);
					
			System.err.println(actAccount1Data);
			System.err.println(expAccount1Data);
			
			if(actAccount1Data.equalsIgnoreCase(expAccount1Data))
			{
				result=true;
			}
		 	
		}
			
		 	 	}	
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());	
			ClickUsingJs(gmailUserBtn);

			Thread.sleep(3000);
			getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(2000);
			
			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
			
			getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
			
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 	removeAccountBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 	removeDeleteBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 	yesRemoveBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
			getDriver().switchTo().window(newTabs.get(1)).close();
		 	Thread.sleep(2000);
		 	
		 	
		 	getDriver().switchTo().window(newTabs.get(0));
		 	Thread.sleep(2000);
		 
			
			
			
		 	 		
		 	if (result && count!=0) 
			{
			
				System.err.println("PDF file is as Expected");
				return true;
			}
			else
			{
									
				System.err.println("PDF file is not as Expected");
				return false;
			}
			
			
		
		
	}
		
		
		 public boolean checkSavingMultipleRowsPurchaseOrderVouchertoAuthorizeFromHomePage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		 {
			 
			 	getDriver().navigate().refresh();
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
				invTransactionsMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
				inventoryTransactionsPurchasesMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
				purchasesOrdersVoucher.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
				newBtn.click();

				Thread.sleep(3000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
				MRpurchaseAccountTxt.click();
				MRpurchaseAccountTxt.sendKeys("STD RATE COGS ACC INV");
				Thread.sleep(3000);
				MRpurchaseAccountTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
				vendorAccountTxt.click();
				vendorAccountTxt.sendKeys("Vendor B");
				Thread.sleep(3000);
				vendorAccountTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValuetxt));
				departmentValuetxt.click();
				departmentValuetxt.sendKeys("INDIA");
				Thread.sleep(3000);
				departmentValuetxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				select1stRow_1stColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				enter_ItemTxt.click();
				enter_ItemTxt.sendKeys(Keys.SPACE);

				Thread.sleep(2000);
				
				enter_ItemTxt.sendKeys("STD RATE COGS ITEM");
				Thread.sleep(1500);

				/*int pvvGridItemListCount = pvvGridItemList.size();
				for (int i = 0; i < pvvGridItemListCount; i++) {
					String Item = pvvGridItemList.get(i).getText();
					if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
						pvvGridItemList.get(i).click();
						break;
					}
				}*/
				enter_ItemTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
				select1stRow_3rdColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
				enter_Quantity.sendKeys("11");

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
				select1stRow_4thColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
				enter_Rate.sendKeys("10.56");
				enter_Rate.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
				select2ndRow_1stColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				enter_ItemTxt.click();

				enter_ItemTxt.sendKeys(Keys.END);
				enter_ItemTxt.sendKeys(Keys.SHIFT, Keys.HOME);
				enter_ItemTxt.sendKeys("BR COGS ITEM");

				Thread.sleep(2000);

				enter_ItemTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
				select2ndRow_3rdColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
				enter_Quantity.sendKeys("12");

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
				select2ndRow_4thColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
				enter_Rate.sendKeys("25.89");
				enter_Rate.sendKeys(Keys.TAB);

				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
				select3rdRow_1stColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				enter_ItemTxt.click();

				enter_ItemTxt.sendKeys(Keys.END);
				enter_ItemTxt.sendKeys(Keys.SHIFT, Keys.HOME);
				enter_ItemTxt.sendKeys("FIFO COGS ITEM");

				Thread.sleep(2000);

				enter_ItemTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_3rdColumn));
				select3rdRow_3rdColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
				enter_Quantity.sendKeys("18");

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_4thColumn));
				select3rdRow_4thColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
				enter_Rate.sendKeys("75.21");
				enter_Rate.sendKeys(Keys.TAB);

				Thread.sleep(1000);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
				select4thRow_1stColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
				enter_ItemTxt.click();

				enter_ItemTxt.sendKeys(Keys.END);
				enter_ItemTxt.sendKeys(Keys.SHIFT, Keys.HOME);
				enter_ItemTxt.sendKeys("WA COGS ITEM");

				Thread.sleep(2000);

				enter_ItemTxt.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_3rdColumn));
				select4thRow_3rdColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
				enter_Quantity.sendKeys("5");

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_4thColumn));
				select4thRow_4thColumn.click();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
				enter_Rate.sendKeys("11");
				enter_Rate.sendKeys(Keys.TAB);
				
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
				voucherSaveBtn.click();

				String expMessage1 = "Voucher saved successfully";

				String actMessage = checkValidationMessage(expMessage1);
				// String expMessage2 = ": 3";

				System.out.println("Actual Message    : " + actMessage);
				System.out.println("Expected Message  : " + expMessage1);

				if (actMessage.startsWith(expMessage1)/* && actMessage.endsWith(expMessage2) */ )

				{
					return true;
				} else {
					return false;
				}

			
		 }
		 
		 @FindBy(xpath="//span[@class='icon-new d-none d-md-block hiconright2']")
		 public static WebElement PO_EntryNewBtn;
		 public boolean checkSavingMultipleVouchersinPurchaseOrderVouchertoAuthorizeFromHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		 {
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			 documentNumberTxt.click();
			 documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			 documentNumberTxt.sendKeys("5");
			 Thread.sleep(1500);
			 documentNumberTxt.sendKeys(Keys.TAB);
			 Thread.sleep(2000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
			 toggleBtn.click();
			 Thread.sleep(2000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(copyToClipBoardOption));
			 copyToClipBoardOption.click();
			 Thread.sleep(1500);
			 
			 boolean flag1=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag2=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag3=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag4=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag5=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag6=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 
			 boolean flag7=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag8=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag9=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag10=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag11=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag12=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag13=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 
			 boolean flag14=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag15=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag16=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag17=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag18=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			 boolean flag19=checkSavingPOVoucherusingCopyandPaste();
			 Thread.sleep(2000);
			 
			
			 
			 
			 if(flag1&&flag2&&flag3&&flag4&& flag5 && flag6 && flag7&&flag8&&flag9&&flag10&& flag11)
			 {
				 return true;
			 }
			 else
			 {
				 return false;
			 }
		 }
		 
		 @FindBy(xpath="//span[@class='icon-menudots hiconright2']")
		 public static WebElement toggleBtn;
		 
		 public boolean checkSavingPOVoucherusingCopyandPaste() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		 {
			 Thread.sleep(1500);
				 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PO_EntryNewBtn));
			 PO_EntryNewBtn.click();
			 Thread.sleep(1500);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
			 toggleBtn.click();
			 Thread.sleep(2000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pasteFromClipBoardOption));
			 pasteFromClipBoardOption.click();
			 Thread.sleep(2000);
			 checkValidationMessage("");
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
			 voucherSaveBtn.click();
			 Thread.sleep(2000);
				String expMessage = "Voucher saved successfully";

				String actMessage = checkValidationMessage(expMessage);
				// String expMessage2 = ": 3";

				System.out.println("Actual Message    : " + actMessage);
				System.out.println("Expected Message  : " + expMessage);

				if (actMessage.startsWith(expMessage) )

				{
					return true;
				} else {
					return false;
				}
		 }
		 

			@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr//td[5]")
			public static List<WebElement> voucherNoList;
			
			@FindBy(xpath = "//*[@id='frstPage']")
			public static WebElement voucherFirstPageBtn;
			

			@FindBy(xpath = "//*[@id='id_nextPage']")
			public static WebElement voucherNxtPageBtn;
			
			
			
			@FindBy(xpath = "//*[@data-tooltip='Select']//span")
			public static WebElement emailSelectAllCheckBox;
			
			
			  @FindBy(xpath="(//div[@data-tooltip='Delete']//div)[1]")
			   public  static WebElement mailSelectAllDeleteButton;
			
			
			
			public int checkEmailsAfterAuthorizationofVouchers() throws InterruptedException, AWTException
			{
				
				
				Robot robot = new Robot();
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.keyRelease(KeyEvent.VK_T);

				Thread.sleep(4000);

				ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

				System.out.println("openTabs" + openTabs);

				getDriver().switchTo().window(openTabs.get(0));

				getDriver().switchTo().window(openTabs.get(1));

				getDriver().get(
						"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
				userNameTxt.click();

				userNameTxt.sendKeys("emailvalidationone@gmail.com");

				Thread.sleep(2000);

				System.out.println("User enter text");

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
				NextBtn.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
				PasswordTxt.click();

				PasswordTxt.sendKeys("validationone");

				Thread.sleep(2000);

				NextBtn.click();

				Thread.sleep(12000);

				int count1 = row1mailFromList.size();
				
				
				int actMailsCount=0;

				if (emailBody2.getText().isEmpty() == false) {
					int count = emailSubjectList.size();
					System.err.println("No.of mails in the list" + count);

					for (int i = 0; i < count; i++) {
						String mailFromTxt = emailSubjectList.get(i).getText();
						Thread.sleep(2000);
						if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) {
							Thread.sleep(2000);
							
							actMailsCount=actMailsCount+1;
						}

					}
				}
					Thread.sleep(4000);
					
					
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSelectAllCheckBox));
		emailSelectAllCheckBox.click();
					
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailSelectAllDeleteButton));
					
		mailSelectAllDeleteButton.click();	
		
		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

		int actOpenWindowsCount = getDriver().getWindowHandles().size();
		int expOpenWindowsCount = 2;

		System.out.println(
				"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

		
		ClickUsingJs(gmailUserBtn);

		Thread.sleep(3000);
		getDriver().switchTo().frame(SignOutFrame);
		Thread.sleep(1000);
		getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		Thread.sleep(6000);
		
		getDriver().get(
				"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		removeAccountBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		removeDeleteBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		yesRemoveBtn.click();

		Thread.sleep(2000);

		getDriver().switchTo().window(newTabs.get(1)).close();
		Thread.sleep(2000);
		getDriver().switchTo().window(newTabs.get(0));
		Thread.sleep(2000);
		
		
		return actMailsCount;
			}
			
			
			
			public void checkDeleteAllEmailsBeforeAuthorizationofVouchers() throws InterruptedException, AWTException
			{
				Thread.sleep(4000);
				
				Robot robot = new Robot();
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.keyRelease(KeyEvent.VK_T);

				Thread.sleep(4000);

				ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

				System.out.println("openTabs" + openTabs);

				getDriver().switchTo().window(openTabs.get(0));

				getDriver().switchTo().window(openTabs.get(1));

				getDriver().get(
						"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
				userNameTxt.click();

				userNameTxt.sendKeys("emailvalidationone@gmail.com");

				Thread.sleep(2000);

				System.out.println("User enter text");

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
				NextBtn.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
				PasswordTxt.click();

				PasswordTxt.sendKeys("validationone");

				Thread.sleep(2000);

				NextBtn.click();

				Thread.sleep(12000);

				int count1 = row1mailFromList.size();
				
				
				int actMailsCount=0;

				if (emailBody2.getText().isEmpty() == false)
				{
					
					
					
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSelectAllCheckBox));
		emailSelectAllCheckBox.click();
					
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailSelectAllDeleteButton));
					
		mailSelectAllDeleteButton.click();	
		
				}
		
		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

		int actOpenWindowsCount = getDriver().getWindowHandles().size();
		int expOpenWindowsCount = 2;

		System.out.println(
				"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

		
		ClickUsingJs(gmailUserBtn);

		Thread.sleep(3000);
		getDriver().switchTo().frame(SignOutFrame);
		Thread.sleep(1000);
		getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		Thread.sleep(6000);
		
		getDriver().get(
				"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		removeAccountBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		removeDeleteBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		yesRemoveBtn.click();

		Thread.sleep(2000);

		getDriver().switchTo().window(newTabs.get(1)).close();
		Thread.sleep(2000);
		getDriver().switchTo().window(newTabs.get(0));
		Thread.sleep(2000);
		
			
				
			}
			
			
		 public boolean checkAuthorizingTwoVouchersinPOVoucherfromHomePage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
		 {
				Thread.sleep(2000);
				
				
				checkDeleteAllEmailsBeforeAuthorizationofVouchers();
				Thread.sleep(8000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
				invTransactionsMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
				inventoryTransactionsPurchasesMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
				purchasesOrdersVoucher.click();

				Thread.sleep(3000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

				
				boolean result= voucherNxtPageBtn.isDisplayed();
				System.out.println(result);
						
				if(result)
				{
					((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,300)","");
					Thread.sleep(2000);
				
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherNxtPageBtn));
					voucherNxtPageBtn.click();
					Thread.sleep(6000);
				}
				
				int vcount = authorizationStatusList.size();
				System.out.println(vcount);
				int count = 0;
				
				
				while(count<2)
				{
					for (int i = 0; i < vcount; i++) {
						String VoucherAuth = authorizationStatusList.get(i).getText();
						String VoucherNo=voucherNoList.get(i).getText();
						
						
						if(VoucherAuth.equalsIgnoreCase("Pending"))
						{
							switch(VoucherNo)
							{
							case "5":
								
								grid_CheckBoxList.get(i).click();
								count++;
								break;
								
							case "6":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
								
							}
						}
					}
				}
				
				//getAction().moveToElement(authorizeBtn).build().perform();
				Thread.sleep(3000);
				

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
				authorizeBtn.click();
				Thread.sleep(2000);

				String expMessage="Voucher Approoved Successfully";
				String actMessage=checkValidationMessage(expMessage);
				if (actMessage.equalsIgnoreCase(expMessage)) {
					/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
					voucherhomeCloseBtn.click();*/

					return true;
				}

				else {
					return false;
				}
			
		 }
		 
		 
		 @FindBy(xpath="(//i[@class='icon-close hiconright2'])[1]")
			public static WebElement voucherhomeCloseBtn;
		 
			@FindBy(xpath="//*[@class='f gW']")
			public static List<WebElement> noofAttachements;
			
			
			public boolean checkEmailAttachementsAfterAuthorizingTwoVouchers() throws InterruptedException, AWTException, IOException
			{
				//int actAttachements=checkEmailAttachementsAfterAuthorizationofPOVouchers();
				
				int actAttachements=checkEmailsAfterAuthorizationofVouchers();
				int expAttachements=2;
				
				
				System.out.println("Actual Attachements=	"	+actAttachements);
				System.out.println("Expected Attachements	"	+expAttachements);
				
				if(actAttachements==expAttachements)
				{
					return true;
				}
				else
				{
					return false;
				}
			}
			
			
			@FindBy(xpath="(//div[@class='vh'])[1]//div[@class='bBe']")
			public static WebElement notificationCloseBtn;

		 
		 
		 public int checkEmailAttachementsAfterAuthorizationofPOVouchers() throws InterruptedException, AWTException, IOException
		 {
			 


				Thread.sleep(2000);
				File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\POAuthorizeFromHomePageTriggersEmail.pdf");

				if (Efile1.exists()) {
					Efile1.delete();
				}

				Robot robot = new Robot();
				robot.keyPress(KeyEvent.VK_CONTROL);
				robot.keyPress(KeyEvent.VK_T);
				robot.keyRelease(KeyEvent.VK_CONTROL);
				robot.keyRelease(KeyEvent.VK_T);

				Thread.sleep(4000);

				ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

				System.out.println("openTabs" + openTabs);

				getDriver().switchTo().window(openTabs.get(0));

				getDriver().switchTo().window(openTabs.get(1));

				getDriver().get(
						"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
				userNameTxt.click();

				userNameTxt.sendKeys("emailvalidationone@gmail.com");

				Thread.sleep(2000);

				System.out.println("User enter text");

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
				NextBtn.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
				PasswordTxt.click();

				PasswordTxt.sendKeys("validationone");

				Thread.sleep(2000);

				NextBtn.click();

				Thread.sleep(12000);

				int count1 = row1mailFromList.size();
				//System.out.println("No.of mails in the list" + count1);
				
				//String expEmailBody="Automation Company : 08/10/2020 5:43 PM < > STD RATE COGS ITEM 2.0000000000 10.0000000000 20.0000000 PurOrd:2 INDIA";
				//String actEmailBody=emailBody.getText();
				
				//System.err.println("Email Body *" + emailBody.getText() + "*");
				//System.err.println("Expected Email Body *" + expEmailBody);
				
				int actAttachments=0;

				if (emailBody2.getText().isEmpty() == false) {
					int count = emailSubjectList.size();
					System.err.println("No.of mails in the list" + count);

					for (int i = 0; i < count; i++) {
						String mailFromTxt = emailSubjectList.get(i).getText();
						Thread.sleep(2000);
						if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) {
							Thread.sleep(2000);
							
							emailSubjectList.get(i).click();
							break;
						}

					}

					Thread.sleep(3000);

					Thread.sleep(3000);
					
					/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(notificationCloseBtn));
					notificationCloseBtn.click();
					Thread.sleep(2000);*/
					
					
					getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
					
					actAttachments=noofAttachements.size();
					
					
					System.err.println("No of Attachemants" +noofAttachements.size());
					

					
					ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

					int actOpenWindowsCount = getDriver().getWindowHandles().size();
					int expOpenWindowsCount = 2;

					System.out.println(
							"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
					mailDeleteButton.click();

					Thread.sleep(2000);

					Thread.sleep(2000);

					/*
					 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
					 * gmailUserBtn)); gmailUserBtn.click();
					 */
					
					ClickUsingJs(gmailUserBtn);

					Thread.sleep(3000);
					getDriver().switchTo().frame(SignOutFrame);
					Thread.sleep(1000);
					getAction().moveToElement(gmailSignOutBtn).click().build().perform();
					Thread.sleep(6000);
					
					getDriver().get(
							"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

					Thread.sleep(2000);
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
					removeAccountBtn.click();

					Thread.sleep(1000);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
					removeDeleteBtn.click();

					Thread.sleep(2000);

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
					yesRemoveBtn.click();

					Thread.sleep(2000);

					getDriver().switchTo().window(newTabs.get(1)).close();
					Thread.sleep(2000);
					getDriver().switchTo().window(newTabs.get(0));
					Thread.sleep(2000);

					
				}
				
				return actAttachments;


		 }
		 
		 
		 
		 public boolean checkAuthorizingThreeVouchersFromHomePage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		 {

				Thread.sleep(5000);

				/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
				invTransactionsMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
				inventoryTransactionsPurchasesMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
				purchasesOrdersVoucher.click();

				Thread.sleep(3000);*/
				
				/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherFirstPageBtn));
				voucherFirstPageBtn.click();
				Thread.sleep(2000);*/

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

				int vcount = authorizationStatusList.size();
				int count = 0;
				
				((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,300)","");
				Thread.sleep(2000);
						
				while(count<3)
				{
					for (int i = 0; i < vcount; i++) {
						String VoucherAuth = authorizationStatusList.get(i).getText();
						String VoucherNo=voucherNoList.get(i).getText();
						if(VoucherAuth.equalsIgnoreCase("Pending"))
						{
							switch(VoucherNo)
							{
							case "7":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							case "8":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							case "9":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							}
						}
					}
				}
				
				
				getAction().moveToElement(authorizeBtn).build().perform();
				Thread.sleep(2000);
				


				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
				authorizeBtn.click();
				Thread.sleep(2000);

				String expMessage="Voucher Approoved Successfully";
				String actMessage=checkValidationMessage(expMessage);
				if (actMessage.equalsIgnoreCase(expMessage)) {
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
					voucherhomeCloseBtn.click();

					return true;
				}

				else {
					return false;
				}
			

		 }
		 
		 public boolean checkEmailAttachementsAfterAuthorizingThreeVouchers() throws InterruptedException, AWTException, IOException
			{
				//int actAttachements=checkEmailAttachementsAfterAuthorizationofPOVouchers();
				int actAttachements=checkEmailsAfterAuthorizationofVouchers();
				int expAttachements=3;
				
				
				System.out.println("Actual Attachements=	"	+actAttachements);
				System.out.println("Expected Attachements	"	+expAttachements);
				
				if(actAttachements==expAttachements)
				{
					return true;
				}
				else
				{
					return false;
				}
			}
		 
		 
		 
		 
		 public boolean checkAuthorizingFiveVouchersinPOVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		 {
			 Thread.sleep(7000);
			 getDriver().navigate().refresh();
				Thread.sleep(4000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
				invTransactionsMenu.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
				inventoryTransactionsPurchasesMenu.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
				purchasesOrdersVoucher.click();

				Thread.sleep(4000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

				int vcount = authorizationStatusList.size();
				int count = 0;
				((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,300)","");
				Thread.sleep(2000);
				while(count<5)
				{
					for (int i = 0; i < vcount; i++) {
						String VoucherAuth = authorizationStatusList.get(i).getText();
						String VoucherNo=voucherNoList.get(i).getText();
						if(VoucherAuth.equalsIgnoreCase("Pending"))
						{
							switch(VoucherNo)
							{
							case "10":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							case "11":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							case "12":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							
							case "13":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							case "14":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							}
						}
					}
				}
				
				
				getAction().moveToElement(authorizeBtn).build().perform();
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
				authorizeBtn.click();
				Thread.sleep(2000);

				String expMessage="Voucher Approoved Successfully";
				String actMessage=checkValidationMessage(expMessage);
				if (actMessage.equalsIgnoreCase(expMessage)) {
					

					return true;
				}

				else {
					return false;
				}
			


		 }
		 
		 
		 
		 public boolean checkEmailAttachementsAfterAuthorizingFiveVouchers() throws InterruptedException, AWTException, IOException
			{
			//	int actAttachements=checkEmailAttachementsAfterAuthorizationofPOVouchers();
			 
				int actAttachements=checkEmailsAfterAuthorizationofVouchers();
				int expAttachements=5;
				
				
				System.out.println("Actual Attachements=	"	+actAttachements);
				System.out.println("Expected Attachements	"	+expAttachements);
				
				if(actAttachements==expAttachements)
				{
					return true;
				}
				else
				{
					return false;
				}
			}
		 
		 public boolean checkAuthorizingTenVouchersInPOVoucherFromHomePage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		 {
				Thread.sleep(2000);

				/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
				invTransactionsMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
				inventoryTransactionsPurchasesMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
				purchasesOrdersVoucher.click();

				Thread.sleep(3000);
		*/
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

				int vcount = authorizationStatusList.size();
				int count = 0;
						
				while(count<10)
				{
					for (int i = 0; i < vcount; i++) {
						String VoucherAuth = authorizationStatusList.get(i).getText();
						String VoucherNo=voucherNoList.get(i).getText();
						if(VoucherAuth.equalsIgnoreCase("Pending"))
						{
							switch(VoucherNo)
							{
							case "15":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							case "16":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							case "17":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
								
							case "18":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							case "19":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
								
							case "20":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
								
							case "21":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
								
							case "22":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
								
							case "23":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
								
							case "24":
								grid_CheckBoxList.get(i).click();
								count++;
								break;
							}
						}
					}
				}
				
				
				

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
				authorizeBtn.click();
				Thread.sleep(2000);

				String expMessage="Voucher Approoved Successfully";
				String actMessage=checkValidationMessage(expMessage);
				if (actMessage.equalsIgnoreCase(expMessage)) {
					

					return true;
				}

				else {
					return false;
				}
			



		 }
		 
		 
		 public boolean checkEmailAttachementsAfterAuthorizingTenVouchers() throws InterruptedException, AWTException, IOException
			{
				//int actAttachements=checkEmailAttachementsAfterAuthorizationofPOVouchers();
				int actAttachements=checkEmailsAfterAuthorizationofVouchers();
				int expAttachements=10;
				
				
				System.out.println("Actual Attachements=	"	+actAttachements);
				System.out.println("Expected Attachements	"	+expAttachements);
				
				if(actAttachements==expAttachements)
				{
					return true;
				}
				else
				{
					return false;
				}
			}
		 

		
		
		 
		 
		 public boolean checkEmailTemplateofLoginforPurchaseOrderVoucher()
					throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException {
				getDriver().navigate().refresh();
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
				System.out.println("Setting buton is enabled");
				serachMenuTextHomePage.click();
				serachMenuTextHomePage.sendKeys("Configure Transactions");

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
				searchMenuTextClick.click();

				Thread.sleep(2000);

			/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
				documentCustomization45Btn.click();
*/
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseOrdersBtn));
				purchaseOrdersBtn.click();

				Thread.sleep(4000);
				/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toogleExpandBtn));
				toogleExpandBtn.click();

				Thread.sleep(2000);*/
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
				triggersBtn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
				editLayoutTriggerName.click();
				editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
				Thread.sleep(1000);
				editLayoutTriggerName.sendKeys("Authorize Trigger");
				Thread.sleep(2000);
				editLayoutTriggerName.sendKeys(Keys.TAB);
				Thread.sleep(5000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailTxtBox));
				if (sendEmailTxtBox.isSelected() == false) {
					sendEmailTxtBox.click();
				}
				sendEmailTxtBox.sendKeys(Keys.HOME, Keys.SHIFT, Keys.END);
				sendEmailTxtBox.sendKeys("POTriggersEmailLayout");
				Thread.sleep(3000);
				sendEmailTxtBox.sendKeys(Keys.TAB);
				
				getAction().moveToElement(customerChkBox).build().perform();
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerChkBoxSelected));
				if (customerChkBox.isSelected() == true) {
					customerChkBoxSelected.click();
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginChkBoxSelected));
				if (loginChkBox.isSelected() == false) {
					loginChkBoxSelected.click();
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectUsersList));
				Select s1 = new Select(selectUsersList);
				s1.selectByVisibleText("TriggerUser");
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
				editLayoutTriggersSaveBtn.click();
				Thread.sleep(2500);

				String expSaveMessage = "Data saved successfully";
				String actSaveMessage = checkValidationMessage(expSaveMessage);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
				updateBtn.click();

				String expUpdateMessage = "Data saved successfully";
				String actUpdateMessage = checkValidationMessage(expUpdateMessage);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
				settings_closeBtn.click();

				if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
					return true;
				} else {
					return false;
				}

			}		 

		 
		 
		 public boolean checkVerifingDetailsOfSavedPurchaseVouchers4InHomepage()
					throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
				financialsMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
				financialsTransactionMenu.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesExpandBtn));
				purchasesExpandBtn.click();
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersBtn));
				purchaseVouchersBtn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

				Calendar cal = Calendar.getInstance();
				SimpleDateFormat todayDate = new SimpleDateFormat("dd/MM/yyyy");
				String currentDate = todayDate.format(cal.getTime());
				System.out.println("--------------Date Before Adding:: " + currentDate);

				int count = voucherHomeBodyLists1st.size();

				ArrayList<String> actVouchersList = new ArrayList<String>();

				for (int i = 0; i < count; i++) {
					String data = voucherHomeBodyLists1st.get(i).getText();
					System.out.println(data);
					if (data.isEmpty() == false && i != 9 && i != 10 && i!=3 ) {
						System.out.println(data);

						actVouchersList.add(data);
					}
				}

				ArrayList<String> expVouchersList = new ArrayList<String>();

				expVouchersList.add("1");
				//expVouchersList.add(currentDate);
				expVouchersList.add("26");
				expVouchersList.add("UserAllOptionsST");
				expVouchersList.add("UserAllOptionsST");
				expVouchersList.add(currentDate);
				expVouchersList.add(currentDate);
				expVouchersList.add("True");
				expVouchersList.add("Authorized");
				/*expVouchersList.add(" ");
				expVouchersList.add("Total");*/

				System.out.println(
						"**********************checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage*********************");

				System.out.println(actVouchersList);
				System.out.println(expVouchersList);

				if (actVouchersList.equals(expVouchersList)) {
					return true;
				} else {
					return false;
				}
			}
		 
		 
		 
		 public boolean checkEmailAfterEnablingLoginOptionAfterAuthorizationofPurchaseOrderVoucherfromHomePage() throws IOException, AWTException, InterruptedException
		 {
		 	Thread.sleep(7000);
		 	File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\POAuthorizeFromHomePageTriggersEmailLoginOption.pdf");

		 	if (Efile1.exists()) {
		 		Efile1.delete();
		 	}

		 	Robot robot = new Robot();
		 	robot.keyPress(KeyEvent.VK_CONTROL);
		 	robot.keyPress(KeyEvent.VK_T);
		 	robot.keyRelease(KeyEvent.VK_CONTROL);
		 	robot.keyRelease(KeyEvent.VK_T);

		 	Thread.sleep(4000);

		 	ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

		 	System.out.println("openTabs" + openTabs);

		 	getDriver().switchTo().window(openTabs.get(0));
		 	Thread.sleep(2000);
		 	getDriver().switchTo().window(openTabs.get(1));
		 	Thread.sleep(2000);
		 	getDriver().get(
		 			"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

		 	Thread.sleep(2000);
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
		 	userNameTxt.click();

		 	userNameTxt.sendKeys("emailvalidationone@gmail.com");

		 	Thread.sleep(2000);

		 	System.out.println("User enter text");

		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
		 	NextBtn.click();

		 	Thread.sleep(2000);

		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
		 	PasswordTxt.click();

		 	PasswordTxt.sendKeys("validationone");

		 	Thread.sleep(2000);

		 	NextBtn.click();

		 	Thread.sleep(10000);

		 	int count1 = row1mailFromList.size();
		 	/*System.out.println("No.of mails in the list" + count1);
		 	
		 	String expEmailBody="Automation Company : 08/10/2020 5:43 PM< >STD RATE COGS ITEM2.000000000010.000000000020.0000000PurOrd:24";
		 	String actEmailBody=emailBody.getText();
		 	
		 	System.err.println("Email Body *" + emailBody.getText() + "*");
		 	System.err.println("Expected Email Body *" + expEmailBody);*/

		 	if (emailBody2.getText().isEmpty() == false) {
		 		int count = emailSubjectList.size();
		 		System.err.println("No.of mails in the list" + count);

		 		for (int i = 0; i < count; i++) {
		 			String mailFromTxt = emailSubjectList.get(i).getText();
		 			Thread.sleep(2000);
		 			if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) {
		 				Thread.sleep(2000);
		 				
		 				emailSubjectList.get(i).click();
		 				break;
		 			}

		 		}

		 		Thread.sleep(3000);

		 		Thread.sleep(3000);
		 		
		 	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(notificationCloseBtn));
		 		notificationCloseBtn.click();
		 		Thread.sleep(2000);*/
		 		
		 		
		 		getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
		 		Thread.sleep(2000);


		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
		 		mailAttachmentDownloadBtn1.click();

		 		Thread.sleep(4000);

		 		robot.keyPress(KeyEvent.VK_CONTROL);
		 		robot.keyPress(KeyEvent.VK_J);
		 		robot.keyRelease(KeyEvent.VK_J);
		 		robot.keyRelease(KeyEvent.VK_CONTROL);

		 		Thread.sleep(5000);

		 		robot.keyPress(KeyEvent.VK_TAB);
		 		robot.keyRelease(KeyEvent.VK_TAB);
		 		robot.keyPress(KeyEvent.VK_TAB);
		 		robot.keyRelease(KeyEvent.VK_TAB);

		 		robot.keyPress(KeyEvent.VK_ENTER);
		 		robot.keyRelease(KeyEvent.VK_ENTER);

		 		Thread.sleep(3000);

		 		robot.keyPress(KeyEvent.VK_CONTROL);
		 		robot.keyPress(KeyEvent.VK_S);
		 		robot.keyRelease(KeyEvent.VK_S);
		 		robot.keyRelease(KeyEvent.VK_CONTROL);

		 		Thread.sleep(2000);

		 		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\POAuthorizeFromHomePageTriggersEmailLoginOption.exe");

		 		Thread.sleep(7000);

		 		robot.keyPress(KeyEvent.VK_CONTROL);
		 		robot.keyPress(KeyEvent.VK_J);
		 		robot.keyRelease(KeyEvent.VK_J);
		 		robot.keyRelease(KeyEvent.VK_CONTROL);

		 		Thread.sleep(2000);

		 		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

		 		int actOpenWindowsCount = getDriver().getWindowHandles().size();
		 		int expOpenWindowsCount = 4;

		 		System.out.println(
		 				"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

		 		getDriver().switchTo().window(newTabs.get(3)).close();
		 		Thread.sleep(1000);
		 		getDriver().switchTo().window(newTabs.get(2)).close();
		 		Thread.sleep(1000);
		 		getDriver().switchTo().window(newTabs.get(1));
		 		Thread.sleep(1000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 		mailDeleteButton.click();

		 		Thread.sleep(2000);

		 		Thread.sleep(2000);

		 		/*
		 		 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 		 * gmailUserBtn)); gmailUserBtn.click();
		 		 */
		 		
		 		ClickUsingJs(gmailUserBtn);

		 		Thread.sleep(3000);
		 		getDriver().switchTo().frame(SignOutFrame);
		 		Thread.sleep(1000);
		 		getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 		Thread.sleep(6000);
		 		
		 		getDriver().get(
						"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

				Thread.sleep(2000);
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 		removeAccountBtn.click();

		 		Thread.sleep(1000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 		removeDeleteBtn.click();

		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 		yesRemoveBtn.click();

		 		Thread.sleep(2000);

		 		getDriver().switchTo().window(newTabs.get(1)).close();
		 		Thread.sleep(2000);
		 		getDriver().switchTo().window(newTabs.get(0));
		 		Thread.sleep(2000);

		 		String actAccount1PDF = getBaseDir() + "\\autoIt\\ExportFiles\\POAuthorizeFromHomePageTriggersEmailLoginOption.pdf";
		 		String expAccount1PDF = getBaseDir() + "\\autoIt\\ImportFiles\\POAuthorizeFromHomePageTriggersEmailLoginOption.pdf";

		 		PDFUtil pdfutil = new PDFUtil();

		 		boolean result = pdfutil.compare(actAccount1PDF, expAccount1PDF);
		 		Calendar cal = Calendar.getInstance();
		 		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		 		String currentDate = df.format(cal.getTime());

		 		String actAccount1Data = pdfutil.getText(actAccount1PDF);
		 		String expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("17/12/2024", currentDate);

		 		System.err.println(actAccount1Data);
		 		System.err.println(expAccount1Data);

		 		if (actAccount1Data.equalsIgnoreCase(expAccount1Data) ) {

		 			System.err.println("PDF file is as Expected");
		 			return true;
		 		} else {

		 			System.err.println("PDF file is not as Expected");
		 			return false;
		 		}

		 	} else {

		 		System.err.println("Subject : " + emailRow1Subject.getText());

		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsChckbx));
		 		selectAllMailsChckbx.click();

		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsDltBtn));
		 		selectAllMailsDltBtn.click();

		 		Thread.sleep(4000);

		 		Thread.sleep(2000);

		 		/*
		 		 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 		 * gmailUserBtn)); gmailUserBtn.click();
		 		 */
		 		
		 		ClickUsingJs(gmailUserBtn);

		 		Thread.sleep(3000);
		 		getDriver().switchTo().frame(SignOutFrame);
		 		Thread.sleep(1000);
		 		getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 		Thread.sleep(6000);
		 		
		 		getDriver().get(
						"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

				Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 		removeAccountBtn.click();

		 		Thread.sleep(1000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 		removeDeleteBtn.click();

		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 		yesRemoveBtn.click();

		 		Thread.sleep(2000);
		 		return false;
		 	}
		 }
		 
		 
		 
		 @FindBy(xpath="(//*[contains(text(),'Miscellaneous')])[3]")		
			public static WebElement miscellaneousTab1;
		 
		 @FindBy(xpath="//select[@id='misc_PickEmailFrom']")
			public static WebElement pickEmailFromSelect;
		 
			@FindBy(xpath = "//input[@id='triggers_chksendTypeExtraField']")
			public static WebElement extraFieldChkBox;
			
			@FindBy(xpath = "//input[@id='triggers_chksendTypeExtraField']/following-sibling::span")
			public static WebElement extraFieldChkBoxSelected;

		 public boolean checkEnablingPickEmailFromNarrationinPOVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		 {
		 	 	Thread.sleep(2000);
		 	    getDriver().navigate().refresh();
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		 		System.out.println("Setting buton is enabled");
		 		serachMenuTextHomePage.click();
		 		serachMenuTextHomePage.sendKeys("Configure Transactions");

		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		 		searchMenuTextClick.click();

		 		Thread.sleep(2000);

		 		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
		 		documentCustomization45Btn.click();*/

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseOrdersBtn));
		 		purchaseOrdersBtn.click();

		 		Thread.sleep(3000);

		 		  
		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTab1));
		 				miscellaneousTab1.click();
		 		
		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickEmailFromSelect));
		 				pickEmailFromSelect.click();
		 				Select s=new Select(pickEmailFromSelect);
		 				s.selectByVisibleText("Narration");
		 				Thread.sleep(2000);
		 				
		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		 				updateBtn.click();
		 				
		 				String expMessage1 = "Data saved successfully";
		 				String actMessage1 = checkValidationMessage(expMessage1);
		 				
		 				/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toogleExpandBtn));
		 				toogleExpandBtn.click();
*/
		 				Thread.sleep(2000);
		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		 				triggersBtn.click();

		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		 				editLayoutTriggerName.click();
		 				editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
		 				Thread.sleep(1000);
		 				editLayoutTriggerName.sendKeys("Authorize Trigger");
		 				Thread.sleep(2000);
		 				editLayoutTriggerName.sendKeys(Keys.TAB);
		 				Thread.sleep(5000);

		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailTxtBox));
		 				if (sendEmailTxtBox.isSelected() == false) {
		 					sendEmailTxtBox.click();
		 				}
		 				sendEmailTxtBox.sendKeys(Keys.HOME, Keys.SHIFT, Keys.END);
		 				sendEmailTxtBox.sendKeys("POTriggersEmailLayout");
		 				Thread.sleep(3000);
		 				
		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerChkBoxSelected));
		 				if (customerChkBox.isSelected() == true) {
		 					customerChkBoxSelected.click();
		 				}

		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginChkBoxSelected));
		 				if (loginChkBox.isSelected() == true) {
		 					loginChkBoxSelected.click();
		 				}

		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFieldChkBoxSelected));
		 				if(extraFieldChkBox.isSelected()==false)
		 				{
		 					extraFieldChkBoxSelected.click();
		 				}
		 				
		 				Thread.sleep(3000);
		 				
		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		 				editLayoutTriggersSaveBtn.click();
		 				Thread.sleep(2500);

		 				String expSaveMessage = "Data saved successfully";
		 				String actSaveMessage = checkValidationMessage(expSaveMessage);

		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		 				updateBtn.click();

		 				String expUpdateMessage = "Data saved successfully";
		 				String actUpdateMessage = checkValidationMessage(expUpdateMessage);

		 				Thread.sleep(2000);
		 				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		 				settings_closeBtn.click();

		 				
		 				
		 				if ( actSaveMessage.equalsIgnoreCase(expSaveMessage) && actMessage1.equalsIgnoreCase(expMessage1))
		 				{
		 					return true;
		 				}
		 				else
		 				{
		 					return false;
		 				}
		 				
		 }

		  public boolean checkSavingPOVoucher1toSendEmailfromExtraField() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		  
		  {


		 		getDriver().navigate().refresh();
		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		 		inventoryMenu.click();

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
		 		invTransactionsMenu.click();

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		 		inventoryTransactionsPurchasesMenu.click();

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		 		purchasesOrdersVoucher.click();

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		 		newBtn.click();

		 		Thread.sleep(3000);

		 		 boolean result=checkSavingPOVouchertoSendEmailfromExtraField("emailvalidationone@gmail.com");
		 		 if(result)
		 		 {
		 			 return true;
		 		 }
		 		 else
		 		 {
		 			 return false;
		 		 }
		 	 
		  }
		  
		  public boolean checkSavingPOVoucher2toSendEmailfromExtraField() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		  {
		 	 boolean result=checkSavingPOVouchertoSendEmailfromExtraField("emailvalidationtwo@gmail.com");
		 	 if(result)
		 	 {
		 		 return true;
		 	 }
		 	 else
		 	 {
		 		 return false;
		 	 }
		 	 }
		  
		  
		  
		  public boolean checkSavingPOVoucher3toSendEmailfromExtraField() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		  {
		 	 boolean result=checkSavingPOVouchertoSendEmailfromExtraField("emailvalidationfour@gmail.com");
		 	 if(result)
		 	 {
		 		 return true;
		 	 }
		 	 else
		 	 {
		 		 return false;
		 	 }
		 	 }
		  
		  public boolean checkSavingPOVouchertoSendEmailfromExtraField(String narrationTxt) throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		  {
		 	 Thread.sleep(3000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		 		MRpurchaseAccountTxt.click();
		 		MRpurchaseAccountTxt.sendKeys("STD RATE COGS ACC INV");
		 		Thread.sleep(3000);
		 		MRpurchaseAccountTxt.sendKeys(Keys.TAB);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		 		vendorAccountTxt.click();
		 		vendorAccountTxt.sendKeys("Vendor B");
		 		Thread.sleep(3000);
		 		vendorAccountTxt.sendKeys(Keys.TAB);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValuetxt));
		 		departmentValuetxt.click();
		 		departmentValuetxt.sendKeys("INDIA");
		 		Thread.sleep(3000);
		 		departmentValuetxt.sendKeys(Keys.TAB);

		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PO_NarrationHeader));
		 		PO_NarrationHeader.click();
		 		PO_NarrationHeader.sendKeys(narrationTxt);
		 		Thread.sleep(3000);
		 		PO_NarrationHeader.sendKeys(Keys.TAB);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		 		select1stRow_1stColumn.click();
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		 		enter_ItemTxt.click();
		 		enter_ItemTxt.sendKeys(Keys.SPACE);
		 		

		 		Thread.sleep(1000);
		 		enter_ItemTxt.sendKeys("STD RATE COGS ITEM");
		 		Thread.sleep(2000);

		 		/*int pvvGridItemListCount = pvvGridItemList.size();
		 		for (int i = 0; i < pvvGridItemListCount; i++) {
		 			String Item = pvvGridItemList.get(i).getText();
		 			if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
		 				pvvGridItemList.get(i).click();
		 				Thread.sleep(2000);
		 				break;
		 			}
		 		}*/
		 		enter_ItemTxt.sendKeys(Keys.TAB);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		 		select1stRow_3rdColumn.click();
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		 		enter_Quantity.sendKeys("2");

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		 		select1stRow_4thColumn.click();
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		 		enter_Rate.sendKeys("10");
		 		enter_Rate.sendKeys(Keys.TAB);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		 		select2ndRow_1stColumn.click();
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		 		enter_ItemTxt.click();

		 		enter_ItemTxt.sendKeys(Keys.END);
		 		enter_ItemTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		 		enter_ItemTxt.sendKeys("STD RATE COGS ITEM");

		 		Thread.sleep(2000);

		 		enter_ItemTxt.sendKeys(Keys.TAB);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		 		select2ndRow_3rdColumn.click();
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		 		enter_Quantity.sendKeys("1");

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
		 		select2ndRow_4thColumn.click();
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		 		enter_Rate.sendKeys("10");
		 		enter_Rate.sendKeys(Keys.TAB);

		 		Thread.sleep(1000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		 		voucherSaveBtn.click();

		 		String expMessage1 = "Voucher saved successfully";

		 		String actMessage = checkValidationMessage(expMessage1);
		 		// String expMessage2 = ": 3";

		 		System.out.println("Actual Message    : " + actMessage);
		 		System.out.println("Expected Message  : " + expMessage1);

		 		if (actMessage.startsWith(expMessage1)/* && actMessage.endsWith(expMessage2) */ )

		 		{
		 			return true;
		 		} else {
		 			return false;
		 		}
		  }
		  
		  
		  
		  @FindBy(xpath="//input[@id='id_header_67108882']")
		  public static WebElement PO_NarrationHeader;
		  
		  
		  public boolean checkEmailfromExtraFieldAfterauthorizationofVoucher1() throws InterruptedException, AWTException, IOException
		  {
		 	 
		 	 

		 		Thread.sleep(7000);
		 		File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\POTriggersEmail1FromExtraField.pdf");

		 		if (Efile1.exists()) {
		 			Efile1.delete();
		 		}

		 		Robot robot = new Robot();
		 		robot.keyPress(KeyEvent.VK_CONTROL);
		 		robot.keyPress(KeyEvent.VK_T);
		 		robot.keyRelease(KeyEvent.VK_CONTROL);
		 		robot.keyRelease(KeyEvent.VK_T);

		 		Thread.sleep(4000);

		 		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

		 		System.out.println("openTabs" + openTabs);

		 		getDriver().switchTo().window(openTabs.get(0));
		 		Thread.sleep(2000);
		 		getDriver().switchTo().window(openTabs.get(1));
		 		Thread.sleep(2000);
		 		getDriver().get(
		 				"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

		 		Thread.sleep(2000);
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
		 		userNameTxt.click();

		 		userNameTxt.sendKeys("emailvalidationone@gmail.com");

		 		Thread.sleep(2000);

		 		System.out.println("User enter text");

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
		 		NextBtn.click();

		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
		 		PasswordTxt.click();

		 		PasswordTxt.sendKeys("validationone");

		 		Thread.sleep(2000);

		 		NextBtn.click();

		 		Thread.sleep(10000);

		 		int count1 = row1mailFromList.size();
		 	/*	System.out.println("No.of mails in the list" + count1);
		 		
		 		String expEmailBody="Automation Company : 08/10/2020 5:43 PM< >STD RATE COGS ITEM2.000000000010.000000000020.0000000PurOrd:25";
		 		String actEmailBody=emailBody.getText();
		 		
		 		System.err.println("Email Body *" + emailBody.getText() + "*");
		 		System.err.println("Expected Email Body *" + expEmailBody);*/

		 		if (emailBody2.getText().isEmpty() == false) {
		 			int count = emailSubjectList.size();
		 			System.err.println("No.of mails in the list" + count);

		 			for (int i = 0; i < count; i++) {
		 				String mailFromTxt = emailSubjectList.get(i).getText();
		 				Thread.sleep(2000);
		 				if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) {
		 					Thread.sleep(2000);
		 					
		 					emailSubjectList.get(i).click();
		 					break;
		 				}

		 			}

		 			Thread.sleep(3000);

		 			Thread.sleep(3000);
		 			
		 		/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(notificationCloseBtn));
		 			notificationCloseBtn.click();
		 			Thread.sleep(2000);*/
		 			
		 			
		 			getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
		 			Thread.sleep(2000);


		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
		 			mailAttachmentDownloadBtn1.click();

		 			Thread.sleep(4000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(5000);

		 			robot.keyPress(KeyEvent.VK_TAB);
		 			robot.keyRelease(KeyEvent.VK_TAB);
		 			robot.keyPress(KeyEvent.VK_TAB);
		 			robot.keyRelease(KeyEvent.VK_TAB);

		 			robot.keyPress(KeyEvent.VK_ENTER);
		 			robot.keyRelease(KeyEvent.VK_ENTER);

		 			Thread.sleep(3000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_S);
		 			robot.keyRelease(KeyEvent.VK_S);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(2000);

		 			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\POTriggersEmail1FromExtraField.exe");

		 			Thread.sleep(7000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(2000);

		 			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

		 			int actOpenWindowsCount = getDriver().getWindowHandles().size();
		 			int expOpenWindowsCount = 4;

		 			System.out.println(
		 					"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

		 			getDriver().switchTo().window(newTabs.get(3)).close();
		 			Thread.sleep(1000);
		 			getDriver().switchTo().window(newTabs.get(2)).close();
		 			Thread.sleep(1000);
		 			getDriver().switchTo().window(newTabs.get(1));
		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 			mailDeleteButton.click();

		 			Thread.sleep(2000);

		 			Thread.sleep(2000);

		 			/*
		 			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 			 * gmailUserBtn)); gmailUserBtn.click();
		 			 */
		 			
		 			ClickUsingJs(gmailUserBtn);

		 			Thread.sleep(3000);
		 			getDriver().switchTo().frame(SignOutFrame);
		 			Thread.sleep(1000);
		 			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 			Thread.sleep(6000);
		 			
		 			getDriver().get(
							"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

					Thread.sleep(2000);
		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 			removeAccountBtn.click();

		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 			removeDeleteBtn.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 			yesRemoveBtn.click();

		 			Thread.sleep(2000);

		 			getDriver().switchTo().window(newTabs.get(1)).close();
		 			Thread.sleep(2000);
		 			getDriver().switchTo().window(newTabs.get(0));
		 			Thread.sleep(2000);

		 			String actAccount1PDF = getBaseDir() + "\\autoIt\\ExportFiles\\POTriggersEmail1FromExtraField.pdf";
		 			String expAccount1PDF = getBaseDir() + "\\autoIt\\ImportFiles\\POTriggersEmail1FromExtraField.pdf";

		 			PDFUtil pdfutil = new PDFUtil();

		 			boolean result = pdfutil.compare(actAccount1PDF, expAccount1PDF);
		 			Calendar cal = Calendar.getInstance();
		 			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		 			String currentDate = df.format(cal.getTime());

		 			String actAccount1Data = pdfutil.getText(actAccount1PDF);
		 			String expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("03/01/2024", currentDate);

		 			System.err.println(actAccount1Data);
		 			System.err.println(expAccount1Data);

		 			if (actAccount1Data.equalsIgnoreCase(expAccount1Data) ) {

		 				System.err.println("PDF file is as Expected");
		 				return true;
		 			} else {

		 				System.err.println("PDF file is not as Expected");
		 				return false;
		 			}

		 		} else {

		 			System.err.println("Subject : " + emailRow1Subject.getText());

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsChckbx));
		 			selectAllMailsChckbx.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsDltBtn));
		 			selectAllMailsDltBtn.click();

		 			Thread.sleep(4000);

		 			Thread.sleep(2000);

		 			/*
		 			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 			 * gmailUserBtn)); gmailUserBtn.click();
		 			 */
		 			
		 			ClickUsingJs(gmailUserBtn);

		 			Thread.sleep(3000);
		 			getDriver().switchTo().frame(SignOutFrame);
		 			Thread.sleep(1000);
		 			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 			Thread.sleep(6000);
		 			
		 			getDriver().get(
							"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

					Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 			removeAccountBtn.click();

		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 			removeDeleteBtn.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 			yesRemoveBtn.click();

		 			Thread.sleep(2000);
		 			return false;
		 		}

		 	 
		  }
		  
		  
		  
		  public boolean checkAuthorizingVouchersFromHomePageforExtraFieldEmail() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		  {
		 	 
		 	 
		 	 getDriver().navigate().refresh();
		 		Thread.sleep(4000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		 		inventoryMenu.click();
		 		Thread.sleep(2000);
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
		 		invTransactionsMenu.click();
		 		Thread.sleep(2000);
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		 		inventoryTransactionsPurchasesMenu.click();
		 		Thread.sleep(2000);
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		 		purchasesOrdersVoucher.click();

		 		Thread.sleep(3000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		 		int vcount = authorizationStatusList.size();
		 		int count = 0;
		 		
		 		//((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,300)","");
		 		Thread.sleep(2000);
		 				
		 		/*while(count<3)
		 		{
		 			for (int i = 0; i < vcount; i++) {
		 				String VoucherAuth = authorizationStatusList.get(i).getText();
		 				String VoucherNo=voucherNoList.get(i).getText();
		 				if(VoucherAuth.equalsIgnoreCase("Pending"))
		 				{
		 					switch(VoucherNo)
		 					{
		 					case "25":
		 						grid_CheckBoxList.get(i).click();
		 						count++;
		 						break;
		 					case "26":
		 						grid_CheckBoxList.get(i).click();
		 						count++;
		 						break;
		 					case "27":
		 						grid_CheckBoxList.get(i).click();
		 						count++;
		 						break;
		 					}
		 				}
		 			}
		 		}*/
		 	
		 		
		 		for (int i = 0; i < vcount; i++) {
	 				String VoucherAuth = authorizationStatusList.get(i).getText();
	 				
	 				if(VoucherAuth.equalsIgnoreCase("Pending"))
	 				{
	 						 					
	 						grid_CheckBoxList.get(i).click();
	 				}
		 		}
	 						
		 		
		 		
		 		//((JavascriptExecutor)getDriver()).executeScript("window.scrollTo(0, 0)","");
		 		getAction().moveToElement(authorizeBtn).build().perform();
		 		
		 		 Thread.sleep(4000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
		 		authorizeBtn.click();
		 		Thread.sleep(5000);

		 		String expMessage="Voucher Approoved Successfully";
		 		String actMessage=checkValidationMessage(expMessage);
		 		Thread.sleep(2000);
		 		if (actMessage.equalsIgnoreCase(expMessage)) {
		 			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
		 			voucherhomeCloseBtn.click();*/

		 			return true;
		 		}

		 		else {
		 			return false;
		 		}
		 	
		 	 
		  }
		  
		  
		  
		  
		  public boolean checkEmailfromExtraFieldAfterauthorizationofVoucher2() throws InterruptedException, AWTException, IOException
		  {
		 	 
		 	 

		 		Thread.sleep(10000);
		 		File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\POTriggersEmail2FromExtraField.pdf");

		 		if (Efile1.exists()) {
		 			Efile1.delete();
		 		}

		 		Robot robot = new Robot();
		 		robot.keyPress(KeyEvent.VK_CONTROL);
		 		robot.keyPress(KeyEvent.VK_T);
		 		robot.keyRelease(KeyEvent.VK_CONTROL);
		 		robot.keyRelease(KeyEvent.VK_T);

		 		Thread.sleep(4000);

		 		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

		 		System.out.println("openTabs" + openTabs);

		 		getDriver().switchTo().window(openTabs.get(0));
		 		Thread.sleep(2000);
		 		getDriver().switchTo().window(openTabs.get(1));
		 		Thread.sleep(2000);
		 		getDriver().get(
		 				"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

		 		Thread.sleep(2000);
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
		 		userNameTxt.click();

		 		userNameTxt.sendKeys("emailvalidationtwo@gmail.com");

		 		Thread.sleep(2000);

		 		System.out.println("User enter text");

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
		 		NextBtn.click();

		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
		 		PasswordTxt.click();

		 		PasswordTxt.sendKeys("validationtwo");

		 		Thread.sleep(2000);

		 		NextBtn.click();

		 		Thread.sleep(10000);

		 		int count1 = row1mailFromList.size();
		 		System.out.println("No.of mails in the list" + count1);
		 		
		 		/*String expEmailBody="Automation Company : 08/10/2020 5:43 PM< >STD RATE COGS ITEM2.000000000010.000000000020.0000000PurOrd:26";
		 		String actEmailBody=emailBody.getText();
		 		
		 		System.err.println("Email Body *" + emailBody.getText() + "*");
		 		System.err.println("Expected Email Body *" + expEmailBody);*/

		 		if (emailBody2.getText().isEmpty() == false) {
		 			int count = emailSubjectList.size();
		 			System.err.println("No.of mails in the list" + count);

		 			for (int i = 0; i < count; i++) {
		 				String mailFromTxt = emailSubjectList.get(i).getText();
		 				Thread.sleep(2000);
		 				if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) {
		 					Thread.sleep(2000);
		 					
		 					emailSubjectList.get(i).click();
		 					break;
		 				}

		 			}

		 			Thread.sleep(3000);

		 			Thread.sleep(3000);
		 			
		 			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(notificationCloseBtn));
		 			notificationCloseBtn.click();
		 			Thread.sleep(2000);*/
		 			
		 			
		 			getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
		 			Thread.sleep(2000);


		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
		 			mailAttachmentDownloadBtn1.click();

		 			Thread.sleep(4000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(5000);

		 			robot.keyPress(KeyEvent.VK_TAB);
		 			robot.keyRelease(KeyEvent.VK_TAB);
		 			robot.keyPress(KeyEvent.VK_TAB);
		 			robot.keyRelease(KeyEvent.VK_TAB);

		 			robot.keyPress(KeyEvent.VK_ENTER);
		 			robot.keyRelease(KeyEvent.VK_ENTER);

		 			Thread.sleep(3000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_S);
		 			robot.keyRelease(KeyEvent.VK_S);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(2000);

		 			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\POTriggersEmail2FromExtraField.exe");

		 			Thread.sleep(7000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(2000);

		 			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

		 			int actOpenWindowsCount = getDriver().getWindowHandles().size();
		 			int expOpenWindowsCount = 4;

		 			System.out.println(
		 					"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

		 			getDriver().switchTo().window(newTabs.get(3)).close();
		 			Thread.sleep(1000);
		 			getDriver().switchTo().window(newTabs.get(2)).close();
		 			Thread.sleep(1000);
		 			getDriver().switchTo().window(newTabs.get(1));
		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 			mailDeleteButton.click();

		 			Thread.sleep(2000);

		 			Thread.sleep(2000);

		 			/*
		 			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 			 * gmailUserBtn)); gmailUserBtn.click();
		 			 */
		 			
		 			ClickUsingJs(gmailUserBtn);

		 			Thread.sleep(3000);
		 			getDriver().switchTo().frame(SignOutFrame);
		 			Thread.sleep(2000);
		 			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 			Thread.sleep(6000);
		 			getDriver().get(
							"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

					Thread.sleep(2000);
		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 			removeAccountBtn.click();

		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 			removeDeleteBtn.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 			yesRemoveBtn.click();

		 			Thread.sleep(2000);

		 			getDriver().switchTo().window(newTabs.get(1)).close();
		 			Thread.sleep(2000);
		 			getDriver().switchTo().window(newTabs.get(0));
		 			Thread.sleep(2000);

		 			String actAccount1PDF = getBaseDir() + "\\autoIt\\ExportFiles\\POTriggersEmail2FromExtraField.pdf";
		 			String expAccount1PDF = getBaseDir() + "\\autoIt\\ImportFiles\\POTriggersEmail2FromExtraField.pdf";

		 			PDFUtil pdfutil = new PDFUtil();

		 			boolean result = pdfutil.compare(actAccount1PDF, expAccount1PDF);
		 			Calendar cal = Calendar.getInstance();
		 			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		 			String currentDate = df.format(cal.getTime());

		 			String actAccount1Data = pdfutil.getText(actAccount1PDF);
		 			String expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("14/02/2024", currentDate);

		 			System.err.println(actAccount1Data);
		 			System.err.println(expAccount1Data);

		 			if (actAccount1Data.equalsIgnoreCase(expAccount1Data) ) {

		 				System.err.println("PDF file is as Expected");
		 				return true;
		 			} else {

		 				System.err.println("PDF file is not as Expected");
		 				return false;
		 			}

		 		} else {

		 			System.err.println("Subject : " + emailRow1Subject.getText());

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsChckbx));
		 			selectAllMailsChckbx.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsDltBtn));
		 			selectAllMailsDltBtn.click();

		 			Thread.sleep(4000);

		 			Thread.sleep(2000);

		 			/*
		 			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 			 * gmailUserBtn)); gmailUserBtn.click();
		 			 */
		 			
		 			ClickUsingJs(gmailUserBtn);

		 			Thread.sleep(3000);
		 			getDriver().switchTo().frame(SignOutFrame);
		 			Thread.sleep(1000);
		 			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 			Thread.sleep(6000);
		 			getDriver().get(
							"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

					Thread.sleep(2000);
		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 			removeAccountBtn.click();

		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 			removeDeleteBtn.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 			yesRemoveBtn.click();

		 			Thread.sleep(2000);
		 			return false;
		 		}

		 	 
		  }
		  
		  
		  
		  public boolean checkEmailfromExtraFieldAfterauthorizationofVoucher3() throws InterruptedException, AWTException, IOException
		  {
		 	 
		 	 

		 		Thread.sleep(8000);
		 		File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\POTriggersEmail3FromExtraField.pdf");

		 		if (Efile1.exists()) {
		 			Efile1.delete();
		 		}

		 		Robot robot = new Robot();
		 		robot.keyPress(KeyEvent.VK_CONTROL);
		 		robot.keyPress(KeyEvent.VK_T);
		 		robot.keyRelease(KeyEvent.VK_CONTROL);
		 		robot.keyRelease(KeyEvent.VK_T);

		 		Thread.sleep(4000);

		 		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

		 		System.out.println("openTabs" + openTabs);

		 		getDriver().switchTo().window(openTabs.get(0));
		 		Thread.sleep(2000);
		 		getDriver().switchTo().window(openTabs.get(1));
		 		Thread.sleep(2000);
		 		getDriver().get(
		 				"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

		 		Thread.sleep(2000);
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
		 		userNameTxt.click();

		 		userNameTxt.sendKeys("emailvalidationfour@gmail.com");

		 		Thread.sleep(2000);

		 		System.out.println("User enter text");

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
		 		NextBtn.click();

		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
		 		PasswordTxt.click();

		 		PasswordTxt.sendKeys("validationfour");

		 		Thread.sleep(2000);

		 		NextBtn.click();

		 		Thread.sleep(10000);

		 		int count1 = row1mailFromList.size();
		 		System.out.println("No.of mails in the list" + count1);
		 		
		 		String expEmailBody="Automation Company : 08/10/2020 5:43 PM< >STD RATE COGS ITEM2.000000000010.000000000020.0000000PurOrd:27";
		 		String actEmailBody=emailBody.getText();
		 		
		 		System.err.println("Email Body *" + emailBody.getText() + "*");
		 		System.err.println("Expected Email Body *" + expEmailBody);

		 		if (emailBody.getText().isEmpty() == false) {
		 			int count = emailSubjectList.size();
		 			System.err.println("No.of mails in the list" + count);

		 			for (int i = 0; i < count; i++) {
		 				String mailFromTxt = emailSubjectList.get(i).getText();
		 				Thread.sleep(2000);
		 				if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) {
		 					Thread.sleep(2000);
		 					
		 					emailSubjectList.get(i).click();
		 					break;
		 				}

		 			}

		 			Thread.sleep(3000);

		 			Thread.sleep(3000);
		 			
		 			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(notificationCloseBtn));
		 			notificationCloseBtn.click();
		 			Thread.sleep(2000);
		 			*/
		 			
		 			getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
		 			Thread.sleep(2000);


		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
		 			mailAttachmentDownloadBtn1.click();

		 			Thread.sleep(4000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(5000);

		 			robot.keyPress(KeyEvent.VK_TAB);
		 			robot.keyRelease(KeyEvent.VK_TAB);
		 			robot.keyPress(KeyEvent.VK_TAB);
		 			robot.keyRelease(KeyEvent.VK_TAB);

		 			robot.keyPress(KeyEvent.VK_ENTER);
		 			robot.keyRelease(KeyEvent.VK_ENTER);

		 			Thread.sleep(3000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_S);
		 			robot.keyRelease(KeyEvent.VK_S);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(2000);

		 			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\POTriggersEmail3FromExtraField.exe");

		 			Thread.sleep(7000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(2000);

		 			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

		 			int actOpenWindowsCount = getDriver().getWindowHandles().size();
		 			int expOpenWindowsCount = 4;

		 			System.out.println(
		 					"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

		 			getDriver().switchTo().window(newTabs.get(3)).close();
		 			Thread.sleep(1000);
		 			getDriver().switchTo().window(newTabs.get(2)).close();
		 			Thread.sleep(1000);
		 			getDriver().switchTo().window(newTabs.get(1));
		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 			mailDeleteButton.click();

		 			Thread.sleep(2000);

		 			Thread.sleep(2000);

		 			/*
		 			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 			 * gmailUserBtn)); gmailUserBtn.click();
		 			 */
		 			
		 			ClickUsingJs(gmailUserBtn);

		 			Thread.sleep(3000);
		 			getDriver().switchTo().frame(SignOutFrame);
		 			Thread.sleep(1000);
		 			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 			Thread.sleep(6000);
		 			getDriver().get("https://accounts.google.com/AccountChooser/signinchooser?service=mail&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&flowName=GlifWebSignIn&flowEntry=AccountChooser");
					 Thread.sleep(2000);
		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 			removeAccountBtn.click();

		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 			removeDeleteBtn.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 			yesRemoveBtn.click();

		 			Thread.sleep(2000);

		 			getDriver().switchTo().window(newTabs.get(1)).close();
		 			Thread.sleep(2000);
		 			getDriver().switchTo().window(newTabs.get(0));
		 			Thread.sleep(2000);

		 			String actAccount1PDF = getBaseDir() + "\\autoIt\\ExportFiles\\POTriggersEmail3FromExtraField.pdf";
		 			String expAccount1PDF = getBaseDir() + "\\autoIt\\ImportFiles\\POTriggersEmail3FromExtraField.pdf";

		 			PDFUtil pdfutil = new PDFUtil();

		 			boolean result = pdfutil.compare(actAccount1PDF, expAccount1PDF);
		 			Calendar cal = Calendar.getInstance();
		 			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		 			String currentDate = df.format(cal.getTime());

		 			String actAccount1Data = pdfutil.getText(actAccount1PDF);
		 			String expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("14/02/2024", currentDate);

		 			System.err.println(actAccount1Data);
		 			System.err.println(expAccount1Data);

		 			if (actAccount1Data.equalsIgnoreCase(expAccount1Data) ) {

		 				System.err.println("PDF file is as Expected");
		 				return true;
		 			} else {

		 				System.err.println("PDF file is not as Expected");
		 				return false;
		 			}

		 		} else {

		 			System.err.println("Subject : " + emailRow1Subject.getText());

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsChckbx));
		 			selectAllMailsChckbx.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectAllMailsDltBtn));
		 			selectAllMailsDltBtn.click();

		 			Thread.sleep(4000);

		 			Thread.sleep(2000);

		 			/*
		 			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 			 * gmailUserBtn)); gmailUserBtn.click();
		 			 */
		 			
		 			ClickUsingJs(gmailUserBtn);

		 			Thread.sleep(3000);
		 			getDriver().switchTo().frame(SignOutFrame);
		 			Thread.sleep(1000);
		 			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 			Thread.sleep(6000);
		 			getDriver().get(
							"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

					Thread.sleep(2000);
		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 			removeAccountBtn.click();

		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 			removeDeleteBtn.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 			yesRemoveBtn.click();

		 			Thread.sleep(2000);
		 			return false;
		 		}

		 	 
		  }
		  
		  
	
	
	
	public boolean checkSavingTriggerinPOVouchertoSendMailtoRole() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		

		getDriver().navigate().refresh();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Configure Transactions");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();

		Thread.sleep(2000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseOrdersBtn));
		purchaseOrdersBtn.click();

		Thread.sleep(2000);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Authorize Trigger");
		Thread.sleep(4000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(10000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailTxtBox));
		if (sendEmailTxtBox.isSelected() == false) {
			sendEmailTxtBox.click();
		}
		/*
		 * sendEmailTxtBox.sendKeys(Keys.HOME, Keys.SHIFT, Keys.END);
		 * sendEmailTxtBox.sendKeys("POTriggersEmailLayout");
		 */
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerChkBoxSelected));
		if (customerChkBox.isSelected() == true) {
			customerChkBoxSelected.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginChkBoxSelected));
		if (loginChkBox.isSelected() == true) {
			loginChkBoxSelected.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFieldChkBoxSelected));
		if (extraFieldChkBox.isSelected() == true) {
			extraFieldChkBoxSelected.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(roleChkBoxSelected));
		if (roleChkBox.isSelected() == false) {
			roleChkBoxSelected.click();
		}
		
		getAction().moveToElement(roleTxtBox).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(roleTxtBox));
		roleTxtBox.click();
		Thread.sleep(4000);
	
		getAction().moveToElement(role1ChkBox).build().perform();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(role1ChkBox));
		role1ChkBox.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		Thread.sleep(2000);

		String expSaveMessage = "Data saved successfully";
		String actSaveMessage = checkValidationMessage(expSaveMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();

		if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
			return true;
		} else {
			return false;
		}


		
	}


	public boolean checkSavingPOVoucherforRoleEmailValidation() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		

		
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
		invTransactionsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();
		Thread.sleep(4000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		MRpurchaseAccountTxt.click();
		MRpurchaseAccountTxt.sendKeys("STD RATE COGS ACC INV");
		Thread.sleep(3000);
		MRpurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();
		vendorAccountTxt.sendKeys("Vendor B");
		Thread.sleep(3000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValuetxt));
		departmentValuetxt.click();
		departmentValuetxt.sendKeys("INDIA");
		Thread.sleep(3000);
		departmentValuetxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		

		Thread.sleep(1000);
		enter_ItemTxt.sendKeys("STD RATE COGS ITEM");
		Thread.sleep(2000);

		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys("2");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();

		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		enter_ItemTxt.sendKeys("STD RATE COGS ITEM");

		Thread.sleep(2000);

		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		select2ndRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys("1");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
		select2ndRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		String expMessage1 = "Voucher saved successfully";

		String actMessage = checkValidationMessage(expMessage1);
		// String expMessage2 = ": 3";

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage1);

		if (actMessage.startsWith(expMessage1)/* && actMessage.endsWith(expMessage2) */ )

		{
			return true;
		} else {
			return false;
		}


	}

		  
	public boolean checkEmailValidationforUser1inRoleAfterAuth() throws InterruptedException, AWTException, IOException
	{
		Thread.sleep(6000);
		
		File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser1forAuth.pdf");
		
		if(Efile1.exists())
		{
			Efile1.delete();
			System.out.println("File deleted");
		}
		
		
		Robot robot = new Robot();   
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);
		
		Thread.sleep(6000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

	 	System.out.println("openTabs"+openTabs);

	 	getDriver().switchTo().window(openTabs.get(0));
	 	Thread.sleep(2000);
	 	getDriver().switchTo().window(openTabs.get(1));
		

	 	getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");
	 	
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
	 	userNameTxt.click();
	 	
	 	userNameTxt.sendKeys("emailvalidationtwo@gmail.com");
	 	
	 	Thread.sleep(2000);
	 	
	 	System.out.println("User enter text");
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
	 	NextBtn.click();
	 	
	 	Thread.sleep(2000);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
	 	PasswordTxt.click();
	 	
	 	PasswordTxt.sendKeys("validationtwo");
	 	
	 	Thread.sleep(2000);
	 	
	 	NextBtn.click();
	 	
	 	Thread.sleep(30000);
	 	
	 boolean result=false;
	
	 		int count = emailSubjectList.size();
	 		
	 		System.out.println(emailBody2.getText().isEmpty());
	 		if (emailBody2.getText().isEmpty()==false) 
		 	 	{
	 		
		 	for (int i = 0; i < count; i++) 
	 	{
	 		String mailFromTxt = emailSubjectList.get(i).getText();
	 		System.out.println(mailFromTxt);
	 	 	
	 		if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) 
	 		{
	 			emailSubjectList.get(i).click();     
	 			break;
	 		}
	 	}
		 	
	 			 	
		 	Thread.sleep(8000);
		
		
		
		if(refreshBtn.isDisplayed())
		{
			
			System.out.println("Subject is not matching");
		 	
		}
		
	 		
	 	
		else
		{
			
			System.out.println("Subject is  matching");
			getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
		mailAttachmentDownloadBtn1.click();
		
		Thread.sleep(6000);
		
		 
		robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_CONTROL);


	Thread.sleep(5000);

	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);
	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);

	robot.keyPress(KeyEvent.VK_ENTER);
	robot.keyRelease(KeyEvent.VK_ENTER);

	Thread.sleep(3000);

	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_CONTROL);
		
	Thread.sleep(2000);

			
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailValidationforUser1forAuth.exe");
			
			Thread.sleep(8000);
			
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(4000);
			
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
			int actOpenWindowsCount = getDriver().getWindowHandles().size();
			int expOpenWindowsCount = 4;
			
			System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
			
			getDriver().switchTo().window(newTabs.get(3)).close();
		 	Thread.sleep(2000);
		 	getDriver().switchTo().window(newTabs.get(2)).close();
		 	Thread.sleep(2000);
		 	getDriver().switchTo().window(newTabs.get(1));
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 	mailDeleteButton.click();
		 	
		 	Thread.sleep(2000);
		 	

		 	Thread.sleep(2000);
		 	
		 	String actAccount1PDF = getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser1forAuth.pdf";
			String expAccount1PDF = getBaseDir()+"\\autoIt\\ImportFiles\\EmailValidationforUser1forAuth.pdf";
			
		
		 
			PDFUtil pdfutil = new PDFUtil();
		
			boolean result1 = pdfutil.compare(actAccount1PDF, expAccount1PDF);
			Calendar cal=Calendar.getInstance();
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = df.format(cal.getTime());
			
			String  actAccount1Data = pdfutil.getText(actAccount1PDF);
			String  expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("16/04/2024", currentDate);
					
			System.err.println(actAccount1Data);
			System.err.println(expAccount1Data);
			
			if(actAccount1Data.equalsIgnoreCase(expAccount1Data))
			{
				result=true;
			}
		 	
		}
			
		 	 	}	
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());	
			ClickUsingJs(gmailUserBtn);

			Thread.sleep(3000);
			getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(2000);
			
			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
			
			getDriver().get(
					"https://accounts.google.com/AccountChooser/signinchooser?service=mail&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&flowName=GlifWebSignIn&flowEntry=AccountChooser");
			Thread.sleep(2000);
			
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 	removeAccountBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 	removeDeleteBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 	yesRemoveBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
			getDriver().switchTo().window(newTabs.get(1)).close();
		 	Thread.sleep(2000);
		 	
		 	
		 	getDriver().switchTo().window(newTabs.get(0));
		 	Thread.sleep(2000);
		 
			
			
			
		 	 		
		 	if (result && count!=0) 
			{
			
				System.err.println("PDF file is as Expected");
				return true;
			}
			else
			{
									
				System.err.println("PDF file is not as Expected");
				return false;
			}
			
			
		} 
		

	public boolean checkEmailValidationforUser2inRoleAfterAuth() throws InterruptedException, AWTException, IOException
	{
		Thread.sleep(6000);
	File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser2forAuth.pdf");
		
		if(Efile1.exists())
		{
			Efile1.delete();
			System.out.println("File deleted");
		}
		
		
		Robot robot = new Robot();   
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);
		
		Thread.sleep(4000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

	 	System.out.println("openTabs"+openTabs);

	 	getDriver().switchTo().window(openTabs.get(0));
	 	Thread.sleep(2000);
	 	getDriver().switchTo().window(openTabs.get(1));
		

	 	getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");
	 	
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
	 	userNameTxt.click();
	 	
	 	userNameTxt.sendKeys("emailvalidationfour@gmail.com");
	 	
	 	Thread.sleep(2000);
	 	
	 	System.out.println("User enter text");
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
	 	NextBtn.click();
	 	
	 	Thread.sleep(2000);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
	 	PasswordTxt.click();
	 	
	 	PasswordTxt.sendKeys("validationfour");
	 	
	 	Thread.sleep(2000);
	 	
	 	NextBtn.click();
	 	
	 	Thread.sleep(30000);
	 	
	 boolean result=false;
	
	 		int count = emailSubjectList.size();
	 		System.err.println("No.of mails in the list"+count);
			
	 		
	 		System.out.println(emailBody2.getText().isEmpty());
	 		
	 		if (emailBody2.getText().isEmpty()==false) 
		 	 	{
	 		
		 	for (int i = 0; i < count; i++) 
	 	{
	 		String mailFromTxt = emailSubjectList.get(i).getText();
	 		System.out.println(mailFromTxt);
	 		Thread.sleep(2000);
	 		if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) 
	 		{
	 			emailSubjectList.get(i).click();     
	 			break;
	 		}
	 	}
		 	
	 			 	
		 	Thread.sleep(8000);
		
		//Thread.sleep(3000);
		
		if(refreshBtn.isDisplayed())
		{
			
			System.out.println("Subject is not matching");
		 	
		}
		
	 		
	 	
		else
		{
			
			System.out.println("Subject is  matching");
			getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
		mailAttachmentDownloadBtn1.click();
		
		Thread.sleep(6000);
		
		 
		robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_CONTROL);


	Thread.sleep(5000);

	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);
	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);

	robot.keyPress(KeyEvent.VK_ENTER);
	robot.keyRelease(KeyEvent.VK_ENTER);

	Thread.sleep(3000);

	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_CONTROL);
		
	Thread.sleep(4000);

			
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailValidationforUser2forAuth.exe");
			
			Thread.sleep(8000);
			
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(2000);
			
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
			int actOpenWindowsCount = getDriver().getWindowHandles().size();
			int expOpenWindowsCount = 4;
			
			System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
			
			getDriver().switchTo().window(newTabs.get(3)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(2)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(1));
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 	mailDeleteButton.click();
		 	
		 	Thread.sleep(2000);
		 	

		 	Thread.sleep(2000);
		 	
		 	String actAccount1PDF = getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser2forAuth.pdf";
			String expAccount1PDF = getBaseDir()+"\\autoIt\\ImportFiles\\EmailValidationforUser2forAuth.pdf";
			
		
		 
			PDFUtil pdfutil = new PDFUtil();
		
			boolean result1 = pdfutil.compare(actAccount1PDF, expAccount1PDF);
			Calendar cal=Calendar.getInstance();
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = df.format(cal.getTime());
			
			String  actAccount1Data = pdfutil.getText(actAccount1PDF);
			String  expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("08/05/2024", currentDate);
					
			System.err.println(actAccount1Data);
			System.err.println(expAccount1Data);
			
			if(actAccount1Data.equalsIgnoreCase(expAccount1Data))
			{
				result=true;
			}
		 	
		}
			
		 	 	}	
		 	
			Thread.sleep(2000);
		 	getDriver().navigate().refresh();
		 	Thread.sleep(8000);
		 	
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());	
			ClickUsingJs(gmailUserBtn);

			Thread.sleep(3000);
			getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(2000);
			
			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
			
			getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
			
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 	removeAccountBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 	removeDeleteBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 	yesRemoveBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
			getDriver().switchTo().window(newTabs.get(1)).close();
		 	Thread.sleep(2000);
		 	
		 	
		 	getDriver().switchTo().window(newTabs.get(0));
		 	Thread.sleep(2000);
		 
			
			
			
		 	 		
		 	if (result && count!=0) 
			{
			
				System.err.println("PDF file is as Expected");
				return true;
			}
			else
			{
									
				System.err.println("PDF file is not as Expected");
				return false;
			}
			
			
		} 
		

	public boolean checkEmailValidationforUser3inRoleAfterAuth() throws InterruptedException, AWTException, IOException
	{
	File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser3forAuth.pdf");
		
		if(Efile1.exists())
		{
			Efile1.delete();
			System.out.println("File deleted");
		}
		
		
		Robot robot = new Robot();   
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_T);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		robot.keyRelease(KeyEvent.VK_T);
		
		Thread.sleep(4000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

	 	System.out.println("openTabs"+openTabs);

	 	getDriver().switchTo().window(openTabs.get(0));
	 	Thread.sleep(2000);
	 	getDriver().switchTo().window(openTabs.get(1));
		

	 	getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");
	 	
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
	 	userNameTxt.click();
	 	
	 	userNameTxt.sendKeys("emailvalidationfive@gmail.com");
	 	
	 	Thread.sleep(2000);
	 	
	 	System.out.println("User enter text");
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
	 	NextBtn.click();
	 	
	 	Thread.sleep(2000);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
	 	PasswordTxt.click();
	 	
	 	PasswordTxt.sendKeys("validationfive");
	 	
	 	Thread.sleep(2000);
	 	
	 	NextBtn.click();
	 	
	 	Thread.sleep(30000);
	 	
	 	boolean result=false;
	 
	 		int count = emailSubjectList.size();
	 		System.err.println("No.of mails in the list"+count);
			if (emailBody2.getText().isEmpty()==false) 
		 	 	{
	 		
		 	for (int i = 0; i < count; i++) 
	 	{
	 		String mailFromTxt = emailSubjectList.get(i).getText();
	 		System.out.println(mailFromTxt);
	 	 	
	 		if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher For Authorization")) 
	 		{
	 			emailSubjectList.get(i).click();     
	 			break;
	 		}
	 	}
		 	
	 			 	
		 	Thread.sleep(8000);
		
		//Thread.sleep(3000);
		
		if(refreshBtn.isDisplayed())
		{
			
			System.out.println("Subject is not matching");
		 	
		}
		
	 		
	 	
		else
		{
			
			System.out.println("Subject is  matching");
			getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
		mailAttachmentDownloadBtn1.click();
		
		Thread.sleep(3000);
		
		 
		robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_J);
	robot.keyRelease(KeyEvent.VK_CONTROL);


	Thread.sleep(5000);

	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);
	robot.keyPress(KeyEvent.VK_TAB);
	robot.keyRelease(KeyEvent.VK_TAB);

	robot.keyPress(KeyEvent.VK_ENTER);
	robot.keyRelease(KeyEvent.VK_ENTER);

	Thread.sleep(3000);

	robot.keyPress(KeyEvent.VK_CONTROL);
	robot.keyPress(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_S);
	robot.keyRelease(KeyEvent.VK_CONTROL);
		
	Thread.sleep(2000);

			
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailValidationforUser3forAuth.exe");
			
			Thread.sleep(4000);
			
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(2000);
			
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
			int actOpenWindowsCount = getDriver().getWindowHandles().size();
			int expOpenWindowsCount = 4;
			
			System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
			
			getDriver().switchTo().window(newTabs.get(3)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(2)).close();
		 	Thread.sleep(1000);
		 	getDriver().switchTo().window(newTabs.get(1));
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 	mailDeleteButton.click();
		 	
		 	Thread.sleep(2000);
		 	

		 	Thread.sleep(2000);
		 	
		 	String actAccount1PDF = getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser3forAuth.pdf";
			String expAccount1PDF = getBaseDir()+"\\autoIt\\ImportFiles\\EmailValidationforUser3forAuth.pdf";
			
		
		 
			PDFUtil pdfutil = new PDFUtil();
		
			boolean result1 = pdfutil.compare(actAccount1PDF, expAccount1PDF);
			Calendar cal=Calendar.getInstance();
			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
			String currentDate = df.format(cal.getTime());
			
			String  actAccount1Data = pdfutil.getText(actAccount1PDF);
			String  expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("16/04/2024", currentDate);
					
			System.err.println(actAccount1Data);
			System.err.println(expAccount1Data);
			
			if(actAccount1Data.equalsIgnoreCase(expAccount1Data))
			{
				result=true;
			}
		 	
		}
			
		 	 	}	
		 	
		 	
		 	
			Thread.sleep(2000);
		 	getDriver().navigate().refresh();
		 	Thread.sleep(6000);
		 	
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());	
			ClickUsingJs(gmailUserBtn);

			Thread.sleep(3000);
			getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(2000);
			
			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
			
			getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
			
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 	removeAccountBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 	removeDeleteBtn.click();
		 	
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 	yesRemoveBtn.click();
		 	
		 	Thread.sleep(1000);
		 	
			getDriver().switchTo().window(newTabs.get(1)).close();
		 	Thread.sleep(2000);
		 	
		 	
		 	getDriver().switchTo().window(newTabs.get(0));
		 	Thread.sleep(2000);
		 
			
			
			
		 	 		
		 	if (result && count!=0) 
			{
			
				System.err.println("PDF file is as Expected");
				return true;
			}
			else
			{
									
				System.err.println("PDF file is not as Expected");
				return false;
			}
			
			
		} 
		  
		  
		  
		  
	@FindBy(xpath="(//*[@id='12'])[2]/span")
	private static WebElement  triggerAccountExpansionBtn;
   
	@FindBy(xpath="(//*[@id='5002'])[1]")
	private static WebElement  triggerAccountName;	  
		  
		  
		  
		  @FindBy(xpath="//i[@class='icon-settings hiconright2']")
		  public static WebElement receiptsVoucherSettingsBtn;
		  
		  
		  @FindBy(xpath = "(//ul[@class='treeview filtertreedata col-xs-12']//a[@id='4']//span[@class='icon-right-arrow'])[1]")
		  public static WebElement triggerCashBankACExpansionBtn;
		  
		  
		  
		  
		  public boolean checkSavingEmailTemplateinTriggerinReceiptsVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		  { 
			  Thread.sleep(2000);
			
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			 financialsMenu.click();
			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			 financialsTransactionMenu.click();
			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
			 cashAndBankMenu.click();
			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsVoucher));
			 receiptsVoucher.click();
			 Thread.sleep(4000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsVoucherSettingsBtn));
			 receiptsVoucherSettingsBtn.click();
			 
			

				Thread.sleep(4000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
				triggersBtn.click();

				Thread.sleep(4000);

				

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
				editLayoutTriggerName.click();
				editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
				Thread.sleep(1000);
				editLayoutTriggerName.sendKeys("Receipts On Save  Trigger");
				Thread.sleep(4000);
				editLayoutTriggerName.sendKeys(Keys.TAB);
				Thread.sleep(4000);

				triggersConjuctionDrpdwn.click();
				Thread.sleep(2000);
				triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);
				triggersConjuctionDrpdwn.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
				triggersSelectFieldTxt.click();

				Thread.sleep(3000);

				//getAction().moveToElement(triggerCashBankACExpansionBtn).build().perform();
				//Thread.sleep(1000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerAccountExpansionBtn));
				triggerAccountExpansionBtn.click();

				Thread.sleep(3000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerAccountName));
				triggerAccountName.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
				triggersSelectOperdrpdwn.click();
				triggersSelectOperdrpdwn.sendKeys("Equal");

				Thread.sleep(5000);
				triggersSelectOperdrpdwn.sendKeys(Keys.TAB);

				Thread.sleep(5000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
				triggersCompareWithdrpdwn.click();
				triggersCompareWithdrpdwn.sendKeys("value");
				// triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
				Thread.sleep(2000);
				triggersCompareWithdrpdwn.sendKeys(Keys.TAB);

				Thread.sleep(2000);

				triggersValueTxt.sendKeys("Vendor B");
				Thread.sleep(1500);
				triggersValueTxt.sendKeys(Keys.TAB);

				Thread.sleep(1500);

				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
				sendEmailChkBoxSelected.click();
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createNewEmailButton));
				createNewEmailButton.click();
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailNewBtn));
				emailNewBtn.click();
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailNewLayoutBtn));
				emailNewLayoutBtn.click();
				Thread.sleep(4000);

				Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\ReceiptsEmailInvoice.exe");

				Thread.sleep(200000);

				getAction().moveToElement(emailSaveBtn).build().perform();
				Thread.sleep(1000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveBtn));
				emailSaveBtn.click();
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveOptionBtn));
				emailSaveOptionBtn.click();
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSavePopupTitle));
				boolean actTitleDisplay = emailSavePopupTitle.isDisplayed();
				boolean expTitleDisplay = true;

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveTxt));
				emailSaveTxt.click();
				emailSaveTxt.sendKeys("ReceiptsTriggersEmailLayout");
				Thread.sleep(1500);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSaveLayoutBtn));
				emailSaveLayoutBtn.click();
				Thread.sleep(2000);
				String expSaveLayotMsg = "Layout saved successfully";
				String actSaveLayoutMsg = checkValidationMessage(expSaveLayotMsg);

				System.out.println("Actual Message: 	" + actSaveLayoutMsg);
				System.out.println("Expected Message:	" + expSaveLayotMsg);

				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailLayoutExitBtn));
				emailLayoutExitBtn.click();
				Thread.sleep(2000);
				
				
				/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
				 financialsMenu.click();
				 Thread.sleep(2000);
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
				 financialsTransactionMenu.click();
				 Thread.sleep(2000);
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
				 cashAndBankMenu.click();
				 Thread.sleep(2000);
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsVoucher));
				 receiptsVoucher.click();
				 Thread.sleep(4000);
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receiptsVoucherSettingsBtn));
				 receiptsVoucherSettingsBtn.click();
				 
				

					Thread.sleep(4000);
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
					triggersBtn.click();

					Thread.sleep(4000);

					

					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
					editLayoutTriggerName.click();
					editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
					Thread.sleep(1000);
					editLayoutTriggerName.sendKeys("Receipts On Save  Trigger");
					Thread.sleep(4000);
					editLayoutTriggerName.sendKeys(Keys.TAB);
					Thread.sleep(4000);
*/
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
					
					if(sendEmailChkBox.isSelected()==false)
					{
						sendEmailChkBoxSelected.click();
					}
					Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailTxtBox));
				sendEmailTxtBox.click();
				sendEmailTxtBox.sendKeys(Keys.HOME, Keys.SHIFT, Keys.END);
				sendEmailTxtBox.sendKeys("ReceiptsTriggersEmailLayout");
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailSubjectTxt));
				emailSubjectTxt.sendKeys("Receipts Email Trigger");
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateBtn));
				emailHeaderTemplateBtn.click();
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newButtonInCreateTemplate));
				newButtonInCreateTemplate.click();
				Thread.sleep(1500);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplates));
				emailTemplates.sendKeys("Receipts Email Template");
				emailTemplates.sendKeys(Keys.TAB);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
				companyNameTemplateArea.click();
				Thread.sleep(1500);

				Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\ReceiptsEmailTemplate.exe");

				Thread.sleep(140000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
				saveBtnInEmailTemplateHeader.click();

				Thread.sleep(2000);
				System.err.println("Template save button clicked");

				String expEmailTemplateSaveMsg = "Template saved successfully";
				String actEmailTemplateSaveMsg = checkValidationMessage(expEmailTemplateSaveMsg);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
				ClickUsingJs(closeButtonInCreateTemplate);
				//closeButtonInCreateTemplate.click();
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateTxt));
				emailHeaderTemplateTxt.click();
				emailHeaderTemplateTxt.sendKeys("Receipts Email Template");
				Thread.sleep(2000);

				

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerChkBoxSelected));
				if (customerChkBox.isSelected() == false) {
					customerChkBoxSelected.click();
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
				editLayoutTriggersSaveBtn.click();
				Thread.sleep(2500);

				String expSaveMessage = "Data saved successfully";
				String actSaveMessage = checkValidationMessage(expSaveMessage);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
				updateBtn.click();

				String expUpdateMessage = "Data saved successfully";
				String actUpdateMessage = checkValidationMessage(expUpdateMessage);
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
				settings_closeBtn.click();

				if (actTitleDisplay == expTitleDisplay && actSaveLayoutMsg.equalsIgnoreCase(expSaveLayotMsg)) {
					return true;
				} else {
					return false;
				}
				
				
			

			 
		 }
		 
		  
		  @FindBy(xpath="(//*[@class='icon-new hiconright2'])[1]")
			public static WebElement newbtn;
		 
		  
		  public boolean checkSaveReceiptsVoucherafterEmailInvoiceCreated() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		  {
			  
			  Thread.sleep(8000);
		 	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newbtn));
		 	 newbtn.click();
		 	 Thread.sleep(2000);
		 	 
		 	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		 	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		 	 vendorAccountTxt.click();

		 		vendorAccountTxt.sendKeys(Keys.SPACE);

		 		ArrayList<String> actVendorAccountList = new ArrayList<String>();

		 		int count = vendorAccountListCount.size();

		 		System.err.println(count);

		 		for (int i = 0; i < count; i++) {
		 			String data = vendorAccountListCount.get(i).getText();

		 			if (data.equalsIgnoreCase("Accounts Receivable")) {
		 				vendorAccountListCount.get(i).click();
		 				break;
		 			}
		 		}

		 		vendorAccountTxt.sendKeys(Keys.TAB);

		 		Thread.sleep(3000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValuetxt));
		 		departmentValuetxt.click();
		 		departmentValuetxt.sendKeys("INDIA");
		 		Thread.sleep(3000);
		 		departmentValuetxt.sendKeys(Keys.TAB);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		 		select1stRow_1stColumn.click();
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_DebitACTxt));
		 		enter_DebitACTxt.click();
		 		enter_DebitACTxt.sendKeys(Keys.SPACE);

		 		Thread.sleep(2000);
		 		
		 		enter_DebitACTxt.sendKeys("Vendor B");
		 		Thread.sleep(1000);
		 		
		 		enter_DebitACTxt.sendKeys(Keys.TAB);
		 	 
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		 		enter_Amount.sendKeys("100");
		 		Thread.sleep(2000);
		 		enter_Amount.sendKeys(Keys.TAB);
		 		
		 		
		 		Thread.sleep(4000);

		 	
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		 		billRefNewReferenceTxt.click();
		 		Thread.sleep(1000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		 		pickBtn.click();

		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		 		Bill_OkBtn.click();

		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		 		voucherSaveBtn.click();
		 		
		 		Thread.sleep(2000);

		 		String expMessage1 = "Voucher saved successfully";

		 		String actMessage = checkValidationMessage(expMessage1);
		 		String expMessage2 = ": 1";

		 		System.out.println("Actual Message    : " + actMessage);
		 		System.out.println("Expected Message  : " + expMessage1);

		 		if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)) {
		 			return true;
		 		} else {
		 			return false;
		 		}
		 		
		  }
		  
		  
		  public boolean checkEmailValidationAfterReceiptsVoucherSaved() throws InterruptedException, AWTException, IOException
		  {
		 	 


		 		Thread.sleep(4000);
		 		File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\ReceiptsEmailTrigger.pdf");

		 		if (Efile1.exists()) {
		 			Efile1.delete();
		 		}

		 		Robot robot = new Robot();
		 		robot.keyPress(KeyEvent.VK_CONTROL);
		 		robot.keyPress(KeyEvent.VK_T);
		 		robot.keyRelease(KeyEvent.VK_CONTROL);
		 		robot.keyRelease(KeyEvent.VK_T);

		 		Thread.sleep(4000);

		 		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

		 		System.out.println("openTabs" + openTabs);

		 		getDriver().switchTo().window(openTabs.get(0));
		 		Thread.sleep(2000);
		 		getDriver().switchTo().window(openTabs.get(1));
		 		Thread.sleep(2000);
		 		getDriver().get(
		 				"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

		 		Thread.sleep(2000);
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
		 		userNameTxt.click();

		 		userNameTxt.sendKeys("emailvalidationone@gmail.com");

		 		Thread.sleep(2000);

		 		System.out.println("User enter text");

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
		 		NextBtn.click();

		 		Thread.sleep(2000);

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
		 		PasswordTxt.click();

		 		PasswordTxt.sendKeys("validationone");

		 		Thread.sleep(2000);

		 		NextBtn.click();

		 		Thread.sleep(10000);

		 		int count1 = row1mailFromList.size();
		 		/*System.out.println("No.of mails in the list" + count1);
		 		
		 		String expEmailBody="Automation Company : 08/10/2020 5:43 PM< >Vendor B100.0000000Rct:1";
		 		String actEmailBody=emailBody1.getText();
		 		
		 		System.err.println("Email Body " + emailBody1.getText());
		 		System.err.println("Expected Email Body *" + expEmailBody);*/

		 		if (emailBody2.getText().isEmpty() == false) {
		 			int count = emailSubjectList.size();
		 			System.err.println("No.of mails in the list" + count);

		 			for (int i = 0; i < count; i++) {
		 				String mailFromTxt = emailSubjectList.get(i).getText();
		 				Thread.sleep(2000);
		 				if (mailFromTxt.equalsIgnoreCase("Receipts Email Trigger")) {
		 					emailSubjectList.get(i).click();
		 					break;
		 				}

		 			}

		 			Thread.sleep(3000);

		 			Thread.sleep(3000);
		 		/*	
		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(notificationCloseBtn));
		 			notificationCloseBtn.click();
		 			Thread.sleep(2000);*/

		 			getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
		 			mailAttachmentDownloadBtn1.click();

		 			Thread.sleep(3000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(5000);

		 			robot.keyPress(KeyEvent.VK_TAB);
		 			robot.keyRelease(KeyEvent.VK_TAB);
		 			robot.keyPress(KeyEvent.VK_TAB);
		 			robot.keyRelease(KeyEvent.VK_TAB);

		 			robot.keyPress(KeyEvent.VK_ENTER);
		 			robot.keyRelease(KeyEvent.VK_ENTER);

		 			Thread.sleep(3000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_S);
		 			robot.keyRelease(KeyEvent.VK_S);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(2000);

		 			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\ReceiptsEmailTrigger.exe");

		 			Thread.sleep(4000);

		 			robot.keyPress(KeyEvent.VK_CONTROL);
		 			robot.keyPress(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_J);
		 			robot.keyRelease(KeyEvent.VK_CONTROL);

		 			Thread.sleep(2000);

		 			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

		 			int actOpenWindowsCount = getDriver().getWindowHandles().size();
		 			int expOpenWindowsCount = 4;

		 			System.out.println(
		 					"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

		 			getDriver().switchTo().window(newTabs.get(3)).close();
		 			Thread.sleep(1000);
		 			getDriver().switchTo().window(newTabs.get(2)).close();
		 			Thread.sleep(1000);
		 			getDriver().switchTo().window(newTabs.get(1));
		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
		 			mailDeleteButton.click();

		 			Thread.sleep(2000);

		 			Thread.sleep(2000);

		 			/*
		 			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
		 			 * gmailUserBtn)); gmailUserBtn.click();
		 			 */
		 			
		 			ClickUsingJs(gmailUserBtn);

		 			Thread.sleep(3000);
		 			getDriver().switchTo().frame(SignOutFrame);
		 			Thread.sleep(1000);
		 			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 			Thread.sleep(6000);
		 			
		 			getDriver().get(
							"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

					Thread.sleep(2000);
		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 			removeAccountBtn.click();

		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 			removeDeleteBtn.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 			yesRemoveBtn.click();

		 			Thread.sleep(2000);

		 			getDriver().switchTo().window(newTabs.get(1)).close();
		 			Thread.sleep(2000);
		 			getDriver().switchTo().window(newTabs.get(0));
		 			Thread.sleep(2000);

		 			String actAccount1PDF = getBaseDir() + "\\autoIt\\ExportFiles\\ReceiptsEmailTrigger.pdf";
		 			String expAccount1PDF = getBaseDir() + "\\autoIt\\ImportFiles\\ReceiptsEmailTrigger.pdf";

		 			PDFUtil pdfutil = new PDFUtil();

		 			boolean result = pdfutil.compare(actAccount1PDF, expAccount1PDF);
		 			Calendar cal = Calendar.getInstance();
		 			DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
		 			String currentDate = df.format(cal.getTime());

		 			String actAccount1Data = pdfutil.getText(actAccount1PDF);
		 			String expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("28/02/2025", currentDate);

		 			System.err.println(actAccount1Data);
		 			System.err.println(expAccount1Data);

		 			if (actAccount1Data.equalsIgnoreCase(expAccount1Data) ) {

		 				System.err.println("PDF file is as Expected");
		 				return true;
		 			} else {

		 				System.err.println("PDF file is not as Expected");
		 				return false;
		 			}

		 		} else {

		 			System.err.println("Subject : " + emailRow1Subject.getText());

		 			
		 			ClickUsingJs(gmailUserBtn);

		 			Thread.sleep(3000);
		 			getDriver().switchTo().frame(SignOutFrame);
		 			Thread.sleep(1000);
		 			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
		 			Thread.sleep(6000);
		 			
		 			getDriver().get(
							"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

					Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
		 			removeAccountBtn.click();

		 			Thread.sleep(1000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
		 			removeDeleteBtn.click();

		 			Thread.sleep(2000);

		 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
		 			yesRemoveBtn.click();

		 			Thread.sleep(2000);
		 			return false;
		 		}

		 	
		 	 
		  }
		 
		  
		  
		///Material Receipts Note Triggers
		  
		  @FindBy(xpath="//i[@class='icon-settings hiconright2']")
		  public static WebElement mrnSettingsBtn;
		  
		  @FindBy(xpath="(//*[@id='134217831'])[1]")
		  public static WebElement triggerBPAAmt;
		  
		  @FindBy(xpath="(//*[@id='134217832'])[1]")
		  public static WebElement triggerBPADAmt;
		  
		  @FindBy(xpath="(//*[@id='134217833'])[1]")
		  public static WebElement triggerDebitNoteAmount;
		  
		  @FindBy(xpath = "//*[@id='0_1_AdvanceFilter_']/table/tbody/tr/td[5]")
		   private static WebElement triggersValueToEnterTxt1;
		  
		  @FindBy(xpath = "//*[@id='17_1_AdvanceFilter_']/table/tbody/tr/td[5]/input")
		 	private static WebElement triggersValueTxt2;
		  

public boolean checkSaveTriggerinMRNVoucherbasedonBPAAmountwithEnableRaiseandOnSaveOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{ 
	 Thread.sleep(2000);
	 getDriver().navigate().refresh();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Material Receipt Notes");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();

		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mrnSettingsBtn));
		mrnSettingsBtn.click();
		Thread.sleep(4000);

	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();
		Thread.sleep(2000);
	 
	 
	 boolean flag=checkSavingTriggerinMRNVoucher("BPA Trigger On Save", triggerBPAAmt, "BPA Amount");
if(flag)
{
	 return true;
	 
}
else
{
	 return false;
}
}


@FindBy(xpath="//*[@id='id_header_4']")
public static WebElement mrn_VendorAcc;

@FindBy(xpath="//*[@id='id_header_268435460']")
public static WebElement mrn_WarehouseTxt;

@FindBy(xpath="//*[@id='id_header_268435459']")
public static WebElement mrn_Department;

@FindBy(xpath="(//i[@class='icon-new hiconright2'])[1]")
public static WebElement mnr_NewBtn;



@FindBy(xpath="//tbody[@id='id_header_268435460_table_body']/tr/td[2]")
public static List<WebElement> warehouseListCount;



@FindBy(xpath="//*[@id='id_footer_134217831']")
public static WebElement footer_BPAAmountTxt;

@FindBy(xpath="//*[@id='id_footer_134217832']")
public static WebElement footer_BPADAmountTxt;

@FindBy(xpath="//*[@id='id_footer_134217833']")
public static WebElement footer_DebitNoteAmountTxt;


public boolean checkSavingMRNVouchertoRaiseaTriggerVoucherBasedonBAPAmount() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	 
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mnr_NewBtn));
	 mnr_NewBtn.click();
	 Thread.sleep(6000);
	 
	 checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();

		vendorAccountTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actVendorAccountList = new ArrayList<String>();

		int count = vendorAccountListCount.size();

		System.err.println(count);

		for (int i = 0; i < count; i++) {
			String data = vendorAccountListCount.get(i).getText();

			if (data.equalsIgnoreCase("Vendor B")) {
				vendorAccountListCount.get(i).click();
				break;
			}
		}

		vendorAccountTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mrn_WarehouseTxt));
		mrn_WarehouseTxt.click();

		mrn_WarehouseTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actWarehouseList = new ArrayList<String>();

		int count1 = warehouseListCount.size();

		System.err.println(count);

		for (int i = 0; i < count; i++) {
			String data = warehouseListCount.get(i).getText();

			if (data.equalsIgnoreCase("HYDERABAD")) {
				warehouseListCount.get(i).click();
				break;
			}
		}

		mrn_WarehouseTxt.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actDepartmentList = new ArrayList<String>();

		int departmentCount = departmentListCount.size();

		System.err.println(departmentCount);

		for (int i = 0; i < departmentCount; i++) {
			String data = departmentListCount.get(i).getText();

			if (data.equalsIgnoreCase("INDIA")) {
				departmentListCount.get(i).click();
				break;
			}
		}

		departmentTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		
		int pvvGridItemListCount = pvvGridItemList.size();
		
		for (int i = 0; i < pvvGridItemListCount; i++) {
			
			String Item = pvvGridItemList.get(i).getText();
			
			if (Item.equalsIgnoreCase("BR COGS ITEM")) 
			{
				pvvGridItemList.get(i).click();
				break;
			}
		}
		enter_ItemTxt.sendKeys(Keys.TAB);	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		select1stRow_5thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.click();
		enter_Quantity.clear();
		enter_Quantity.sendKeys("2");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
		select1stRow_7thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		select1stRow_9thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys("MRN#V1");
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal = Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 2);

		String next2DaysDate = df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : " + next2DaysDate);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next2DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footer_BPAAmountTxt));
		footer_BPAAmountTxt.click();
		footer_BPAAmountTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		footer_BPAAmountTxt.sendKeys("25");
		Thread.sleep(1500);
		footer_BPAAmountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footer_BPADAmountTxt));
		footer_BPADAmountTxt.click();
		footer_BPADAmountTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		footer_BPADAmountTxt.sendKeys("50");
		Thread.sleep(1500);
		footer_BPADAmountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footer_DebitNoteAmountTxt));
		footer_DebitNoteAmountTxt.click();
		footer_DebitNoteAmountTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		footer_DebitNoteAmountTxt.sendKeys("25");
		Thread.sleep(1500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		Thread.sleep(2000);

		String expMessage1 = "Voucher saved successfully";

		String actMessage = checkValidationMessage(expMessage1);
		//String expMessage2 = ": 1";

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage1);

		if (actMessage.startsWith(expMessage1) ) {
			return true;
		} else {
			return false;
		}
	

}


public ArrayList<String> checkMRNVoucher(String DocNo) throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException 
{
	 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys(DocNo);
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);

		checkValidationMessage("Voucher loaded Successfully");
		Thread.sleep(3000);
		
		String actVendoracc=vendorAccountTxt.getAttribute("value");

		//String actWarehouseTxt=mrn_WarehouseTxt.getAttribute("value");
				
		String actDepartmentTxt=departmentTxt.getAttribute("value");
		
		//String actR1Item = select1stRow_2ndColumn.getText();
		String actR1PurchaseAcc = select1stRow_2ndColumn.getText();
		
		String actR1Quantity = select1stRow_5thColumn.getText();
		

		String actR1Rate = select1stRow_7thColumn.getText();
		

		String actBPAAmount = footer_BPAAmountTxt.getAttribute("value");
		
		String actBPADAmount = footer_BPADAmountTxt.getAttribute("value");
		
		String actDebitNoteAmount = footer_DebitNoteAmountTxt.getAttribute("value");
		

		ArrayList<String> actArray = new ArrayList<String>();
		actArray.add(actVendoracc);
		//actArray.add(actWarehouseTxt);
		actArray.add(actDepartmentTxt);
		actArray.add(actR1PurchaseAcc);
		//actArray.add(actR1Quantity);
		//actArray.add(actR1Rate);
		actArray.add(actBPAAmount);
		actArray.add(actBPADAmount);
		actArray.add(actDebitNoteAmount);
		
		
		return actArray;

}

@FindBy(xpath="//a[@id='2032']/span[text()='Journal Entries']")
public static WebElement journalEntriesVoucher;

public boolean checkJournalEntriesRaisedVoucherafterSavingTriggerMRNVOucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	 ArrayList<String> actArry = new ArrayList<String>();
	 actArry=checkMRNVoucher("1");
	 System.out.println(actArry.toString());
	 
	 Thread.sleep(6000);
	 
	 
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
	 
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsJournalsMenu));
		financialsTransactionsJournalsMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(journalEntriesVoucher));
		journalEntriesVoucher.click();
		Thread.sleep(2000);

		
		System.out.println("Number of Vouchers raised after trigger saved Before Authorization  "  +((voucherGridDocNo.size())-1));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for (int i = 0; i < voucherGridDocNoCount; i++) {
			String data = "1";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(4000);
		
		String actDepartmentTxt1=departmentTxt.getAttribute("value");
		String expDepartmentTxt1=actArry.get(1);
		
		String actDebitAcc1 = select1stRow_1stColumn.getText();
		String expDebitAcc1=actArry.get(0);
		
		String actCreditAcc1 = select1stRow_2ndColumn.getText();
		String expCreditAcc1=actArry.get(2);

		String actAmount1 = select1stRow_3rdColumn.getText();
		String expBPAAmount=actArry.get(3);
		
		
		
		System.err.println("****************************Trigger Voucher Based on BPA Amount*************************");
		System.err.println("Actual Department 		"	+	actDepartmentTxt1	+		"Expected         	"+expDepartmentTxt1);
		System.err.println("Actual Debit Account 	"	+	actDebitAcc1		+		"Expected        	"+expDebitAcc1);
		System.err.println("Actual Credit Account	"	+	actCreditAcc1		+		"Expected        	"+expCreditAcc1);
		System.err.println("Actual Amount 			"	+	actAmount1			+		"Expected 	        "+expBPAAmount);
		
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("2");
		Thread.sleep(1500);
		documentNumberTxt.sendKeys(Keys.TAB);
		Thread.sleep(6000);
		String actDepartmentTxt2=departmentTxt.getAttribute("value");
		String expDepartmentTxt2=actArry.get(1);
		
		String actDebitAcc2 = select1stRow_1stColumn.getText();
		String expDebitAcc2=actArry.get(0);
		
		String actCreditAcc2 = select1stRow_2ndColumn.getText();
		String expCreditAcc2=actArry.get(2);

		String actAmount2 = select1stRow_3rdColumn.getText();
		String expBPADAmount=actArry.get(4);
		
		
		
		System.err.println("****************************Trigger Voucher Based on BPAD Amount*************************");
		System.err.println("Actual Department 		"	+	actDepartmentTxt2	+		"Expected         	"+expDepartmentTxt2);
		System.err.println("Actual Debit Account 	"	+	actDebitAcc2		+		"Expected        	"+expDebitAcc2);
		System.err.println("Actual Credit Account	"	+	actCreditAcc2		+		"Expected        	"+expCreditAcc2);
		System.err.println("Actual Amount 			"	+	actAmount2			+		"Expected 	        "+expBPADAmount);
		
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("3");
		Thread.sleep(1500);
		documentNumberTxt.sendKeys(Keys.TAB);
		Thread.sleep(6000);
		
		String actDepartmentTxt3=departmentTxt.getAttribute("value");
		String expDepartmentTxt3=actArry.get(1);
		
		String actDebitAcc3 = select1stRow_1stColumn.getText();
		String expDebitAcc3=actArry.get(0);
		
		String actCreditAcc3 = select1stRow_2ndColumn.getText();
		String expCreditAcc3=actArry.get(2);

		String actAmount3 = select1stRow_3rdColumn.getText();
		String expDebitNoteAmount=actArry.get(5);
		
		
		
		System.err.println("****************************Trigger Voucher Based on Debit Note Amount*************************");
		System.err.println("Actual Department 		"	+	actDepartmentTxt3		+		"Expected         	"+expDepartmentTxt3);
		System.err.println("Actual Debit Account 	"	+	actDebitAcc3			+		"Expected        	"+expDebitAcc3);
		System.err.println("Actual Credit Account	"	+	actCreditAcc3			+		"Expected        	"+expCreditAcc3);
		System.err.println("Actual Amount 			"	+	actAmount3				+		"Expected 	        "+expDebitNoteAmount);
		
		
		/*
		 * ArrayList<String> expArray = new ArrayList<String>();
		 * 
		 * expArray.add(actR1DebitAcc); expArray.add(actR1DepatmentTxt);
		 * expArray.add(actR1CreditAcc); expArray.add(actR1Amount);
		 * 
		 * 
		 * String actRowsList=actArry.toString(); String expRowList=expArray.toString();
		 * 
		 * System.err.println(actRowsList); System.err.println(expRowList);
		 */
		
		
		if(actDepartmentTxt1.equalsIgnoreCase(expDepartmentTxt1) && actDebitAcc1.equalsIgnoreCase(expDebitAcc1) && actCreditAcc1.equalsIgnoreCase(expCreditAcc1)
				&& actAmount1.equalsIgnoreCase(expBPAAmount)&&actDepartmentTxt2.equalsIgnoreCase(expDepartmentTxt2) && actDebitAcc2.equalsIgnoreCase(expDebitAcc2) && actCreditAcc2.equalsIgnoreCase(expCreditAcc2)
				&& actAmount2.equalsIgnoreCase(expBPADAmount) && actDepartmentTxt3.equalsIgnoreCase(expDepartmentTxt3) && actDebitAcc3.equalsIgnoreCase(expDebitAcc3) && actCreditAcc3.equalsIgnoreCase(expCreditAcc3)
				&& actAmount3.equalsIgnoreCase(expDebitNoteAmount))
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
}



public boolean checkSaveTriggerinMRNVoucherbasedonBPAAmountwithEnableRaiseandOnAuthorizeOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	 boolean flag=checkSavingTriggerinMRNVoucher("BPA Trigger On Authorization",triggerBPAAmt,"BPA Amount");
	 if(flag)
	 {
		 return true;
		 
	 }
	 else
	 {
		 return false;
	 }
}


public boolean checkSaveTriggerinMRNVoucherbasedonBPADAmountwithEnableRaiseandOnSaveOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	 boolean flag=checkSavingTriggerinMRNVoucher("BPAD Trigger On Save",triggerBPADAmt,"BPAD Amount");
	 if(flag)
	 {
		 return true;
		 
	 }
	 else
	 {
		 return false;
	 }
}



public boolean checkSaveTriggerinMRNVoucherbasedonBPADAmountwithEnableRaiseandOnAuthorizeOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	 boolean flag=checkSavingTriggerinMRNVoucher("BPAD Trigger On Authorization",triggerBPADAmt,"BPAD Amount");
	 if(flag)
	 {
		 return true;
		 
	 }
	 else
	 {
		 return false;
	 }
}


public boolean checkSaveTriggerinMRNVoucherbasedonDebitNoteAmountwithEnableRaiseandOnSaveOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	 boolean flag=checkSavingTriggerinMRNVoucher("Debit Note  Trigger On Save",triggerDebitNoteAmount,"Debit Note amount");
	 if(flag)
	 {
		 return true;
		 
	 }
	 else
	 {
		 return false;
	 }
}



public boolean checkSaveTriggerinMRNVoucherbasedonDebitNoteAmountwithEnableRaiseandOnAuthorizeOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	 boolean flag=checkSavingTriggerinMRNVoucher("Debit Note Trigger On Authorization",triggerDebitNoteAmount,"Debit Note amount");
	 
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		 Thread.sleep(2000);

		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);

		
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn)); 
		  settings_closeBtn.click();
		  Thread.sleep(2000);
	 
	 if(flag)
	 {
		 return true;
		 
	 }
	 else
	 {
		 return false;
	 }
}


@FindBy(xpath = "//td[@id='triggers_docTable_col_2-1']")
private static WebElement trigger2ndRow1stCol;

@FindBy(xpath = "//*[@id='rdbOnAuthorization']")
private static WebElement triggersOnAuthRadioBtn;

@FindBy(xpath = "//td[@id='triggers_docTable_col_2-2']")
private static WebElement trigger2ndRow2ndCol;

@FindBy(xpath = "//td[@id='triggers_docTable_col_3-1']")
private static WebElement trigger3rdRow1stCol;

@FindBy(xpath = "//td[@id='triggers_docTable_col_3-2']")
private static WebElement trigger3rdRow2ndCol;

@FindBy(xpath = "//td[@id='triggers_docTable_col_4-1']")
private static WebElement trigger4thRow1stCol;

@FindBy(xpath = "//td[@id='triggers_docTable_col_4-2']")
private static WebElement trigger4thRow2ndCol;

@FindBy(xpath = "//td[@id='triggers_docTable_col_5-1']")
private static WebElement trigger5thRow1stCol;

@FindBy(xpath = "//td[@id='triggers_docTable_col_5-2']")
private static WebElement trigger5thRow2ndCol;


public boolean checkSavingTriggerinMRNVoucher(String triggerName, WebElement sname, String AmtTxt) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	    Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT, Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys(triggerName);
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(6000);
		triggersConjuctionDrpdwn.click();
		Thread.sleep(2000);
		triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);
		triggersConjuctionDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
		triggersSelectFieldTxt.click();

		Thread.sleep(3000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sname));
		sname.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
		triggersSelectOperdrpdwn.click();
		triggersSelectOperdrpdwn.sendKeys("Greater than");

		Thread.sleep(5000);
		triggersSelectOperdrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
		//triggersCompareWithdrpdwn.click();
		triggersCompareWithdrpdwn.sendKeys("value");
		Thread.sleep(2000);
		triggersCompareWithdrpdwn.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		triggersValueTxt2.sendKeys("0");
		Thread.sleep(1500);
		triggersValueTxt2.sendKeys(Keys.TAB);

		Thread.sleep(1500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
		if(editLayoutTriggerRaiseDoc.isSelected()==false)
		{
			editLayoutTriggerRaiseDocSelected.click();
		}
		
		
		switch(triggerName)
		{
		case "BPA Trigger On Authorization":
		case "BPAD Trigger On Authorization":
		case "Debit Note Trigger On Authorization":
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersOnAuthRadioBtn));
				triggersOnAuthRadioBtn.click();
				break;
		default: break;
		
		}
		
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
		Select s = new Select(editLayoutTriggersAllVouchers);
		s.selectByVisibleText("Journal Entries");
		Thread.sleep(3000);

		getAction().moveToElement(trigger1stRow1stCol).build().perform();

		ClickUsingJs(trigger1stRow1stCol);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectFieldDropdown));

		Thread.sleep(1000);
		Select s1 = new Select(triggerSelectFieldDropdown);
		s1.selectByVisibleText("Date");
		Thread.sleep(3000);
		triggerSelectFieldDropdown.sendKeys(Keys.TAB);

		ClickUsingJs(trigger1stRow2ndCol);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectMapDropdown));
		Select s2 = new Select(triggerSelectMapDropdown);
		s2.selectByVisibleText("Date");
		Thread.sleep(2000);

		ClickUsingJs(trigger2ndRow1stCol);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectFieldDropdown));

		Thread.sleep(1000);
		Select s3 = new Select(triggerSelectFieldDropdown);
		s3.selectByVisibleText("VendorAC");
		Thread.sleep(3000);
		triggerSelectFieldDropdown.sendKeys(Keys.TAB);

		ClickUsingJs(trigger2ndRow2ndCol);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectMapDropdown));
		Select s4 = new Select(triggerSelectMapDropdown);
		s4.selectByVisibleText("DrAccount");
		Thread.sleep(2000);
		
		ClickUsingJs(trigger3rdRow1stCol);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectFieldDropdown));

		Thread.sleep(1000);
		Select s5 = new Select(triggerSelectFieldDropdown);
		s5.selectByVisibleText("PurchaseAC");
		Thread.sleep(3000);
		triggerSelectFieldDropdown.sendKeys(Keys.TAB);

		ClickUsingJs(trigger3rdRow2ndCol);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectMapDropdown));
		Select s6 = new Select(triggerSelectMapDropdown);
		s6.selectByVisibleText("CrAccount");
		Thread.sleep(2000);
		
		ClickUsingJs(trigger4thRow1stCol);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectFieldDropdown));

		Thread.sleep(1000);
		Select s7 = new Select(triggerSelectFieldDropdown);
		s7.selectByVisibleText("Department");
		Thread.sleep(3000);
		triggerSelectFieldDropdown.sendKeys(Keys.TAB);

		ClickUsingJs(trigger4thRow2ndCol);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectMapDropdown));
		Select s8 = new Select(triggerSelectMapDropdown);
		s8.selectByVisibleText("Department");
		Thread.sleep(2000);
		
		
		ClickUsingJs(trigger5thRow1stCol);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectFieldDropdown));

		Thread.sleep(1000);
		Select s9 = new Select(triggerSelectFieldDropdown);
		s9.selectByVisibleText(AmtTxt);
		Thread.sleep(3000);
		triggerSelectFieldDropdown.sendKeys(Keys.TAB);

		ClickUsingJs(trigger5thRow2ndCol);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectMapDropdown));
		Select s10 = new Select(triggerSelectMapDropdown);
		s10.selectByVisibleText("Amount");
		Thread.sleep(2000);
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		Thread.sleep(2500);

		String expSaveMessage = "Data saved successfully";
		String actSaveMessage = checkValidationMessage(expSaveMessage);

		
		 

		if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
			System.out.println("Test Pass : Trigger  Added Successfully");

			return true;
		} else {
			System.out.println("Test Fail : Trigger NOT Added Successfully");

			return false;
		}
}






@FindBy(xpath = "//*[@id='1280']")
	private static WebElement transactionAuthorizationsMRN;

	@FindBy(xpath = "//*[@id='masterChild_1280']")
	private static WebElement transactionAuthorizationTextMRN;



public static boolean checkAuthorizationInMRNVoucherWithUserAllOptionsST()throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
	 
	  Thread.sleep(4000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(noOfAlertsDisplay));
		noOfAlertsDisplaySelcted.click();
		 Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(alertTransactions));
		alertTransactions.click();
		 Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(transactionAuthorizationsMRN));
		transactionAuthorizationsMRN.click();
		 Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(transactionAuthorizationTextMRN));
		transactionAuthorizationTextMRN.click();
		 Thread.sleep(2000);
		checkValidationMessage("Voucher loaded Successfully");

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(POauthorizeBtn));
		POauthorizeBtn.click();
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(openingStocksNewHomeRow1AuthorizationStatus));

		String actStatus = openingStocksNewHomeRow1AuthorizationStatus.getText();
		String expStatus = "Authorized";

		System.out.println("Authorization Status In home Page 	 :  " + actStatus + " Value Expected : " + expStatus);

		if (actStatus.equalsIgnoreCase(expStatus)) {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();

			return true;
		} 
		else 
		{
			return false;
		}
		
}


public boolean checkJournalEntriesVoucherAfterAuthorizationMRNVoucher() throws InterruptedException
{
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
	 
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsJournalsMenu));
		financialsTransactionsJournalsMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(journalEntriesVoucher));
		journalEntriesVoucher.click();
		Thread.sleep(2000);

		
		System.out.println("Number of Vouchers raised after trigger saved After Authorization  "  +((voucherGridDocNo.size())-1));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for (int i = 0; i < voucherGridDocNoCount; i++) {
			String data = "4";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(2000);
		
		String actDepartmentTxt1=departmentTxt.getAttribute("value");
		String expDepartmentTxt1="INDIA";
		
		String actDebitAcc1 = select1stRow_1stColumn.getText();
		String expDebitAcc1="Vendor B";
		
		String actCreditAcc1 = select1stRow_2ndColumn.getText();
		String expCreditAcc1="BR COGS ACC INV";

		String actAmount1 = select1stRow_3rdColumn.getText();
		String expBPAAmount="25.00";
		
		
		
		System.err.println("****************************Trigger Voucher Based on BPA Amount*************************");
		System.err.println("Actual Department 		"	+	actDepartmentTxt1	+		"Expected         	"+expDepartmentTxt1);
		System.err.println("Actual Debit Account 	"	+	actDebitAcc1		+		"Expected        	"+expDebitAcc1);
		System.err.println("Actual Credit Account	"	+	actCreditAcc1		+		"Expected        	"+expCreditAcc1);
		System.err.println("Actual Amount 			"	+	actAmount1			+		"Expected 	        "+expBPAAmount);
		
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("5");
		Thread.sleep(1500);
		documentNumberTxt.sendKeys(Keys.TAB);
		Thread.sleep(3000);
		String actDepartmentTxt2=departmentTxt.getAttribute("value");
		String expDepartmentTxt2="INDIA";
		
		String actDebitAcc2 = select1stRow_1stColumn.getText();
		String expDebitAcc2="Vendor B";
		
		String actCreditAcc2 = select1stRow_2ndColumn.getText();
		String expCreditAcc2="BR COGS ACC INV";

		String actAmount2 = select1stRow_3rdColumn.getText();
		String expBPADAmount="50.00";
		
		
		
		System.err.println("****************************Trigger Voucher Based on BPAD Amount*************************");
		System.err.println("Actual Department 		"	+	actDepartmentTxt2	+		"Expected         	"+expDepartmentTxt2);
		System.err.println("Actual Debit Account 	"	+	actDebitAcc2		+		"Expected        	"+expDebitAcc2);
		System.err.println("Actual Credit Account	"	+	actCreditAcc2		+		"Expected        	"+expCreditAcc2);
		System.err.println("Actual Amount 			"	+	actAmount2			+		"Expected 	        "+expBPADAmount);
		
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("6");
		Thread.sleep(1500);
		documentNumberTxt.sendKeys(Keys.TAB);
		Thread.sleep(3000);
		
		String actDepartmentTxt3=departmentTxt.getAttribute("value");
		String expDepartmentTxt3="INDIA";
		
		String actDebitAcc3 = select1stRow_1stColumn.getText();
		String expDebitAcc3="Vendor B";
		
		String actCreditAcc3 = select1stRow_2ndColumn.getText();
		String expCreditAcc3="BR COGS ACC INV";

		String actAmount3 = select1stRow_3rdColumn.getText();
		String expDebitNoteAmount="25.00";
		
		
		
		System.err.println("****************************Trigger Voucher Based on Debit Note Amount*************************");
		System.err.println("Actual Department 		"	+	actDepartmentTxt3		+		"Expected         	"+expDepartmentTxt3);
		System.err.println("Actual Debit Account 	"	+	actDebitAcc3			+		"Expected        	"+expDebitAcc3);
		System.err.println("Actual Credit Account	"	+	actCreditAcc3			+		"Expected        	"+expCreditAcc3);
		System.err.println("Actual Amount 			"	+	actAmount3				+		"Expected 	        "+expDebitNoteAmount);
		
		
		/*
		 * ArrayList<String> expArray = new ArrayList<String>();
		 * 
		 * expArray.add(actR1DebitAcc); expArray.add(actR1DepatmentTxt);
		 * expArray.add(actR1CreditAcc); expArray.add(actR1Amount);
		 * 
		 * 
		 * String actRowsList=actArry.toString(); String expRowList=expArray.toString();
		 * 
		 * System.err.println(actRowsList); System.err.println(expRowList);
		 */
		
		
		if(actDepartmentTxt1.equalsIgnoreCase(expDepartmentTxt1) && actDebitAcc1.equalsIgnoreCase(expDebitAcc1) && actCreditAcc1.equalsIgnoreCase(expCreditAcc1)
				&& actAmount1.equalsIgnoreCase(expBPAAmount)&&actDepartmentTxt2.equalsIgnoreCase(expDepartmentTxt2) && actDebitAcc2.equalsIgnoreCase(expDebitAcc2) && actCreditAcc2.equalsIgnoreCase(expCreditAcc2)
				&& actAmount2.equalsIgnoreCase(expBPADAmount) && actDepartmentTxt3.equalsIgnoreCase(expDepartmentTxt3) && actDebitAcc3.equalsIgnoreCase(expDebitAcc3) && actCreditAcc3.equalsIgnoreCase(expCreditAcc3)
				&& actAmount3.equalsIgnoreCase(expDebitNoteAmount))
		{
			return true;
		}
		else
		{
			return false;
		}
}
		  


public boolean checkEditingMRNVoucherAfterRaiseDocumentOnSave() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	 
	 
	 Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Material Receipt Notes");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();

		Thread.sleep(4000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mnr_NewBtn));
		 mnr_NewBtn.click();
		 Thread.sleep(2000);
		 
		 checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("1");
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(4000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(footer_BPAAmountTxt));
		footer_BPAAmountTxt.click();
		footer_BPAAmountTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		Thread.sleep(1500);
		footer_BPAAmountTxt.sendKeys("100");
		footer_BPAAmountTxt.sendKeys(Keys.TAB);
		
	
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
		billRefNewReferenceTxt.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		Thread.sleep(2000);

		String expMessage1 = "Voucher saved successfully";

		String actMessage = checkValidationMessage(expMessage1);
		//String expMessage2 = ": 1";

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage1);

		if (actMessage.startsWith(expMessage1) ) {
			return true;
		} else {
			return false;
		}
		
		
}


public boolean checkJournalEntriesVoucherAfterEditingMRNVoucher() throws InterruptedException
{
	 Thread.sleep(4000);
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
	 
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsJournalsMenu));
		financialsTransactionsJournalsMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(journalEntriesVoucher));
		journalEntriesVoucher.click();
		Thread.sleep(4000);

		
		System.out.println("Number of Vouchers raised after trigger saved Before Authorization  "  +((voucherGridDocNo.size())-1));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for (int i = 0; i < voucherGridDocNoCount; i++) {
			String data = "1";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(2000);
		
	
		String actAmount1 = select1stRow_3rdColumn.getText();
		String expBPAAmount="100.00";
		
		
		
		System.err.println("****************************Trigger Voucher Based on BPA Amount*************************");
		
		System.err.println("Actual Amount 			"	+	actAmount1			+		"Expected 	        "+expBPAAmount);
		
		Thread.sleep(2000);
		
		if(actAmount1.equalsIgnoreCase(expBPAAmount))
		{
			return true;
		}
		else
		{
			return false;
		}
	 
}



///Creating view

@FindBy(xpath="//a[@id='btnCreateView']//i[@class='icon-new hiconright2']")
	private static WebElement  createViewBtn;

	@FindBy(xpath="//*[@id='txtViewName']")
	private static WebElement  viewNameTxtInCreateView;

	@FindBy(xpath="//*[@id='cmbDataSet']")
	private static WebElement  dataSetInCreateView;

	@FindBy(xpath="//*[@id='cmbViewType']")
	private static WebElement  viewTypeInCreateView;

	@FindBy(xpath="//*[@id='cmbAuthorization']")
	private static WebElement  authorizationInCreateView;

	@FindBy(xpath="//*[@id='cmbCheckStatus']")
	private static WebElement  checkSatusInCreateView;

	@FindBy(xpath="//*[@id='cmbQC']")
	private static WebElement  qcInCreateView;

	@FindBy(xpath="//*[@id='Suspendstatus']")
	private static WebElement  suspendStatusInCreateView;

	@FindBy(xpath="//*[@id='a']")
	private static WebElement  filterOptionInCreateView;

	@FindBy(xpath="//*[@id='dvCreateEditView']//*[@value='Save']")
	private static WebElement  saveOptionInCreateView;

	@FindBy(xpath="//*[@id='id_transaction_viewcontainer']//li//*[contains(text(),'Department')]")
	private static WebElement createViewNameDisplayInHomePage;
	
	
	
	//Filter Fields
	@FindBy(xpath="//*[@id='2064_1_AdvanceFilter_']/table/tbody/tr/td[1]/select")
	public static WebElement viewFilter_WhereCondition;
	
		
	@FindBy(xpath="//*[@id='2064_1_AdvanceFilter_']/table/tbody/tr/td[2]/input")
	public static WebElement viewFilter_selectField;
	
	
	
	@FindBy(xpath="//*[@id='2064_1_AdvanceFilter_']/table/tbody/tr/td[3]/select")
	public static WebElement viewFilter_selectOperator;
	
	
	@FindBy(xpath="//*[@id='2064_1_AdvanceFilter_']/table/tbody/tr/td[4]/select")
	public static WebElement viewFilter_CompareWith;
	
	
	@FindBy(xpath="//*[@id='2064_1_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	public static WebElement viewFilter_ValueField;
	
	@FindBy(xpath="//*[@id='advancefilter_master_2064_1_']")
	public static WebElement viewFilter_ValueFieldTxt;
	
		
	@FindBy(xpath="(//a[contains(text(),'Department')]/span)[1]")
	public static WebElement viewFilter_DepartmentExpandBtn;
	
	@FindBy(xpath="(//a[contains(text(),'Department')]//following::ul//li/a[contains(text(),'Name')])[1]")
	public static WebElement viewFilter_DepartmentName;
	
	@FindBy(xpath="(//*[@id='4']/span)[1]")
	public static WebElement viewFilter_VendorAcExpandBtn;
	
	@FindBy(xpath="(//*[@id='5002'])[1]")
	public static WebElement viewFilter_VendorName;
	
	@FindBy(xpath="(//*[@id='536870916']/span)[1]")
	public static WebElement viewFilter_WarehouseExpandBtn;
	
	@FindBy(xpath="(//*[@id='5058'])[1]")
	public static WebElement viewFilter_WarehouseName;
	
	
	@FindBy(xpath="//p[contains(text(),'Modified by')]")
	public static WebElement Cust_ModifiedByCol;
	
	@FindBy(xpath="//span[contains(text(),'Transaction')]/span")
	public static WebElement cs_TransactionExpBtn;
	
	@FindBy(xpath="//..//span[contains(text(),'Transaction Fields')]/i")
	public static WebElement cs_TrnsactionFieldsExpBtn;
	
	@FindBy(xpath="//..//span[contains(text(),'Extra Fields')]/i")
	public static WebElement cs_ExtraFieldsExpBtn;
	
	@FindBy(xpath="(//../../span[contains(text(),'Warehouse')]//i)[1]")
	public static WebElement cs_TransactionWarehouseFieldExpBtn;
	
	@FindBy(xpath="((//../../span[contains(text(),'Warehouse')]//i)[1]//following::ul//li)[1][contains(text(),'Name')]")
	public static WebElement cs_TransactionWarehouseNameField;
	
	@FindBy(xpath="(//../../span[contains(text(),'Department')]//i)[1]")
	public static WebElement cs_TransactionDepartmentFieldExpBtn;
	
	@FindBy(xpath="((//../../span[contains(text(),'Department')]//i)[1]//following::ul//li)[1][contains(text(),'Name')]")
	public static WebElement cs_TransactionDepartmentNameField;
	
	
	public boolean checkSavingViewinPurchaseVoucherVATVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
		purchaseVouchersVat.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(1000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewBtn));
		createViewBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewNameTxtInCreateView));
		viewNameTxtInCreateView.click();
		viewNameTxtInCreateView.sendKeys("Department INDIA");
		Thread.sleep(2000);
		viewNameTxtInCreateView.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTypeInCreateView));
		Select s=new Select(viewTypeInCreateView);
		s.selectByVisibleText("Detailed");
		Thread.sleep(2000);
		viewTypeInCreateView.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizationInCreateView));
		Select s2=new Select(authorizationInCreateView);
		s2.selectByVisibleText("All");
		Thread.sleep(2000);
		authorizationInCreateView.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(checkSatusInCreateView));
		Select s3=new Select(checkSatusInCreateView);
		s3.selectByVisibleText("Both");
		Thread.sleep(2000);
		checkSatusInCreateView.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filterOptionInCreateView));
		filterOptionInCreateView.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_WhereCondition));
		//viewFilter_WhereCondition.click();
		Select s4=new Select(viewFilter_WhereCondition);
		s4.selectByVisibleText("Where");
		Thread.sleep(1000);
		//viewFilter_WhereCondition.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_selectField));
		viewFilter_selectField.click();
		Thread.sleep(2000);
	
		ScrollIntoView(viewFilter_DepartmentExpandBtn);
	//	getAction().moveToElement(viewFilter_DepartmentExpandBtn).build().perform();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_DepartmentExpandBtn));
		viewFilter_DepartmentExpandBtn.click();
		Thread.sleep(3000);
		getAction().moveToElement(viewFilter_DepartmentName).build().perform();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_DepartmentName));
		viewFilter_DepartmentName.click();
	
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_selectOperator));
		viewFilter_selectOperator.click();
		Select s5= new Select(viewFilter_selectOperator);
		s5.selectByVisibleText("Equal to");
		Thread.sleep(2000);
		//viewFilter_selectOperator.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_CompareWith));
		viewFilter_CompareWith.click();
		Thread.sleep(2000);
		Select s6= new Select(viewFilter_CompareWith);
		s6.selectByVisibleText("value");
		Thread.sleep(2000);
		//viewFilter_CompareWith.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_ValueFieldTxt));
		viewFilter_ValueFieldTxt.click();
		viewFilter_ValueFieldTxt.sendKeys("DUBAI");
		Thread.sleep(2000);
		viewFilter_ValueFieldTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveOptionInCreateView));
		saveOptionInCreateView.click();

		String expSaveViewMessage = "View saved successfully";

		String actSaveViewMessage = checkValidationMessage(expSaveViewMessage);
		Thread.sleep(2000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		*/
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeBtn));
		customizeBtn.click();

		Thread.sleep(4000);
		
		click(Cust_ModifiedByCol);
		
		click(cs_TransactionExpBtn);
		
		click(cs_ExtraFieldsExpBtn);
		
		click(cs_TransactionDepartmentFieldExpBtn);
		
		
		getAction().doubleClick(cs_TransactionDepartmentNameField).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(custSaveBtn));
		custSaveBtn.click();
		
		 Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		 Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(moreViewsBtn));
		 moreViewsBtn.click();
		 Thread.sleep(2000);

		String actcreateViewNameDisplayInHomePage = createdDepartmentViewinHomePage.getText();
		String expcreateViewNameDisplayInHomePage = "Department INDIA";

		System.out.println("View Display In Home page Value Actual    : " + actcreateViewNameDisplayInHomePage   +     "	 Value Expected : " + expcreateViewNameDisplayInHomePage);

		
		
		if(actSaveViewMessage.equalsIgnoreCase(expSaveViewMessage) && actcreateViewNameDisplayInHomePage.equalsIgnoreCase(expcreateViewNameDisplayInHomePage) )
		{
			return true;
		}
		else
		{
			return false;
		} 
		
	}

	

	@FindBy(xpath="//*[@id='id_transaction_viewcontainer']//li//a[contains(text(),'Department INDIA')]")
	public static WebElement createdDepartmentViewinHomePage;
	
	@FindBy(xpath="//*[@id='id_transaction_viewcontainer']//li//*[contains(text(),'VendorA')]")
	public static WebElement createdVendorViewinHomePage;
	
	@FindBy(xpath="//*[@id='id_transaction_viewcontainer']//li//*[contains(text(),'Warehouse HYD')]")
	public static WebElement createdWarehouseViewinHomePage;
	
	@FindBy(xpath="//li[@id='id_transactionentry_close']")
	public static WebElement newCloseBtn;
	
	@FindBy(xpath="//*[@id='tblBodyTransRender']/tr/td[8]")
	public static List<WebElement>  voucherHomePageDepartmentColList;
	
	
	
	public boolean checkCreatedDepartmentViewinPurchaseVoucherVAT() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdDepartmentViewinHomePage));
		createdDepartmentViewinHomePage.click();
		Thread.sleep(2000);
		

		int actvoucherGridDocNoCount = voucherGridDocNo.size()-1;
		System.out.println("Number of Vouchers in Department View	"	+	actvoucherGridDocNoCount);
		int expvoucherGridDocNoCount=8;
		
		ArrayList<String>Departmentarray=new ArrayList<String>();
		
		for(int i=0;i<actvoucherGridDocNoCount;i++)
		{
			Departmentarray.add(voucherHomePageDepartmentColList.get(i).getText());
		}

		
		Thread.sleep(2000);
		
		String actDeptLsit=Departmentarray.toString();
		String expDeptList="";
		
		System.out.println("Actual Department List  " 	+	actDeptLsit);
		
		System.out.println("Expected Department List  " 	+	expDeptList);
		
		
		ArrayList<String>DocumentNum=new ArrayList<String>();
		
		for(int i=0;i<actvoucherGridDocNoCount;i++)
		{
			DocumentNum.add(voucherGridDocNo.get(i).getText());
		}

		
		Thread.sleep(2000);
		
		String actDocNumList=DocumentNum.toString();
		String expDocNumList="[SU/IND/TEXT8, SU/IND/TEXT7, SU/IND/TEXT6, SU/IND/TEXT5, SU/IND/TEXT4, SU/IND/TEXT3, SU/IND/TEXT2, SU/IND/TEXT1]";
		
		System.out.println("Actual Document List:		"	+	actDocNumList);
		System.out.println("Expected Document List:		"+	expDocNumList);

		for (int i = 0; i < actvoucherGridDocNoCount; i++) {
			String data = "SU/IND/TEXT8";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNum=documentNumberTxt.getAttribute("value");
		String expDocNum="SU/IND/TEXT8";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDeptTxt=departmentTxt.getAttribute("value");
		String expDeptTxt="INDIA";
		
		System.out.println("Actual Document Number:		"	+	actDocNum	+"Expected:		"	+	expDocNum);
		System.out.println("Actual Department:		"	+	actDeptTxt	+"Expected:		"	+	expDeptTxt);
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCloseBtn));
		newCloseBtn.click();
		
		
		if(actvoucherGridDocNoCount==expvoucherGridDocNoCount && actDocNum.equalsIgnoreCase(expDocNum)
				&& actDocNumList.equalsIgnoreCase(expDocNumList))
		{
		
			return true;
		}
		else
		{
			return false;
		}
		
		
		
	}
	
public boolean checkSavingVendorViewinPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	 Thread.sleep(2000);
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewBtn));
		createViewBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewNameTxtInCreateView));
		viewNameTxtInCreateView.click();
		viewNameTxtInCreateView.sendKeys("VendorA");
		Thread.sleep(2000);
		viewNameTxtInCreateView.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTypeInCreateView));
		Select s=new Select(viewTypeInCreateView);
		s.selectByVisibleText("Detailed");
		Thread.sleep(2000);
		viewTypeInCreateView.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizationInCreateView));
		Select s2=new Select(authorizationInCreateView);
		s2.selectByVisibleText("All");
		Thread.sleep(2000);
		authorizationInCreateView.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(checkSatusInCreateView));
		Select s3=new Select(checkSatusInCreateView);
		s3.selectByVisibleText("Both");
		Thread.sleep(2000);
		checkSatusInCreateView.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filterOptionInCreateView));
		filterOptionInCreateView.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_WhereCondition));
		//viewFilter_WhereCondition.click();
		Select s4=new Select(viewFilter_WhereCondition);
		s4.selectByVisibleText("Where");
		Thread.sleep(1000);
		viewFilter_WhereCondition.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_selectField));
		viewFilter_selectField.click();
		Thread.sleep(2000);
		
		getAction().moveToElement(viewFilter_VendorAcExpandBtn).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_VendorAcExpandBtn));
		viewFilter_VendorAcExpandBtn.click();
		Thread.sleep(3000);
		getAction().moveToElement(viewFilter_VendorName).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_VendorName));
		viewFilter_VendorName.click();
	
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_selectOperator));
		viewFilter_selectOperator.click();
		Select s5= new Select(viewFilter_selectOperator);
		s5.selectByVisibleText("Not equal to");
		Thread.sleep(2000);
		//viewFilter_selectOperator.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_CompareWith));
		viewFilter_CompareWith.click();
		Select s6= new Select(viewFilter_CompareWith);
		s6.selectByVisibleText("value");
		Thread.sleep(2000);
		//viewFilter_CompareWith.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_ValueFieldTxt));
		viewFilter_ValueFieldTxt.click();
		viewFilter_ValueFieldTxt.sendKeys("Vendor B");
		Thread.sleep(2000);
		viewFilter_ValueFieldTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveOptionInCreateView));
		saveOptionInCreateView.click();

		String expSaveViewMessage = "View saved successfully";

		String actSaveViewMessage = checkValidationMessage(expSaveViewMessage);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(moreViewsBtn));
		moreViewsBtn.click();
		
		String actcreateViewNameDisplayInHomePage = createdVendorViewinHomePage.getText();
		String expcreateViewNameDisplayInHomePage = "VendorA";

		System.out.println("View Display In Home page Value Actual    : " + actcreateViewNameDisplayInHomePage   +     "	 Value Expected : " + expcreateViewNameDisplayInHomePage);

		
		
		if(actSaveViewMessage.equalsIgnoreCase(expSaveViewMessage) && actcreateViewNameDisplayInHomePage.equalsIgnoreCase(expcreateViewNameDisplayInHomePage) )
		{
			return true;
		}
		else
		{
			return false;
		} 
	 
}

public boolean checkSavedViewofVendorinPVVAT() throws InterruptedException
{
	 

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdVendorViewinHomePage));
		createdVendorViewinHomePage.click();
		Thread.sleep(2000);
		

		int actvoucherGridDocNoCount = voucherGridDocNo.size()-1;
		System.out.println("Number of Vouchers in Vendor View	"	+	actvoucherGridDocNoCount);
		int expvoucherGridDocNoCount=5;
		
		ArrayList<String>DocumentNum=new ArrayList<String>();
		
		for(int i=0;i<actvoucherGridDocNoCount;i++)
		{
			DocumentNum.add(voucherGridDocNo.get(i).getText());
		}

		
		Thread.sleep(2000);
		
		String actDocNumList=DocumentNum.toString();
		String expDocNumList="[SU/DUB/TEXT4, SU/DUB/TEXT3, SU/DUB/TEXT2, SU/DUB/TEXT1, SU/IND/TEXT2]";
		
		System.out.println("Actual Document List:		"	+	actDocNumList);
		System.out.println("Expected Document List:		"+	expDocNumList);
		
		
		for (int i = 0; i < actvoucherGridDocNoCount; i++) {
			String data = "SU/DUB/TEXT1";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNum=documentNumberTxt.getAttribute("value");
		String expDocNum="SU/DUB/TEXT1";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendorTxt=vendorAccountTxt.getAttribute("value");
		String expVendorTxt="Vendor A";
		
		System.out.println("Actual Document Number:		"	+	actDocNum	+"Expected:		"	+	expDocNum);
		System.out.println("Actual Vendor:		"	+	actVendorTxt	+"Expected:		"	+	expVendorTxt);
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCloseBtn));
		newCloseBtn.click();
		
		if(actvoucherGridDocNoCount==expvoucherGridDocNoCount && actDocNum.equalsIgnoreCase(expDocNum)
				&& actDocNumList.equalsIgnoreCase(expDocNumList))
		{
		
			return true;
		}
		else
		{
			return false;
		}
		 
	 
}




public boolean checkSavingWarehouseViewinPurchaseVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	 Thread.sleep(2000);
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createViewBtn));
		createViewBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewNameTxtInCreateView));
		viewNameTxtInCreateView.click();
		viewNameTxtInCreateView.sendKeys("Warehouse HYD");
		Thread.sleep(2000);
		viewNameTxtInCreateView.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTypeInCreateView));
		Select s=new Select(viewTypeInCreateView);
		s.selectByVisibleText("Detailed");
		Thread.sleep(2000);
		viewTypeInCreateView.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizationInCreateView));
		Select s2=new Select(authorizationInCreateView);
		s2.selectByVisibleText("All");
		Thread.sleep(2000);
		authorizationInCreateView.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(checkSatusInCreateView));
		Select s3=new Select(checkSatusInCreateView);
		s3.selectByVisibleText("Both");
		Thread.sleep(2000);
		checkSatusInCreateView.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filterOptionInCreateView));
		filterOptionInCreateView.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_WhereCondition));
		//viewFilter_WhereCondition.click();
		Select s4=new Select(viewFilter_WhereCondition);
		s4.selectByVisibleText("Where");
		Thread.sleep(1000);
		viewFilter_WhereCondition.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_selectField));
		viewFilter_selectField.click();
		Thread.sleep(2000);
		
		getAction().moveToElement(viewFilter_WarehouseExpandBtn).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_WarehouseExpandBtn));
		viewFilter_WarehouseExpandBtn.click();
		Thread.sleep(3000);
		getAction().moveToElement(viewFilter_WarehouseName).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_WarehouseName));
		viewFilter_WarehouseName.click();
	
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_selectOperator));
		viewFilter_selectOperator.click();
		Select s5= new Select(viewFilter_selectOperator);
		s5.selectByVisibleText("Equal to");
		Thread.sleep(2000);
		//viewFilter_selectOperator.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_CompareWith));
		viewFilter_CompareWith.click();
		Thread.sleep(2000);
		Select s6= new Select(viewFilter_CompareWith);
		s6.selectByVisibleText("value");
		Thread.sleep(2000);
		//viewFilter_CompareWith.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_ValueFieldTxt));
		viewFilter_ValueFieldTxt.click();
		viewFilter_ValueFieldTxt.sendKeys("Hyderabad");
		Thread.sleep(2000);
		viewFilter_ValueFieldTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveOptionInCreateView));
		saveOptionInCreateView.click();

		String expSaveViewMessage = "View saved successfully";

		String actSaveViewMessage = checkValidationMessage(expSaveViewMessage);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeBtn));
		customizeBtn.click();

		Thread.sleep(4000);
		
		click(Cust_ModifiedByCol);
		
		click(cs_TransactionExpBtn);
		
		click(cs_ExtraFieldsExpBtn);
		
		click(cs_TransactionWarehouseFieldExpBtn);
		
		
		getAction().doubleClick(cs_TransactionWarehouseNameField).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(custSaveBtn));
		custSaveBtn.click();
		
		 Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(moreViewsBtn));
		 moreViewsBtn.click();
		 Thread.sleep(2000);

		String actcreateViewNameDisplayInHomePage = createdWarehouseViewinHomePage.getText();
		String expcreateViewNameDisplayInHomePage = "Warehouse HYD";

		System.out.println("View Display In Home page Value Actual    : " + actcreateViewNameDisplayInHomePage   +     "	 Value Expected : " + expcreateViewNameDisplayInHomePage);

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(moreViewsBtn));
		 moreViewsBtn.click();
		 Thread.sleep(2000);
		
		if(actSaveViewMessage.equalsIgnoreCase(expSaveViewMessage) && actcreateViewNameDisplayInHomePage.equalsIgnoreCase(expcreateViewNameDisplayInHomePage) )
		{
			return true;
		}
		else
		{
			return false;
		} 
	 
}


@FindBy(xpath="//*[@id='btnMoreViews']")
public static WebElement moreViewsBtn;


public boolean checkSavedViewofWarehouseinPVVAT() throws InterruptedException
{
	 Thread.sleep(2000);
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(moreViewsBtn));
	 moreViewsBtn.click();
	 Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdWarehouseViewinHomePage));
		createdWarehouseViewinHomePage.click();
		Thread.sleep(2000);
		

		int actvoucherGridDocNoCount = voucherGridDocNo.size()-1;
		System.out.println("Number of Vouchers in Vendor View	"	+	actvoucherGridDocNoCount);
		int expvoucherGridDocNoCount=4;
		
		
		ArrayList<String>Warehousearray=new ArrayList<String>();
		
		for(int i=0;i<actvoucherGridDocNoCount;i++)
		{
			Warehousearray.add(voucherHomePageDepartmentColList.get(i).getText());
		}

		
		Thread.sleep(2000);
		
		String actWarehouseLsit=Warehousearray.toString();
		String expWarehouseList="";
		
		System.out.println("Actual Department List  " 	+	actWarehouseLsit);
		
		System.out.println("Expected Department List  " 	+	expWarehouseList);
		
		ArrayList<String>DocumentNum=new ArrayList<String>();
		
		for(int i=0;i<actvoucherGridDocNoCount;i++)
		{
			DocumentNum.add(voucherGridDocNo.get(i).getText());
		}

		
		Thread.sleep(2000);
		
		String actDocNumList=DocumentNum.toString();
		String expDocNumList="[SU/IND/TEXT8, SU/IND/TEXT7, SU/IND/TEXT6, SU/IND/TEXT2]";
		
		System.out.println("Actual Document List:		"	+	actDocNumList);
		System.out.println("Expected Document List:		"+	expDocNumList);
		

		for (int i = 0; i < actvoucherGridDocNoCount; i++) {
			String data = "SU/IND/TEXT6";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNum=documentNumberTxt.getAttribute("value");
		String expDocNum="SU/IND/TEXT6";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actVendorTxt=select1stRow_1stColumn.getText();
		String expVendorTxt="Hyderabad";
		
		System.out.println("Actual Document Number:		"	+	actDocNum	+"Expected:		"	+	expDocNum);
		System.out.println("Actual Department:		"	+	actVendorTxt	+"Expected:		"	+	expVendorTxt);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCloseBtn));
		newCloseBtn.click();
		
		if(actvoucherGridDocNoCount==expvoucherGridDocNoCount && actDocNum.equalsIgnoreCase(expDocNum)
				&& actDocNumList.equalsIgnoreCase(expDocNumList))
		{
		
			return true;
		}
		else
		{
			return false;
		}

	 
}


@FindBy(xpath="//*[@id='id_TransHomeFilter']")
public static WebElement homeFilterBtn;


//Filter Fields
	@FindBy(xpath="//*[@id='2064_2_AdvanceFilter_']/table/tbody/tr/td[1]/select")
	public static WebElement homeFilter_WhereCondition;
	
		
	@FindBy(xpath="//*[@id='2064_2_AdvanceFilter_']/table/tbody/tr/td[2]/input")
	public static WebElement homeFilter_selectField;
	
	
	
	@FindBy(xpath="//*[@id='2064_2_AdvanceFilter_']/table/tbody/tr/td[3]/select")
	public static WebElement homeFilter_selectOperator;
	
	
	@FindBy(xpath="//*[@id='2064_2_AdvanceFilter_']/table/tbody/tr/td[4]/select")
	public static WebElement homeFilter_CompareWith;
	
	
	@FindBy(xpath="//*[@id='2064_2_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	public static WebElement homeFilter_ValueField;
	
	@FindBy(xpath="//*[@id='advancefilter_master_2064_2_']")
	public static WebElement homeFilter_ValueFieldTxt;


	@FindBy(xpath="//*[@id='2064_2_AdvanceFilter_']/table/tbody/tr[2]/td[1]/select")
	public static WebElement homeFilterSecondRow_WhereCondition;
	
		
	@FindBy(xpath="//*[@id='2064_2_AdvanceFilter_']/table/tbody/tr[2]/td[2]/input")
	public static WebElement homeFilterSecondRow_selectField;
	
	
	
	@FindBy(xpath="//*[@id='2064_2_AdvanceFilter_']/table/tbody/tr[2]/td[3]/select")
	public static WebElement homeFilterSecondRow_selectOperator;
	
	
	@FindBy(xpath="//*[@id='2064_2_AdvanceFilter_']/table/tbody/tr[2]/td[4]/select")
	public static WebElement homeFilterSecondRow_CompareWith;
	
	
	@FindBy(xpath="//*[@id='2064_2_AdvanceFilter_']/table/tbody/tr[2]/td[5]/input")
	public static WebElement homeFilterSecondRow_ValueField;
	
	@FindBy(xpath="//*[@id='advancefilter_master_2064_2_']")
	public static WebElement homeFilterSecondRow_ValueFieldTxt;

	
		
	@FindBy(xpath="//input[@id='chkSaveFilter']")
	public static WebElement saveFilterPermenantlyChkBox;
	
	
	@FindBy(xpath="//input[@id='chkSaveFilter']//..//span")
	public static WebElement saveFilterPermenantlyChkBoxSelected;
	
	
	@FindBy(xpath="//*[@id='btnSetFilterVal']")
	public static WebElement filterOKBtn;
	
		
	@FindBy(xpath="//*[@id='id_transaction_viewcontainer']/li//*[contains(text(),'All Vouchers')]")
	public static WebElement AllVouchersView;
	
	
	public boolean checkSaveFilterPermenantlycheckBoxinHomeFilterOptioninPVVAT() throws InterruptedException
	{
		Thread.sleep(10000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AllVouchersView));
		AllVouchersView.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilterBtn));
		homeFilterBtn.click();
		Thread.sleep(15000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filterOptionInCreateView));
		filterOptionInCreateView.click();
		Thread.sleep(2000);
	
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filterOptionInCreateView));
		filterOptionInCreateView.click();
		Thread.sleep(2000);*/
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilter_WhereCondition));
		//homeFilter_WhereCondition.click();
		Select s4=new Select(homeFilter_WhereCondition);
		s4.selectByVisibleText("Where");
		Thread.sleep(1000);
		homeFilter_WhereCondition.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilter_selectField));
		homeFilter_selectField.click();
		Thread.sleep(2000);
		
		getAction().moveToElement(viewFilter_DepartmentExpandBtn).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_DepartmentExpandBtn));
		viewFilter_DepartmentExpandBtn.click();
		Thread.sleep(3000);
		getAction().moveToElement(viewFilter_DepartmentName).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_DepartmentName));
		viewFilter_DepartmentName.click();
	
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilter_selectOperator));
		homeFilter_selectOperator.click();
		Select s5= new Select(homeFilter_selectOperator);
		s5.selectByVisibleText("Equal to");
		Thread.sleep(2000);
		//homeFilter_selectOperator.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilter_CompareWith));
		homeFilter_CompareWith.click();
		Select s6= new Select(homeFilter_CompareWith);
		s6.selectByVisibleText("value");
		Thread.sleep(2000);
		//homeFilter_CompareWith.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilter_ValueFieldTxt));
		homeFilter_ValueFieldTxt.click();
		homeFilter_ValueFieldTxt.sendKeys("INDIA");
		Thread.sleep(2000);
		homeFilter_ValueFieldTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filterPlusBtn));
		filterPlusBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilterSecondRow_WhereCondition));
		//homeFilter_WhereCondition.click();
		Select s7=new Select(homeFilterSecondRow_WhereCondition);
		s7.selectByVisibleText("And");
		Thread.sleep(1000);
		homeFilterSecondRow_WhereCondition.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilterSecondRow_selectField));
		homeFilterSecondRow_selectField.click();
		Thread.sleep(2000);
		
		getAction().moveToElement(viewFilter_VendorAcExpandBtn).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_VendorAcExpandBtn));
		viewFilter_VendorAcExpandBtn.click();
		Thread.sleep(3000);
		getAction().moveToElement(viewFilter_VendorName).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewFilter_VendorName));
		viewFilter_VendorName.click();
	
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilterSecondRow_selectOperator));
		homeFilterSecondRow_selectOperator.click();
		Select s8= new Select(homeFilterSecondRow_selectOperator);
		s8.selectByVisibleText("Not equal to");
		Thread.sleep(2000);
		//homeFilter_selectOperator.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilterSecondRow_CompareWith));
		homeFilterSecondRow_CompareWith.click();
		Select s9= new Select(homeFilterSecondRow_CompareWith);
		s9.selectByVisibleText("value");
		Thread.sleep(2000);
		//homeFilterSecondRow_CompareWith.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilter_ValueFieldTxt));
		homeFilter_ValueFieldTxt.click();
		homeFilter_ValueFieldTxt.sendKeys("Vendor A");
		Thread.sleep(2000);
		homeFilter_ValueFieldTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveFilterPermenantlyChkBoxSelected));
		if(saveFilterPermenantlyChkBox.isSelected()==false)
		{
			saveFilterPermenantlyChkBoxSelected.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filterOKBtn));
		filterOKBtn.click();
		Thread.sleep(4000);

		int actvoucherGridDocNoCount = voucherGridDocNo.size()-1;
		System.out.println("Number of Vouchers after Filter	"	+	actvoucherGridDocNoCount);
		int expvoucherGridDocNoCount=7;
		
		ArrayList<String>DocumentNum=new ArrayList<String>();
		
		for(int i=0;i<actvoucherGridDocNoCount;i++)
		{
			DocumentNum.add(voucherGridDocNo.get(i).getText());
		}

		for (int i = 0; i < actvoucherGridDocNoCount; i++) {
			String data = "SU/IND/TEXT3";
			
				
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}
		Thread.sleep(2000);
		
		String actDocNumList=DocumentNum.toString();
		String expDocNumList="[SU/IND/TEXT8, SU/IND/TEXT7, SU/IND/TEXT6, SU/IND/TEXT5, SU/IND/TEXT4, SU/IND/TEXT3, SU/IND/TEXT1]";
		System.out.println("Actual Document List:		"	+	actDocNumList);
		System.out.println("Expected Document List:		"+	expDocNumList);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNum=documentNumberTxt.getAttribute("value");
		String expDocNum="SU/IND/TEXT3";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDeptTxt=departmentTxt.getAttribute("value");
		String expDeptTxt="INDIA";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actWarehouseTxt=select1stRow_1stColumn.getText();
		String expWarehouseTxt="Mumbai";
		
		System.out.println("Actual Document Number:		"	+	actDocNum		+"Expected:		"	+	expDocNum);
		System.out.println("Actual Warehouse:			"	+	actWarehouseTxt	+"Expected:		"	+	expWarehouseTxt);
		System.out.println("Actual Department:			"	+	actDeptTxt		+"Expected:		"	+	expDeptTxt);
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newCloseBtn));
		newCloseBtn.click();
		
		if(actDocNumList.equalsIgnoreCase(expDocNumList)&&actvoucherGridDocNoCount==expvoucherGridDocNoCount && actDocNum.equalsIgnoreCase(expDocNum)
				&&actDeptTxt.equalsIgnoreCase(expDeptTxt) && actWarehouseTxt.equalsIgnoreCase(expWarehouseTxt) )
		{
		
			return true;
		}
		else
		{
			return false;
		}

		
		
	}
	
	
	
	
	public boolean checkDisablingSaveFilterPermenantlyChkBoxinPVVATVoucher() throws InterruptedException
	{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeFilterBtn));
		homeFilterBtn.click();
		Thread.sleep(10000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveFilterPermenantlyChkBoxSelected));
		if(saveFilterPermenantlyChkBox.isSelected()==true)
		{
			saveFilterPermenantlyChkBoxSelected.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(filterOKBtn));
		filterOKBtn.click();
		Thread.sleep(4000);

		int actvoucherGridDocNoCount = voucherGridDocNo.size()-1;
		
		int expvoucherGridDocNoCount=7;
		
		System.out.println("Number of Vouchers after Filter	"	+	actvoucherGridDocNoCount 		+	"Expected	:"	+	expvoucherGridDocNoCount);
		
		if(actvoucherGridDocNoCount==expvoucherGridDocNoCount)
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	@FindBy(xpath="(//span[@class='icon-pluse icon-font6'])[2]")
	public static WebElement filterPlusBtn;

	public boolean checkSavingOpeningBalanceVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		

		Thread.sleep(3000);
		
		getDriver().navigate().refresh();
		Thread.sleep(3000);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsJournalsMenu));
		financialsTransactionsJournalsMenu.click();
		Thread.sleep(2000);
		
		getAction().moveToElement(openingBalancesVoucher).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingBalancesVoucher));
		openingBalancesVoucher.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		String docno=documentNumberTxt.getAttribute("value");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.click();

		voucherHeaderCurrency.sendKeys(Keys.SHIFT,Keys.HOME);

		voucherHeaderCurrency.sendKeys(Keys.SPACE);

		int currencycount=currencyListCount.size();

		System.err.println(currencycount);

		for(int i=0 ; i < currencycount ;i++)
		{
			String data=currencyListCount.get(i).getText();

			if(data.equalsIgnoreCase("INR"))
			{
				currencyListCount.get(i).click();

				break;
			}
		}

		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);

		int OpeningBalDepartmentListCount=openingBalDepartmentList.size();

		for (int i = 0; i < OpeningBalDepartmentListCount; i++) 
		{
			String data=openingBalDepartmentList.get(i).getText();

			if (data.equalsIgnoreCase("INDIA"))
			{
				openingBalDepartmentList.get(i).click();

				break;
			}
		}

		departmentTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AccountTxt));
		enter_AccountTxt.click();
		enter_AccountTxt.sendKeys("C");

		int accountCount=openingBalAccountListInGrid.size();

		System.err.println(accountCount);

		for(int i=0 ; i < accountCount ;i++)
		{
			String data=openingBalAccountListInGrid.get(i).getText();

			if(data.equalsIgnoreCase("Customer A"))
			{
				openingBalAccountListInGrid.get(i).click();

				break;
			}
		}

		enter_AccountTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_DebitTxt));
		enter_DebitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_CreditTxt));
		enter_CreditTxt.sendKeys("100");
		enter_CreditTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingBalancesSaveBtn));
		openingBalancesSaveBtn.click();
		Thread.sleep(2000);
		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": 1";
		
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(2000);
		
		
		
		if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) )
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
		

	@FindBy(xpath="//*[@class='icon-settings hiconright2']")
	public static WebElement pettyCash_SettingsBtn;


	public boolean checkSavingTriggersinPettyCashVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		Thread.sleep(4000);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pettyCash));
		pettyCash.click();
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pettyCash_SettingsBtn));
		pettyCash_SettingsBtn.click();
		
		Thread.sleep(5000);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();

		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Petty Cash Trigger");
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(50000);
		
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
		if(editLayoutTriggerRaiseDoc.isSelected()==false)
		{
			editLayoutTriggerRaiseDocSelected.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
		Select s = new Select(editLayoutTriggersAllVouchers);
		s.selectByVisibleText("Petty Cash New");
		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		Thread.sleep(2000);

		String expSaveMessage = "Data saved successfully";
		String actSaveMessage = checkValidationMessage(expSaveMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();
		Thread.sleep(3000);
		
		if(actSaveMessage.equalsIgnoreCase(expSaveMessage) && actUpdateMessage.equalsIgnoreCase(expUpdateMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
		

	public boolean checkSavingPettyCashVoucherforTriggers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		newBtn.click();
		Thread.sleep(2000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.TAB);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vcustomerAccountTxt));
		vcustomerAccountTxt.click();
		vcustomerAccountTxt.sendKeys("Bank");
		Thread.sleep(2000);
		vcustomerAccountTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.click();
		departmentTxt.sendKeys(Keys.SPACE);

		ArrayList<String> actDepartmentList = new ArrayList<String>();

		int departmentCount = departmentListCount.size();

		System.err.println(departmentCount);

		for (int i = 0; i < departmentCount; i++) {
			String data = departmentListCount.get(i).getText();

			if (data.equalsIgnoreCase("INDIA")) {
				departmentListCount.get(i).click();
				break;
			}
		}

		departmentTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
		enter_PurchaseAccountTxt.sendKeys("Customer A");
		Thread.sleep(2000);
		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Amount));
		enter_Amount.sendKeys("50");
		Thread.sleep(1000);
		enter_Amount.sendKeys(Keys.TAB);

		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
		billRefPickIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
		billRefOkBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingBalancesSaveBtn));
		openingBalancesSaveBtn.click();
		Thread.sleep(2000);
		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": 1";
		
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2))
		{
			return true;
		}
		
		else
		{
			return false;
		}
		
	}

	public ArrayList<String> checkSavedDocumentinPettyCashVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(3000);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		documentNumberTxt.sendKeys("1");
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);
	Thread.sleep(3000);
		checkValidationMessage("Voucher loaded successfully");
		Thread.sleep(3000);
		
		ArrayList<String> doc1 = new ArrayList<String>();
		
		String Account=select1stRow_1stColumn.getText();
		doc1.add(Account);
		
		String Amount=select1stRow_2ndColumn.getText();
		doc1.add(Amount);
		
		String Reference=select1stRow_3rdColumn.getText();
		doc1.add(Reference);

		
		
		System.out.println(doc1);
		return doc1;
		
	}


	@FindBy(xpath="//span[contains(text(),'Petty Cash New')]")
	public static WebElement pettyCashNewVoucher;

	public boolean checkRaisedPettyCashNewVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		ArrayList<String>doc1=checkSavedDocumentinPettyCashVoucher();
		
		Thread.sleep(4000);
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cashAndBankMenu));
		cashAndBankMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pettyCashNewVoucher));
		pettyCashNewVoucher.click();
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for (int i = 0; i < voucherGridDocNoCount; i++) {
			String data = "1";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(2000);
		
		ArrayList<String> doc2 = new ArrayList<String>();
		
		String Account=select1stRow_1stColumn.getText();
		doc2.add(Account);
		
		String Amount=select1stRow_2ndColumn.getText();
		doc2.add(Amount);
		
		String Reference=select1stRow_3rdColumn.getText();
		doc2.add(Reference);

		System.out.println("Actual		:	"	+	doc1);
		System.out.println("Expected	:	"	+	doc2);
		
		if(doc1.equals(doc2))
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
		
	}
	@FindBy(xpath="//*[@id='cmbUserTypeMaster']")
	public static WebElement comboSearchBox;

	@FindBy(xpath="//*[@id='ol_treeNavigation']")
	public static WebElement navigationAcc;;


	public static void checkEmailtoVendorAcc() throws InterruptedException
	{
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu)); 
		 homeMenu.click();
		 Thread.sleep(2000);
		 
				  
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu)); 
				 mastersMenu.click();
				  
				 Thread.sleep(2000);
				 
				  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
				  accounts.click();
				  Thread.sleep(6000);
				 
				  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(comboSearchBox));
					comboSearchBox.click();
					comboSearchBox.clear();
					comboSearchBox.sendKeys("Vendor B");
					Thread.sleep(2000);
					comboSearchBox.sendKeys(Keys.TAB);
					
					Thread.sleep(2000);
					 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterEditBtn));
					masterEditBtn.click();
					Thread.sleep(4000);
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(detailsTab));
					detailsTab.click();
					Thread.sleep(3000);

					getAction().moveToElement(emailTxt).build().perform();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTxt));
					emailTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
					emailTxt.sendKeys("emailvalidationone@gmail.com");
					emailTxt.sendKeys(Keys.TAB);

					Thread.sleep(2500);
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
					saveBtn.click();
					Thread.sleep(2000);
					
					

		
	}
		

	@FindBy(xpath="//*[@id='lblnchecked']/input")

	public static List<WebElement> accChkBoxList;

	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[11]")
	public static List<WebElement> accNameList;


	///PO Trigger to raise a Purchase Return Voucher  for Doc number


	public boolean checkSavingTriggerinPOVoucherBasedOnSaveforDocumentNumberSeries() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Configure Transactions");

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();

		Thread.sleep(3000);

	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentCustomization45Btn));
		documentCustomization45Btn.click();
		Thread.sleep(2000);
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseOrdersBtn));
		purchaseOrdersBtn.click();

		Thread.sleep(2000);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toogleExpandBtn));
		toogleExpandBtn.click();

		Thread.sleep(2000);*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT, Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Doc Trigger");
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(50000);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
		if(editLayoutTriggerRaiseDoc.isSelected()==false)
		{
			editLayoutTriggerRaiseDocSelected.click();
		}
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersOnSaveRadioBtn));
		triggersOnSaveRadioBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
		Select s = new Select(editLayoutTriggersAllVouchers);
		s.selectByVisibleText("Purchases Returns");
		Thread.sleep(3000);

		getAction().moveToElement(trigger1stRow1stCol).build().perform();

		ClickUsingJs(trigger1stRow1stCol);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectFieldDropdown));

		Thread.sleep(1000);
		Select s1 = new Select(triggerSelectFieldDropdown);
		s1.selectByVisibleText("DocNo");
		Thread.sleep(3000);
		triggerSelectFieldDropdown.sendKeys(Keys.TAB);

		ClickUsingJs(trigger1stRow2ndCol);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerSelectMapDropdown));
		Select s2 = new Select(triggerSelectMapDropdown);
		s2.selectByVisibleText("DocNo");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		Thread.sleep(2000);

		String expSaveMessage = "Data saved successfully";
		String actSaveMessage = checkValidationMessage(expSaveMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();

		if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
			System.out.println("Test Pass : Trigger  Added Successfully");

			return true;
		} else {
			System.out.println("Test Fail : Trigger NOT Added Successfully");

			return false;
		}

	}

	public boolean checkSavingPOVoucherforDocumentNumber() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
		invTransactionsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("POV-1");
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		MRpurchaseAccountTxt.click();
		MRpurchaseAccountTxt.sendKeys("Purchase");
		Thread.sleep(3000);
		MRpurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();
		vendorAccountTxt.sendKeys("Vendor A");
		Thread.sleep(3000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValuetxt));
		departmentValuetxt.click();
		departmentValuetxt.sendKeys("AMERICA");
		Thread.sleep(3000);
		departmentValuetxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		

		Thread.sleep(1000);
		enter_ItemTxt.sendKeys("BR COGS ITEM");
		Thread.sleep(2000);

		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys("2");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys("6.5");
		enter_Rate.sendKeys(Keys.TAB);

		

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		String expMessage1 = "Voucher saved successfully";

		String actMessage = checkValidationMessage(expMessage1);
		// String expMessage2 = ": 3";

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage1);

		if (actMessage.startsWith(expMessage1)/* && actMessage.endsWith(expMessage2) */ )

		{
			return true;
		} else {
			return false;
		}


	}

	@FindBy(xpath = "//*[@id='id_transactionentry_previous']")
	public static WebElement previousBtn;


	public String checkDocNumberinPOVoucher() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();
		Thread.sleep(4000);
		
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		 String actDocNo=documentNumberTxt.getAttribute("value");
		 
			/*
			 * getWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * new_DeleteBtn)); new_DeleteBtn.click(); Thread.sleep(2000);
			 * 
			 * getWaitForAlert(); Thread.sleep(2000); getAlert().accept();
			 * Thread.sleep(2000);
			 */
		 
		 return actDocNo;
		
		
	}

	public boolean checkRaisedVoucherPurchseReturnforDocNumber() throws InterruptedException
	{
		Thread.sleep(4000);
		
		String actDocNo=checkDocNumberinPOVoucher();
		System.out.println(actDocNo);
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Purchases Returns");

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();

		Thread.sleep(3000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for (int i = 0; i < voucherGridDocNoCount; i++) {
			String data = "POV-1";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String expDocNo=documentNumberTxt.getAttribute("value");
		System.out.println(expDocNo);
		
		if(actDocNo.equalsIgnoreCase(expDocNo))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}



	public boolean checkSavingTriggerinPOVoucherBasedonOnAuthorizeforDocNumber() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Configure Transactions");

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();

		Thread.sleep(3000);

		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseOrdersBtn));
		purchaseOrdersBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT, Keys.HOME);
		Thread.sleep(1000);
		editLayoutTriggerName.sendKeys("Doc Trigger");
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys(Keys.TAB);
		Thread.sleep(50000);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersOnAuthRadioBtn));
		triggersOnAuthRadioBtn.click();

		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();
		Thread.sleep(2000);

		String expSaveMessage = "Data saved successfully";
		String actSaveMessage = checkValidationMessage(expSaveMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();

		if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
			System.out.println("Test Pass : Trigger  Added Successfully");

			return true;
		} else {
			System.out.println("Test Fail : Trigger NOT Added Successfully");

			return false;
		}

		
	}


	public boolean checkSavingPOVoucherforAuthorization() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
		invTransactionsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.click();
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("POV-2");
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		MRpurchaseAccountTxt.click();
		MRpurchaseAccountTxt.sendKeys("Purchase");
		Thread.sleep(3000);
		MRpurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();
		vendorAccountTxt.sendKeys("Vendor A");
		Thread.sleep(3000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValuetxt));
		departmentValuetxt.click();
		departmentValuetxt.sendKeys("AMERICA");
		Thread.sleep(3000);
		departmentValuetxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.click();
		enter_ItemTxt.sendKeys(Keys.SPACE);
		

		Thread.sleep(1000);
		enter_ItemTxt.sendKeys("BR COGS ITEM");
		Thread.sleep(2000);

		enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys("2");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys("6.5");
		enter_Rate.sendKeys(Keys.TAB);

		

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		String expMessage1 = "Voucher saved successfully";

		String actMessage = checkValidationMessage(expMessage1);
		// String expMessage2 = ": 3";

		System.out.println("Actual Message    : " + actMessage);
		System.out.println("Expected Message  : " + expMessage1);

		if (actMessage.startsWith(expMessage1)/* && actMessage.endsWith(expMessage2) */ )

		{
			return true;
		} else {
			return false;
		}

	}

	@FindBy(xpath="(//li[@id='masterChild_2560'])[2]")
	private static WebElement transactionAuthorizationText1forDocNo;
	
	@FindBy(xpath="//*[@title='Alerts']//i")
	private static WebElement noOfAlertsDisplay1;


	public boolean checkAuthorizingthePOVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

		Thread.sleep(4000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(noOfAlertsDisplay1));
		noOfAlertsDisplay1.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(alertTransactions));
		alertTransactions.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(transactionAuthorizations));
		transactionAuthorizations.click();
		Thread.sleep(2000);
		getWebDriverWait().until(ExpectedConditions.visibilityOf(transactionAuthorizationText1forDocNo));
		transactionAuthorizationText1forDocNo.click();
		Thread.sleep(2000);
		checkValidationMessage("Voucher loaded Successfully");

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(POauthorizeBtn));
		POauthorizeBtn.click();
		Thread.sleep(4000);

		if (getIsAlertPresent()) {
			getAlert().dismiss();

			getDriver().navigate().refresh();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
			invTransactionsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
			inventoryTransactionsPurchasesMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
			purchasesOrdersVoucher.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

			int vcount = grid_VoucherNoList.size();

			for (int i = 0; i < vcount; i++) {
				String VoucherNo = grid_VoucherNoList.get(i).getText();

				if (VoucherNo.equalsIgnoreCase("POV-2")) {
					if (grid_CheckBoxList.get(i).isSelected() == false) {
						grid_CheckBoxList.get(i).click();
						break;
					}
				}
			}

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
			authorizeBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
			newBtn.click();

			checkValidationMessage("Screen Opened");
		}
		
		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(openingStocksNewHomeRow1AuthorizationStatus));

		String actStatus = openingStocksNewHomeRow1AuthorizationStatus.getText();
		String expStatus = "Authorized";

		System.out.println("Authorization Status In home Page 	 :  " + actStatus + " Value Expected : " + expStatus);

		if (actStatus.equalsIgnoreCase(expStatus)) {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();

			return true;
		} else {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(invTransactionsMenu));
			invTransactionsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
			inventoryTransactionsPurchasesMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
			purchasesOrdersVoucher.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

			int vcount = grid_VoucherNoList.size();

			for (int i = 0; i < vcount; i++) {
				String VoucherNo = grid_VoucherNoList.get(i).getText();

				if (VoucherNo.equalsIgnoreCase("1")) {
					if (grid_CheckBoxList.get(i).isSelected() == false) {
						getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
						break;
					}
				}
			}

			checkValidationMessage("Voucher loaded Successfully");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(POauthorizeBtn));
			POauthorizeBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
			new_CloseBtn.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherhomeCloseBtn));
			voucherhomeCloseBtn.click();

			return false;
		}

	}

	public boolean checkRaisedVoucherPurchaseREturnsAfterAuthorizePOVoucher() throws InterruptedException
	{

		Thread.sleep(4000);
		
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
		System.out.println("Setting buton is enabled");
		serachMenuTextHomePage.click();
		serachMenuTextHomePage.sendKeys("Purchases Returns");

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
		searchMenuTextClick.click();

		Thread.sleep(3000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for (int i = 0; i < voucherGridDocNoCount; i++) {
			String data = "POV-2";
			if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
				voucherGridIndexChkBox.get(i).click();
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		editBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo="POV-2";
		System.out.println(actDocNo);
		System.out.println(expDocNo);
		
		if(actDocNo.equalsIgnoreCase(expDocNo))
		{
			return true;
		}
		else
		{
			return false;
		}
		

		
	}
	
	
	///Role
	
	@FindBy(xpath = "//input[@id='triggers_Role']")
	public static WebElement roleChkBox;
	
	@FindBy(xpath = "//input[@id='triggers_Role']//following-sibling::span")
	public static WebElement roleChkBoxSelected;
	
	@FindBy(xpath = "//input[@id='txtbox_triggers_Role_DD']")
	public static WebElement roleTxtBox;
	
	
	@FindBy(xpath = "//*[@id='ul_triggers_Role_DD']/li[6]/div/label/span[1]")
	public static WebElement role1ChkBox;
	
	@FindBy(xpath = "//*[@id='ul_triggers_Role_DD']//li//span[1]")
	public static List<WebElement> roleLabesList;
	
	@FindBy(xpath = "//*[@id='ul_triggers_Role_DD']//li//span[2]")
	public static List<WebElement> roleLabesChkBoxList;
	
	
	
	
	
	 public boolean checkSavingTriggertoSendEmailbyRole() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		 getDriver().navigate().refresh();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
			System.out.println("Setting buton is enabled");
			serachMenuTextHomePage.click();
			serachMenuTextHomePage.sendKeys("Configure Transactions");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
			searchMenuTextClick.click();

			Thread.sleep(6000);

		

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVATBtn));
			purchaseVouchersVATBtn.click();

			Thread.sleep(6000);
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
			triggersBtn.click();
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
			editLayoutTriggerName.click();
			editLayoutTriggerName.sendKeys(Keys.END,Keys.SHIFT, Keys.HOME);
			Thread.sleep(1000);
			editLayoutTriggerName.sendKeys("Map Trigger");
			Thread.sleep(4000);
			editLayoutTriggerName.sendKeys(Keys.TAB);
			Thread.sleep(8000);
			
			getAction().moveToElement(sendEmailChkBox).build().perform();
			Thread.sleep(2000);
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailChkBoxSelected));
			if (sendEmailChkBox.isSelected() == false) {
				sendEmailChkBoxSelected.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendEmailTxtBox));
				sendEmailTxtBox.click();
				sendEmailTxtBox.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
				sendEmailTxtBox.sendKeys("TriggersEmailLayout");
				Thread.sleep(3000);
			}
			Thread.sleep(2000);

			
			getAction().moveToElement(customerChkBoxSelected).build().perform();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerChkBoxSelected));
			if (customerChkBox.isSelected() == true) {
				customerChkBoxSelected.click();
			}

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginChkBoxSelected));
			if (loginChkBox.isSelected() == true) {
				loginChkBoxSelected.click();
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFieldChkBoxSelected));
			if (extraFieldChkBox.isSelected() == true) {
				extraFieldChkBoxSelected.click();
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(roleChkBoxSelected));
			if (roleChkBox.isSelected() == false) {
				roleChkBoxSelected.click();
			}
			
			Thread.sleep(2000);
			getAction().moveToElement(roleTxtBox).build().perform();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(roleTxtBox));
			roleTxtBox.click();
			Thread.sleep(4000);
			
		/*	int count=roleLabesList.size();
			for(int i=0;i<roleLabesList.size();i++)
			{
				String data="Role1";
				if(roleLabesList.get(i).getText().equalsIgnoreCase("Role1"))
				{
						Thread.sleep(2000);
						roleLabesList.get(i).click();
						Thread.sleep(2000);
					
					break;
				}
			}*/
			
			getAction().moveToElement(role1ChkBox).build().perform();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(role1ChkBox));
			role1ChkBox.click();
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
			editLayoutTriggersSaveBtn.click();
			Thread.sleep(2000);

			String expSaveMessage = "Data saved successfully";
			String actSaveMessage = checkValidationMessage(expSaveMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
			updateBtn.click();
			Thread.sleep(5000);
			String expUpdateMessage = "Data saved successfully";
			String actUpdateMessage = checkValidationMessage(expUpdateMessage);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
			settings_closeBtn.click();

			if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
				return true;
			} else {
				return false;
			}

		 
		 
	 }
	 
	
	
	 public boolean checkSavingPurchaseVoucherVATVoucherforSendtoRole() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	 {
	 	

	 	Thread.sleep(2000);
	 	getDriver().navigate().refresh();
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	 	financialsMenu.click();

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	 	financialsTransactionMenu.click();

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	 	financialsTransactionsPurchaseMenu.click();

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
	 	purchaseVouchersVat.click();

	 	Thread.sleep(2000);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	 	newBtn.click();

	 	checkValidationMessage("Screen opened");

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
	 	vendorAccountTxt.click();

	 	vendorAccountTxt.sendKeys(Keys.SPACE);

	 	ArrayList<String> actVendorAccountList = new ArrayList<String>();

	 	int count = vendorAccountListCount.size();

	 	System.err.println(count);

	 	for (int i = 0; i < count; i++) {
	 		String data = vendorAccountListCount.get(i).getText();

	 		if (data.equalsIgnoreCase("Vendor B")) {
	 			vendorAccountListCount.get(i).click();
	 			break;
	 		}
	 	}

	 	Thread.sleep(2000);
	 	vendorAccountTxt.sendKeys(Keys.TAB);

	 	Thread.sleep(1000);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
	 	voucherHeaderDueDate.click();
	 	voucherHeaderDueDate.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	 	voucherHeaderCurrency.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
	 	voucherHeaderCurrency.sendKeys(Keys.SPACE);

	 	int currencycount = currencyListCount.size();

	 	System.err.println(currencycount);

	 	for (int i = 0; i < currencycount; i++) {
	 		String data = currencyListCount.get(i).getText();

	 		if (data.equalsIgnoreCase("INR")) {
	 			currencyListCount.get(i).click();
	 			break;
	 		}
	 	}

	 	voucherHeaderCurrency.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	 	departmentTxt.click();
	 	departmentTxt.sendKeys(Keys.SPACE);

	 	ArrayList<String> actDepartmentList = new ArrayList<String>();

	 	int departmentCount = departmentListCount.size();

	 	System.err.println(departmentCount);

	 	for (int i = 0; i < departmentCount; i++) {
	 		String data = departmentListCount.get(i).getText();

	 		if (data.equalsIgnoreCase("INDIA")) {
	 			departmentListCount.get(i).click();
	 			break;
	 		}
	 	}

	 	departmentTxt.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
	 	placeOFSupplyTxt.click();
	 	placeOFSupplyTxt.clear();
	 	placeOFSupplyTxt.sendKeys("Abu Dhabi");
	 	Thread.sleep(2000);
	 	placeOFSupplyTxt.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
	 	jurisdictionTxt.click();
	 	jurisdictionTxt.sendKeys("DUBAI");
	 	Thread.sleep(2000);
	 	jurisdictionTxt.sendKeys(Keys.TAB);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVATNarrationTxt));
	 	purchaseVATNarrationTxt.click();
	 	purchaseVATNarrationTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	 	purchaseVATNarrationTxt.sendKeys("emailvalidationone@gmail.com");
	 	Thread.sleep(2000);
	 	purchaseVATNarrationTxt.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	select1stRow_1stColumn.click();
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
	 	enter_WarehouseTxt.click();

	 	enter_WarehouseTxt.sendKeys(Keys.SPACE);

	 	int warehousecount = warehouseBodyComboList.size();

	 	for (int i = 0; i < warehousecount; i++) {
	 		String data = warehouseBodyComboList.get(i).getText();

	 		if (data.equalsIgnoreCase("HYDERABAD")) {
	 			warehouseBodyComboList.get(i).click();
	 			break;
	 		}
	 	}

	 	enter_WarehouseTxt.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	 	enter_ItemTxt.click();
	 	enter_ItemTxt.sendKeys(Keys.SPACE);
	 	int pvvGridItemListCount = pvvGridItemList.size();
	 	for (int i = 0; i < pvvGridItemListCount; i++) {
	 		String Item = pvvGridItemList.get(i).getText();
	 		if (Item.equalsIgnoreCase("STD RATE COGS ITEM")) {
	 			pvvGridItemList.get(i).click();
	 			break;
	 		}
	 	}
	 	enter_ItemTxt.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_TaxCode));
	 	enter_TaxCode.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
	 	enter_PurchaseAccountTxt.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
	 	select1stRow_9thColumn.click();
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	 	enter_Quantity.click();
	 	enter_Quantity.clear();
	 	enter_Quantity.sendKeys("2");

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	 	select1stRow_11thColumn.click();
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	 	enter_Rate.click();
	 	enter_Rate.clear();
	 	enter_Rate.sendKeys("10");
	 	enter_Rate.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
	 	enter_Gross.click();
	 	enter_Gross.sendKeys(Keys.TAB);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));
	 	select1stRow_13thColumn.click();
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Discount));
	 	enter_Discount.click();
	 	enter_Discount.clear();
	 	enter_Discount.sendKeys("10");
	 	enter_Discount.sendKeys(Keys.TAB);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	 	select1stRow_14thColumn.click();

	 	/*
	 	 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
	 	 * enter_PvVat)); enter_PvVat.sendKeys(Keys.TAB);
	 	 */

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PvTaxable));
	 	enter_PvTaxable.sendKeys(Keys.TAB);

	 	Thread.sleep(2000);

	 	/*
	 	 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
	 	 * enter_Batch)); enter_Batch.click(); enter_Batch.sendKeys("PVATR1#V1");
	 	 * enter_Batch.sendKeys(Keys.TAB);
	 	 * 
	 	 * DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	 	 * 
	 	 * Calendar cal=Calendar.getInstance(); SimpleDateFormat currentDate = new
	 	 * SimpleDateFormat("dd/MM/yyyy"); cal.add(Calendar.DATE, 2);
	 	 * 
	 	 * String next2DaysDate=df.format(cal.getTime());
	 	 * 
	 	 * System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : "
	 	 * +next2DaysDate);
	 	 * 
	 	 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
	 	 * enter_Expirydate)); enter_Expirydate.click();
	 	 * enter_Expirydate.sendKeys(Keys.HOME);
	 	 * enter_Expirydate.sendKeys(next2DaysDate);
	 	 * enter_Expirydate.sendKeys(Keys.TAB);
	 	 * 
	 	 * Thread.sleep(2000);
	 	 * 
	 	 */

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
	 	voucherSaveBtn.click();

	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
	 	billRefNewReferenceTxt.click();
	 	Thread.sleep(2000);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
	 	pickBtn.click();

	 	Thread.sleep(2000);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
	 	Bill_OkBtn.click();

	 	Thread.sleep(2000);

	 	String expMessage1 = "Voucher saved successfully";

	 	String actMessage = checkValidationMessage(expMessage1);
	 	String expMessage2 = ": SU/IND/TEXT9";

	 	System.out.println("Actual Message    : " + actMessage);
	 	System.out.println("Expected Message  : " + expMessage1);

	 	if (actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2)) {
	 		return true;
	 	} else {
	 		return false;
	 	}


	 }


	 @FindBy(xpath="//*[@data-tooltip='Refresh']")
	 public static WebElement refreshBtn;


	 public boolean checkEmailValidationforUser1inRole() throws InterruptedException, AWTException, IOException
	 {
	 File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser1.pdf");
	 	
	 	if(Efile1.exists())
	 	{
	 		Efile1.delete();
	 		System.out.println("File deleted");
	 	}
	 	
	 	
	 	Robot robot = new Robot();   
	 	robot.keyPress(KeyEvent.VK_CONTROL);
	 	robot.keyPress(KeyEvent.VK_T);
	 	robot.keyRelease(KeyEvent.VK_CONTROL);
	 	robot.keyRelease(KeyEvent.VK_T);
	 	
	 	Thread.sleep(4000);
	 	
	 	ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

	  	System.out.println("openTabs"+openTabs);

	  	getDriver().switchTo().window(openTabs.get(0));
	  	Thread.sleep(2000);
	  	getDriver().switchTo().window(openTabs.get(1));
	 	

	  	getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");
	  	
	  	Thread.sleep(2000);
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
	  	userNameTxt.click();
	  	
	  	userNameTxt.sendKeys("emailvalidationtwo@gmail.com");
	  	
	  	Thread.sleep(2000);
	  	
	  	System.out.println("User enter text");
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
	  	NextBtn.click();
	  	
	  	Thread.sleep(2000);
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
	  	PasswordTxt.click();
	  	
	  	PasswordTxt.sendKeys("validationtwo");
	  	
	  	Thread.sleep(2000);
	  	
	  	NextBtn.click();
	  	
	  	Thread.sleep(10000);
	  	
	  boolean result=false;
	  /*	int count1 = row1mailFromList.size();
	  	System.out.println("No.of mails in the list"+count1);
	  	
	  	System.err.println("Email Body *"+emailBody.getText()+"*");*/
	  	
	  	/*if (emailBody.getText().isEmpty()==false) 
	  	{*/
	  		int count = emailSubjectList.size();
	  		System.err.println("No.of mails in the list"+count);
	 	 	if(count==0)
	 	 	{
	 	 		
	 	 		System.err.println("No Mails in the List   :"+count);
	 	 		
	 	 		
	 	 	}
	 	 	else if (emailBody2.getText().isEmpty()==false) 
	 	 	 	{
	  		
	 	 	for (int i = 0; i < count; i++) 
	  	{
	  		String mailFromTxt = emailSubjectList.get(i).getText();
	  		System.out.println(mailFromTxt);
	  	 	
	  		if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher")) 
	  		{
	  			emailSubjectList.get(i).click();     
	  			break;
	  		}
	  	}
	 	 	
	  			 	
	 	 	Thread.sleep(3000);
	 	
	 	Thread.sleep(3000);
	 	
	 	if(refreshBtn.isDisplayed())
	 	{
	 		
	 		System.out.println("Subject is not matching");
	 	 	
	 	}
	 	
	  		
	  	
	 	else
	 	{
	 		
	 		System.out.println("Subject is  matching");
	 		getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
	 	mailAttachmentDownloadBtn1.click();
	 	
	 	Thread.sleep(6000);
	 	
	 	 
	 	robot.keyPress(KeyEvent.VK_CONTROL);
	 robot.keyPress(KeyEvent.VK_J);
	 robot.keyRelease(KeyEvent.VK_J);
	 robot.keyRelease(KeyEvent.VK_CONTROL);


	 Thread.sleep(5000);

	 robot.keyPress(KeyEvent.VK_TAB);
	 robot.keyRelease(KeyEvent.VK_TAB);
	 robot.keyPress(KeyEvent.VK_TAB);
	 robot.keyRelease(KeyEvent.VK_TAB);

	 robot.keyPress(KeyEvent.VK_ENTER);
	 robot.keyRelease(KeyEvent.VK_ENTER);

	 Thread.sleep(3000);

	 robot.keyPress(KeyEvent.VK_CONTROL);
	 robot.keyPress(KeyEvent.VK_S);
	 robot.keyRelease(KeyEvent.VK_S);
	 robot.keyRelease(KeyEvent.VK_CONTROL);
	 	
	 Thread.sleep(2000);

	 		
	 		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailValidationforUser1.exe");
	 		
	 		Thread.sleep(8000);
	 		
	 		robot.keyPress(KeyEvent.VK_CONTROL);
	 		robot.keyPress(KeyEvent.VK_J);
	 		robot.keyRelease(KeyEvent.VK_J);
	 		robot.keyRelease(KeyEvent.VK_CONTROL);
	 		
	 		Thread.sleep(4000);
	 		
	 		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
	 			
	 		int actOpenWindowsCount = getDriver().getWindowHandles().size();
	 		int expOpenWindowsCount = 4;
	 		
	 		System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
	 		
	 		getDriver().switchTo().window(newTabs.get(3)).close();
	 	 	Thread.sleep(2000);
	 	 	getDriver().switchTo().window(newTabs.get(2)).close();
	 	 	Thread.sleep(2000);
	 	 	getDriver().switchTo().window(newTabs.get(1));
	 	 	Thread.sleep(2000);
	 	 	
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
	 	 	mailDeleteButton.click();
	 	 	
	 	 	Thread.sleep(2000);
	 	 	

	 	 	Thread.sleep(2000);
	 	 	
	 	 	String actAccount1PDF = getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser1.pdf";
	 		String expAccount1PDF = getBaseDir()+"\\autoIt\\ImportFiles\\EmailValidationforUser1.pdf";
	 		
	 	
	 	 
	 		PDFUtil pdfutil = new PDFUtil();
	 	
	 		boolean result1 = pdfutil.compare(actAccount1PDF, expAccount1PDF);
	 		Calendar cal=Calendar.getInstance();
	 		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	 		String currentDate = df.format(cal.getTime());
	 		
	 		String  actAccount1Data = pdfutil.getText(actAccount1PDF);
	 		String  expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("17/04/2024", currentDate);
	 				
	 		System.err.println(actAccount1Data);
	 		System.err.println(expAccount1Data);
	 		
	 		if(actAccount1Data.equalsIgnoreCase(expAccount1Data))
	 		{
	 			result=true;
	 		}
	 	 	
	 	}
	 		
	 	 	 	}	
	 		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());	
	 		ClickUsingJs(gmailUserBtn);

	 		Thread.sleep(3000);
	 		getDriver().switchTo().frame(SignOutFrame);
	 		Thread.sleep(2000);
	 		
	 		getAction().moveToElement(gmailSignOutBtn).click().build().perform();
	 		
	 		Thread.sleep(6000);
	 		
	 		getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
	 		
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
	 	 	removeAccountBtn.click();
	 	 	
	 	 	Thread.sleep(1000);
	 	 	
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
	 	 	removeDeleteBtn.click();
	 	 	
	 	 	Thread.sleep(2000);
	 	 	
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
	 	 	yesRemoveBtn.click();
	 	 	
	 	 	Thread.sleep(1000);
	 	 	
	 		getDriver().switchTo().window(newTabs.get(1)).close();
	 	 	Thread.sleep(2000);
	 	 	
	 	 	
	 	 	getDriver().switchTo().window(newTabs.get(0));
	 	 	Thread.sleep(2000);
	 	 
	 		
	 		
	 		
	 	 	 		
	 	 	if (result && count!=0) 
	 		{
	 		
	 			System.err.println("PDF file is as Expected");
	 			return true;
	 		}
	 		else
	 		{
	 								
	 			System.err.println("PDF file is not as Expected");
	 			return false;
	 		}
	 		
	 		
	 	} 
	 	

	 public boolean checkEmailValidationforUser2inRole() throws InterruptedException, AWTException, IOException
	 {
		 Thread.sleep(2000);
		 
	 File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser2.pdf");
	 	
	 	if(Efile1.exists())
	 	{
	 		Efile1.delete();
	 		System.out.println("File deleted");
	 	}
	 	
	 	
	 	Robot robot = new Robot();   
	 	robot.keyPress(KeyEvent.VK_CONTROL);
	 	robot.keyPress(KeyEvent.VK_T);
	 	robot.keyRelease(KeyEvent.VK_CONTROL);
	 	robot.keyRelease(KeyEvent.VK_T);
	 	
	 	Thread.sleep(4000);
	 	
	 	ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

	  	System.out.println("openTabs"+openTabs);

	  	getDriver().switchTo().window(openTabs.get(0));
	  	Thread.sleep(2000);
	  	getDriver().switchTo().window(openTabs.get(1));
	 	

	  	getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");
	  	
	  	Thread.sleep(2000);
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
	  	userNameTxt.click();
	  	
	  	userNameTxt.sendKeys("emailvalidationfour@gmail.com");
	  	
	  	Thread.sleep(2000);
	  	
	  	System.out.println("User enter text");
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
	  	NextBtn.click();
	  	
	  	Thread.sleep(2000);
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
	  	PasswordTxt.click();
	  	
	  	PasswordTxt.sendKeys("validationfour");
	  	
	  	Thread.sleep(2000);
	  	
	  	NextBtn.click();
	  	
	  	Thread.sleep(10000);
	  	
	  boolean result=false;
	  /*	int count1 = row1mailFromList.size();
	  	System.out.println("No.of mails in the list"+count1);
	  	
	  	System.err.println("Email Body *"+emailBody.getText()+"*");*/
	  	
	  	/*if (emailBody.getText().isEmpty()==false) 
	  	{*/
	  		int count = emailSubjectList.size();
	  		System.err.println("No.of mails in the list"+count);
	 	 	if(count==0)
	 	 	{
	 	 		
	 	 		System.err.println("No Mails in the List   :"+count);
	 	 		
	 	 		
	 	 	}
	 	 	else if (emailBody2.getText().isEmpty()==false) 
	 	 	 	{
	  		
	 	 	for (int i = 0; i < count; i++) 
	  	{
	  		String mailFromTxt = emailSubjectList.get(i).getText();
	  		System.out.println(mailFromTxt);
	  	 	
	  		if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher")) 
	  		{
	  			emailSubjectList.get(i).click();     
	  			break;
	  		}
	  	}
	 	 	
	  			 	
	 	 	Thread.sleep(3000);
	 	
	 	Thread.sleep(3000);
	 	
	 	if(refreshBtn.isDisplayed())
	 	{
	 		
	 		System.out.println("Subject is not matching");
	 	 	
	 	}
	 	
	  		
	  	
	 	else
	 	{
	 		
	 		System.out.println("Subject is  matching");
	 		getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
	 	mailAttachmentDownloadBtn1.click();
	 	
	 	Thread.sleep(3000);
	 	
	 	 
	 	robot.keyPress(KeyEvent.VK_CONTROL);
	 robot.keyPress(KeyEvent.VK_J);
	 robot.keyRelease(KeyEvent.VK_J);
	 robot.keyRelease(KeyEvent.VK_CONTROL);


	 Thread.sleep(5000);

	 robot.keyPress(KeyEvent.VK_TAB);
	 robot.keyRelease(KeyEvent.VK_TAB);
	 robot.keyPress(KeyEvent.VK_TAB);
	 robot.keyRelease(KeyEvent.VK_TAB);

	 robot.keyPress(KeyEvent.VK_ENTER);
	 robot.keyRelease(KeyEvent.VK_ENTER);

	 Thread.sleep(3000);

	 robot.keyPress(KeyEvent.VK_CONTROL);
	 robot.keyPress(KeyEvent.VK_S);
	 robot.keyRelease(KeyEvent.VK_S);
	 robot.keyRelease(KeyEvent.VK_CONTROL);
	 	
	 Thread.sleep(2000);

	 		
	 		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailValidationforUser2.exe");
	 		
	 		Thread.sleep(4000);
	 		
	 		robot.keyPress(KeyEvent.VK_CONTROL);
	 		robot.keyPress(KeyEvent.VK_J);
	 		robot.keyRelease(KeyEvent.VK_J);
	 		robot.keyRelease(KeyEvent.VK_CONTROL);
	 		
	 		Thread.sleep(2000);
	 		
	 		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
	 			
	 		int actOpenWindowsCount = getDriver().getWindowHandles().size();
	 		int expOpenWindowsCount = 4;
	 		
	 		System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
	 		
	 		getDriver().switchTo().window(newTabs.get(3)).close();
	 	 	Thread.sleep(1000);
	 	 	getDriver().switchTo().window(newTabs.get(2)).close();
	 	 	Thread.sleep(1000);
	 	 	getDriver().switchTo().window(newTabs.get(1));
	 	 	Thread.sleep(1000);
	 	 	
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
	 	 	mailDeleteButton.click();
	 	 	
	 	 	Thread.sleep(2000);
	 	 	

	 	 	Thread.sleep(2000);
	 	 	
	 	 	String actAccount1PDF = getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser2.pdf";
	 		String expAccount1PDF = getBaseDir()+"\\autoIt\\ImportFiles\\EmailValidationforUser2.pdf";
	 		
	 	
	 	 
	 		PDFUtil pdfutil = new PDFUtil();
	 	
	 		boolean result1 = pdfutil.compare(actAccount1PDF, expAccount1PDF);
	 		Calendar cal=Calendar.getInstance();
	 		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	 		String currentDate = df.format(cal.getTime());
	 		
	 		String  actAccount1Data = pdfutil.getText(actAccount1PDF);
	 		String  expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("16/04/2024", currentDate);
	 				
	 		System.err.println(actAccount1Data);
	 		System.err.println(expAccount1Data);
	 		
	 		if(actAccount1Data.equalsIgnoreCase(expAccount1Data))
	 		{
	 			result=true;
	 		}
	 	 	
	 	}
	 		
	 	 	 	}	
	 		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());	
	 		ClickUsingJs(gmailUserBtn);

	 		Thread.sleep(3000);
	 		getDriver().switchTo().frame(SignOutFrame);
	 		Thread.sleep(2000);
	 		
	 		getAction().moveToElement(gmailSignOutBtn).click().build().perform();
	 		
	 		Thread.sleep(6000);
	 		getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
	 		
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
	 	 	removeAccountBtn.click();
	 	 	
	 	 	Thread.sleep(1000);
	 	 	
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
	 	 	removeDeleteBtn.click();
	 	 	
	 	 	Thread.sleep(2000);
	 	 	
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
	 	 	yesRemoveBtn.click();
	 	 	
	 	 	Thread.sleep(1000);
	 	 	
	 		getDriver().switchTo().window(newTabs.get(1)).close();
	 	 	Thread.sleep(2000);
	 	 	
	 	 	
	 	 	getDriver().switchTo().window(newTabs.get(0));
	 	 	Thread.sleep(2000);
	 	 
	 		
	 		
	 		
	 	 	 		
	 	 	if (result && count!=0) 
	 		{
	 		
	 			System.err.println("PDF file is as Expected");
	 			return true;
	 		}
	 		else
	 		{
	 								
	 			System.err.println("PDF file is not as Expected");
	 			return false;
	 		}
	 		
	 		
	 	} 
	 	

	 public boolean checkEmailValidationforUser3inRole() throws InterruptedException, AWTException, IOException
	 {
	 File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser3.pdf");
	 	
	 	if(Efile1.exists())
	 	{
	 		Efile1.delete();
	 		System.out.println("File deleted");
	 	}
	 	
	 	
	 	Robot robot = new Robot();   
	 	robot.keyPress(KeyEvent.VK_CONTROL);
	 	robot.keyPress(KeyEvent.VK_T);
	 	robot.keyRelease(KeyEvent.VK_CONTROL);
	 	robot.keyRelease(KeyEvent.VK_T);
	 	
	 	Thread.sleep(4000);
	 	
	 	ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

	  	System.out.println("openTabs"+openTabs);

	  	getDriver().switchTo().window(openTabs.get(0));
	  	Thread.sleep(2000);
	  	getDriver().switchTo().window(openTabs.get(1));
	 	

	  	getDriver().get("https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");
	  	
	  	Thread.sleep(2000);
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameTxt));
	  	userNameTxt.click();
	  	
	  	userNameTxt.sendKeys("emailvalidationfive@gmail.com");
	  	
	  	Thread.sleep(2000);
	  	
	  	System.out.println("User enter text");
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NextBtn));
	  	NextBtn.click();
	  	
	  	Thread.sleep(2000);
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordTxt));
	  	PasswordTxt.click();
	  	
	  	PasswordTxt.sendKeys("validationfive");
	  	
	  	Thread.sleep(2000);
	  	
	  	NextBtn.click();
	  	
	  	Thread.sleep(10000);
	  	
	  boolean result=false;
	  /*	int count1 = row1mailFromList.size();
	  	System.out.println("No.of mails in the list"+count1);
	  	
	  	System.err.println("Email Body *"+emailBody.getText()+"*");*/
	  	
	  	/*if (emailBody.getText().isEmpty()==false) 
	  	{*/
	  		int count = emailSubjectList.size();
	  		System.err.println("No.of mails in the list"+count);
	 	 	if(count==0)
	 	 	{
	 	 		
	 	 		System.err.println("No Mails in the List   :"+count);
	 	 		
	 	 		
	 	 	}
	 	 	else if (emailBody2.getText().isEmpty()==false) 
	 	 	 	{
	  		
	 	 	for (int i = 0; i < count; i++) 
	  	{
	  		String mailFromTxt = emailSubjectList.get(i).getText();
	  		System.out.println(mailFromTxt);
	  	 	
	  		if (mailFromTxt.equalsIgnoreCase("Email Trigger Voucher")) 
	  		{
	  			emailSubjectList.get(i).click();     
	  			break;
	  		}
	  	}
	 	 	
	  			 	
	 	 	Thread.sleep(3000);
	 	
	 	Thread.sleep(3000);
	 	
	 	if(refreshBtn.isDisplayed())
	 	{
	 		
	 		System.out.println("Subject is not matching");
	 	 	
	 	}
	 	
	  		
	  	
	 	else
	 	{
	 		
	 		System.out.println("Subject is  matching");
	 		getAction().moveToElement(mailAttachmentDownloadBtn1).build().perform();
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailAttachmentDownloadBtn1));
	 	mailAttachmentDownloadBtn1.click();
	 	
	 	Thread.sleep(3000);
	 	
	 	 
	 	robot.keyPress(KeyEvent.VK_CONTROL);
	 robot.keyPress(KeyEvent.VK_J);
	 robot.keyRelease(KeyEvent.VK_J);
	 robot.keyRelease(KeyEvent.VK_CONTROL);


	 Thread.sleep(5000);

	 robot.keyPress(KeyEvent.VK_TAB);
	 robot.keyRelease(KeyEvent.VK_TAB);
	 robot.keyPress(KeyEvent.VK_TAB);
	 robot.keyRelease(KeyEvent.VK_TAB);

	 robot.keyPress(KeyEvent.VK_ENTER);
	 robot.keyRelease(KeyEvent.VK_ENTER);

	 Thread.sleep(3000);

	 robot.keyPress(KeyEvent.VK_CONTROL);
	 robot.keyPress(KeyEvent.VK_S);
	 robot.keyRelease(KeyEvent.VK_S);
	 robot.keyRelease(KeyEvent.VK_CONTROL);
	 	
	 Thread.sleep(2000);

	 		
	 		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\EmailValidationforUser3.exe");
	 		
	 		Thread.sleep(4000);
	 		
	 		robot.keyPress(KeyEvent.VK_CONTROL);
	 		robot.keyPress(KeyEvent.VK_J);
	 		robot.keyRelease(KeyEvent.VK_J);
	 		robot.keyRelease(KeyEvent.VK_CONTROL);
	 		
	 		Thread.sleep(2000);
	 		
	 		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
	 			
	 		int actOpenWindowsCount = getDriver().getWindowHandles().size();
	 		int expOpenWindowsCount = 4;
	 		
	 		System.out.println("Number of Windows  : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
	 		
	 		getDriver().switchTo().window(newTabs.get(3)).close();
	 	 	Thread.sleep(1000);
	 	 	getDriver().switchTo().window(newTabs.get(2)).close();
	 	 	Thread.sleep(1000);
	 	 	getDriver().switchTo().window(newTabs.get(1));
	 	 	Thread.sleep(1000);
	 	 	
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
	 	 	mailDeleteButton.click();
	 	 	
	 	 	Thread.sleep(2000);
	 	 	

	 	 	Thread.sleep(2000);
	 	 	
	 	 	String actAccount1PDF = getBaseDir()+"\\autoIt\\ExportFiles\\EmailValidationforUser3.pdf";
	 		String expAccount1PDF = getBaseDir()+"\\autoIt\\ImportFiles\\EmailValidationforUser3.pdf";
	 		
	 	
	 	 
	 		PDFUtil pdfutil = new PDFUtil();
	 	
	 		boolean result1 = pdfutil.compare(actAccount1PDF, expAccount1PDF);
	 		Calendar cal=Calendar.getInstance();
	 		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	 		String currentDate = df.format(cal.getTime());
	 		
	 		String  actAccount1Data = pdfutil.getText(actAccount1PDF);
	 		String  expAccount1Data = pdfutil.getText(expAccount1PDF).replaceAll("16/04/2024", currentDate);
	 				
	 		System.err.println(actAccount1Data);
	 		System.err.println(expAccount1Data);
	 		
	 		if(actAccount1Data.equalsIgnoreCase(expAccount1Data))
	 		{
	 			result=true;
	 		}
	 	 	
	 	}
	 		
	 	 	 	}	
	 	 	
	 	 	getDriver().navigate().refresh();
	 	 	Thread.sleep(4000);
	 		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());	
	 		ClickUsingJs(gmailUserBtn);

	 		Thread.sleep(3000);
	 		getDriver().switchTo().frame(SignOutFrame);
	 		Thread.sleep(2000);
	 		
	 		getAction().moveToElement(gmailSignOutBtn).click().build().perform();
	 		
	 		Thread.sleep(6000);
	 		
	 		
	 		getDriver().get(
					"https://accounts.google.com/ServiceLogin/identifier?service=mail&passive=true&rm=false&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ss=1&scc=1&ltmpl=default&ltmplcache=2&emr=1&osid=1&flowName=GlifWebSignIn&flowEntry=AddSession");

			Thread.sleep(2000);
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeAccountBtn));
	 	 	removeAccountBtn.click();
	 	 	
	 	 	Thread.sleep(1000);
	 	 	
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeDeleteBtn));
	 	 	removeDeleteBtn.click();
	 	 	
	 	 	Thread.sleep(2000);
	 	 	
	 	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(yesRemoveBtn));
	 	 	yesRemoveBtn.click();
	 	 	
	 	 	Thread.sleep(1000);
	 	 	
	 		getDriver().switchTo().window(newTabs.get(1)).close();
	 	 	Thread.sleep(2000);
	 	 	
	 	 	
	 	 	getDriver().switchTo().window(newTabs.get(0));
	 	 	Thread.sleep(2000);
	 	 
	 		
	 		
	 		
	 	 	 		
	 	 	if (result && count!=0) 
	 		{
	 		
	 			System.err.println("PDF file is as Expected");
	 			return true;
	 		}
	 		else
	 		{
	 								
	 			System.err.println("PDF file is not as Expected");
	 			return false;
	 		}
	 		
	 		
	 	} 
	 	


	 
	 @FindBy(xpath="//*[@class='icon-settings hiconright2']")
	 public static WebElement SO_SettingsBtn;

	 @FindBy(xpath="//*[@id='navigationtab14']")
	 public static WebElement SO_InventoryOptionsTab;

	 @FindBy(xpath="//*[@id='inventoryOpt_ReservationType']")
	 public static WebElement SO_ReservationTypeSelect;

	 
	 public boolean checkSavingTriggerinMRNVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
	 	
	 	 Thread.sleep(2000);
	 	 getDriver().navigate().refresh();
	 		Thread.sleep(2000);
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	 		System.out.println("Setting buton is enabled");
	 		serachMenuTextHomePage.click();
	 		serachMenuTextHomePage.sendKeys("Material Receipt Notes");

	 		Thread.sleep(2000);

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	 		searchMenuTextClick.click();

	 		Thread.sleep(4000);
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mrnSettingsBtn));
	 		mrnSettingsBtn.click();
	 		Thread.sleep(4000);

	 		
	 		
	 		getAction().moveToElement(triggersBtn).build().perform();
	 		Thread.sleep(2000);
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
	 		triggersBtn.click();
	 		Thread.sleep(4000);
	 		
	 		  Thread.sleep(4000);
	 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
	 			editLayoutTriggerName.click();
	 			Thread.sleep(2000);
	 			editLayoutTriggerName.sendKeys(Keys.SHIFT, Keys.HOME);
	 			Thread.sleep(1000);
	 			editLayoutTriggerName.sendKeys("MRN Trigger to Sales Order");
	 			Thread.sleep(4000);
	 			editLayoutTriggerName.sendKeys(Keys.TAB);
	 			Thread.sleep(8000);
	 			
	 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerRaiseDocSelected));
	 			if(editLayoutTriggerRaiseDoc.isSelected()==false)
	 			{
	 				editLayoutTriggerRaiseDocSelected.click();
	 			}
	 			
	 			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
	 			Select s = new Select(editLayoutTriggersAllVouchers);
	 			s.selectByVisibleText("Sales Orders");
	 			Thread.sleep(3000);
	 			
	 			getAction().moveToElement(updateBtn).build().perform();
	 			Thread.sleep(2000);
	 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
	 			editLayoutTriggersSaveBtn.click();
	 			Thread.sleep(3000);

	 			String expSaveMessage = "Data saved successfully";
	 			String actSaveMessage = checkValidationMessage(expSaveMessage);

	 		
	 			
	 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
	 			updateBtn.click();
	 			Thread.sleep(4000);
	 			
	 			String expMessage="Data saved succesfully";
	 			String actMessage=checkValidationMessage(expMessage);
	 			
	 			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Settings_closeBtn));
	 			Settings_closeBtn.click();
	 			Thread.sleep(4000);
	 			 

	 			if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) {
	 				System.out.println("Test Pass : Trigger  Added Successfully");

	 				return true;
	 			} else {
	 				System.out.println("Test Fail : Trigger NOT Added Successfully");

	 				return false;
	 			}
	 		
	 }



	 public boolean checkUpdatingSOVouchertoAutoReservation() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
	 	
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(serachMenuTextHomePage));
	 	System.out.println("Setting buton is enabled");
	 	serachMenuTextHomePage.click();
	 	serachMenuTextHomePage.sendKeys("Sales Orders");

	 	Thread.sleep(2000);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchMenuTextClick));
	 	searchMenuTextClick.click();

	 	Thread.sleep(4000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SO_SettingsBtn));
	 	SO_SettingsBtn.click();
	 	
	 	Thread.sleep(4000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SO_InventoryOptionsTab));
	 	SO_InventoryOptionsTab.click();
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SO_ReservationTypeSelect));
	 	SO_ReservationTypeSelect.click();
	 	
	 	Select s=new Select(SO_ReservationTypeSelect);
	 	s.selectByVisibleText("Auto reservation");
	 	Thread.sleep(2000);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
	 	updateBtn.click();
	 	Thread.sleep(4000);
	 	
	 	String expMessage="Data saved successfully";
	 	String actMessage=checkValidationMessage(expMessage);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Settings_closeBtn));
	 	Settings_closeBtn.click();
	 	Thread.sleep(4000);
	 	
	 	if(actMessage.equalsIgnoreCase(expMessage))
	 	{
	 		return true;
	 	}
	 	else
	 	{
	 		return false;
	 	}
	 	
	 		
	 }

	 public boolean checkSavingMRNVoucherAfterTriggerSaved() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
	 	 Thread.sleep(2000);
	 	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mnr_NewBtn));
	 	 mnr_NewBtn.click();
	 	 Thread.sleep(2000);
	 	 
	 	 checkValidationMessage("Screen opened");

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
	 		vendorAccountTxt.click();

	 		vendorAccountTxt.sendKeys(Keys.SPACE);

	 		ArrayList<String> actVendorAccountList = new ArrayList<String>();

	 		int count = vendorAccountListCount.size();

	 		System.err.println(count);

	 		for (int i = 0; i < count; i++) {
	 			String data = vendorAccountListCount.get(i).getText();

	 			if (data.equalsIgnoreCase("Vendor B")) {
	 				vendorAccountListCount.get(i).click();
	 				break;
	 			}
	 		}

	 		vendorAccountTxt.sendKeys(Keys.TAB);

	 		Thread.sleep(3000);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mrn_WarehouseTxt));
	 		mrn_WarehouseTxt.click();

	 		mrn_WarehouseTxt.sendKeys(Keys.SPACE);

	 		ArrayList<String> actWarehouseList = new ArrayList<String>();

	 		int count1 = warehouseListCount.size();

	 		System.err.println(count);

	 		for (int i = 0; i < count; i++) {
	 			String data = warehouseListCount.get(i).getText();

	 			if (data.equalsIgnoreCase("HYDERABAD")) {
	 				warehouseListCount.get(i).click();
	 				break;
	 			}
	 		}

	 		mrn_WarehouseTxt.sendKeys(Keys.TAB);
	 		
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	 		departmentTxt.click();
	 		departmentTxt.sendKeys(Keys.SPACE);

	 		ArrayList<String> actDepartmentList = new ArrayList<String>();

	 		int departmentCount = departmentListCount.size();

	 		System.err.println(departmentCount);

	 		for (int i = 0; i < departmentCount; i++) {
	 			String data = departmentListCount.get(i).getText();

	 			if (data.equalsIgnoreCase("INDIA")) {
	 				departmentListCount.get(i).click();
	 				break;
	 			}
	 		}

	 		departmentTxt.sendKeys(Keys.TAB);

	 		Thread.sleep(3000);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 		select1stRow_1stColumn.click();
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	 		enter_ItemTxt.click();
	 		enter_ItemTxt.sendKeys(Keys.SPACE);
	 		
	 		int pvvGridItemListCount = pvvGridItemList.size();
	 		
	 		for (int i = 0; i < pvvGridItemListCount; i++) {
	 			
	 			String Item = pvvGridItemList.get(i).getText();
	 			
	 			if (Item.equalsIgnoreCase("BR COGS ITEM")) 
	 			{
	 				pvvGridItemList.get(i).click();
	 				break;
	 			}
	 		}
	 		enter_ItemTxt.sendKeys(Keys.TAB);	
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt));
	 		enter_PurchaseAccountTxt.sendKeys(Keys.TAB);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	 		select1stRow_5thColumn.click();
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	 		enter_Quantity.click();
	 		enter_Quantity.clear();
	 		enter_Quantity.sendKeys("2");

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
	 		select1stRow_7thColumn.click();
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	 		enter_Rate.click();
	 		enter_Rate.clear();
	 		enter_Rate.sendKeys("10");
	 		enter_Rate.sendKeys(Keys.TAB);

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
	 		enter_Gross.click();
	 		enter_Gross.sendKeys(Keys.TAB);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
	 		select1stRow_9thColumn.click();
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
	 		enter_Batch.click();
	 		enter_Batch.sendKeys("MRN#SOV1");
	 		enter_Batch.sendKeys(Keys.TAB);

	 		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

	 		Calendar cal = Calendar.getInstance();
	 		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
	 		cal.add(Calendar.DATE, 2);

	 		String next2DaysDate = df.format(cal.getTime());

	 		System.out.println(">>>>>>>>>Expiry Date After Adding 2Days>>>>>>>> : " + next2DaysDate);

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
	 		enter_Expirydate.click();
	 		enter_Expirydate.sendKeys(Keys.HOME);
	 		enter_Expirydate.sendKeys(next2DaysDate);
	 		enter_Expirydate.sendKeys(Keys.TAB);

	 		Thread.sleep(2000);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
	 		voucherSaveBtn.click();

	 		Thread.sleep(2000);
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
	 		billRefNewReferenceTxt.click();
	 		Thread.sleep(1000);

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
	 		pickBtn.click();

	 		Thread.sleep(2000);

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
	 		Bill_OkBtn.click();

	 		Thread.sleep(2000);

	 		String expMessage1 = "Voucher saved successfully";

	 		String actMessage = checkValidationMessage(expMessage1);
	 		//String expMessage2 = ": 1";

	 		System.out.println("Actual Message    : " + actMessage);
	 		System.out.println("Expected Message  : " + expMessage1);

	 		if (actMessage.startsWith(expMessage1) ) {
	 			return true;
	 		} else {
	 			return false;
	 		}
	 	
	 		
	 		
	 		
	 }


	 public boolean checkRaisedSalesOrderVoucherforAutoReservation() throws InterruptedException
	 {
	 	
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
	 	inventoryMenu.click();
	 	Thread.sleep(2000);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
	 	inventoryTransactionsMenu.click();
	 	Thread.sleep(2000);

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsSalesMenu));
	 	inventoryTransactionsSalesMenu.click();
	 	Thread.sleep(2000);
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesOrdersVoucher));
	 	salesOrdersVoucher.click();

	 	Thread.sleep(4000);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

	 	int voucherGridDocNoCount = voucherGridDocNo.size();

	 	for (int i = 0; i < voucherGridDocNoCount; i++) {
	 		String data = "1";
	 		if (voucherGridDocNo.get(i).getText().equalsIgnoreCase(data)) {
	 			voucherGridIndexChkBox.get(i).click();
	 			break;
	 		}
	 	}

	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
	 	editBtn.click();
	 	Thread.sleep(8000);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	 	select1stRow_8thColumn.click();
	 	Thread.sleep(2000);
	 	
	 	String actReserveColumn	=select1stRow_8thColumn.getText();
	 	String expReserveColumn	="Reserved";

	 	System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);
	 	
	 	if(actReserveColumn.equalsIgnoreCase(expReserveColumn))
	 	{
	 		return true;
	 	}
	 	else
	 	{
	 		return false;
	 	}
	 	
	 }
		  
	 public PurchaseVoucherVATandSalesInvoiceVATTriggersPage(WebDriver driver)
	 {
	    	
	    	PageFactory.initElements(driver, this);
	    	
	 }
	
	
	

}
