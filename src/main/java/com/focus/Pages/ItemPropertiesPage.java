package com.focus.Pages;

import java.awt.AWTException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.focus.Pages.LoginPage;
import com.focus.base.BaseEngine;
import com.focus.utilities.DriverUtility;

public class ItemPropertiesPage extends BaseEngine
{
	
	@FindBy(xpath="//*[@id='btnProperties']")
	public static WebElement itemPropBtn;
	
	@FindBy(xpath="//*[@id='chkDonotUpdateStock']")
	public static WebElement DonotUpdateStockChkBox;
	
	@FindBy(xpath="//*[@id='chkDonotUpdateStock']//..//span")
	public static WebElement DonotUpdateStockChkBoxSelected;
	
	@FindBy(xpath="//*[@id='chkDontMaintainStockbyBatch']")
	public static WebElement Chk_DontMaintainStockbyBatch;
	
	@FindBy(xpath="//*[@id='chkDontMaintainStockbyBatch']//..//span")
	public static WebElement DontMaintainStockbyBatch;
	
	@FindBy(xpath="//*[@id='chkDontinputBin']")
	public static WebElement Chk_DontinputBin;
	
	@FindBy(xpath="//*[@id='chkDontinputBin']//..//span")
	public static WebElement DontinputBin;
	
	@FindBy(xpath="//*[@id='chkDontMaintainStocksByRMA']")
	public static WebElement Chk_DontMaintainStocksByRMA;
	
	@FindBy(xpath="//*[@id='chkDontMaintainStocksByRMA']//..//span")
	public static WebElement DontMaintainStocksByRMA;
	
	@FindBy(xpath="//*[@id='btnPropOk']")
	public static WebElement itemProp_OkBtn;

	@FindBy(xpath="//span[contains(text(),'Stock Ledger')]")
	public static WebElement  stockLedger;
	
	@FindBy (xpath="//*[@id='txtsrch-term']")
	public static WebElement accountSearchTextArea;

	@FindBy (xpath="//*[@id='lblnchecked']")
	public static WebElement accountChkBox;
	
	
	@FindBy(xpath="//input[@id='id_body_37']")
	public static WebElement enter_Expirydate;
	
	@FindBy (xpath="//*[@id='div_OK']")
	private static WebElement RMAPopupOkBtn;
	
	
	
	public static boolean checkLogin() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	        Thread.sleep(3000);

	        getDriver().navigate().refresh();
	        Thread.sleep(3000);

	        LoginPage lp=new LoginPage(getDriver()); 

	        lp.checkLoginPageTitleByURLInputInBrowser();

	        String unamelt="su";

	        String pawslt="su";

	        lp.enterUserName(unamelt);
	        
	        Thread.sleep(2000);
	        
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

	       /* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	        companyLogo.click();*/

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
	
	
	@FindBy (xpath="//input[@id='id_body_85']")
	public static WebElement enter_Bin;
	
	
	public static boolean isElementPresent(By locatorKey) {
	    try {
	        getDriver().findElement(locatorKey);
	        return true;
	    } catch (org.openqa.selenium.NoSuchElementException e) {
	        return false;
	    }
	}
	
	@FindBy(xpath="(//*[@id='btnMasterSaveClick'])[2]")
	public static WebElement AccountSaveBtn;
	
	public boolean checkSavingMaterialReceiptNotesVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		

		
		Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(materialReceiptNotesVoucher));
		materialReceiptNotesVoucher.click();
		Thread.sleep(10000);
		new WebDriverWait(getDriver(), 300).until(ExpectedConditions.visibilityOf(newBtn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		//checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		customerAccountTxt.sendKeys("Bank");
		Thread.sleep(2000);
		customerAccountTxt.sendKeys(Keys.TAB);
		
		if(isElementPresent(By.xpath("//h4[contains(text(),'Account')]")))
		{
			AccountSaveBtn.click();
			
		}
		Thread.sleep(4000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.sendKeys("HYDERABAD");
		Thread.sleep(2000);
		warehouseTxt.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys("item3");
		Thread.sleep(1500);
		
		enter_ItemTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		

		
		 /*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_PurchaseAccountTxt)); 
		 enter_PurchaseAccountTxt.sendKeys("STD RATE COGS ACC");
		 Thread.sleep(2000);
		 enter_PurchaseAccountTxt.sendKeys(Keys.TAB);*/
		 
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		//enter_Quantity.click();
		//Thread.sleep(2000);
		enter_Quantity.sendKeys("5");
		//enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
		select1stRow_6thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_Rate.sendKeys("5");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
		select1stRow_12thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
		enter_Batch.click();
		enter_Batch.sendKeys("Batch1");
		enter_Batch.sendKeys(Keys.TAB);

		DateFormat df = new SimpleDateFormat("dd/MM/yyyy");

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
		cal.add(Calendar.DATE, 5); 

		String next5DaysDate=df.format(cal.getTime());

		System.out.println(">>>>>>>>>Expiry Date After Adding 5 Days>>>>>>>> : "+next5DaysDate);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchTxt));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binSearchBtn));
		binSearchBtn.click();
		Thread.sleep(5500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binAutoAllocateBtn));
		binAutoAllocateBtn.click();

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binOkBtn));
		binOkBtn.click();

		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Bin));
		enter_Bin.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
		enter_Expirydate.click();
		enter_Expirydate.sendKeys(Keys.HOME);
		enter_Expirydate.sendKeys(next5DaysDate);
		enter_Expirydate.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaScreenTitle));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaSerialNumberTxtField));
		rmaSerialNumberTxtField.sendKeys("MRN1,5");
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rmaAddBtn));
		rmaAddBtn.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();

		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		Thread.sleep(2000);
		
		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": 1";
		
		

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		System.out.println("Expected Message  : "+expMessage2);
		
		

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			
			return true;
		}
		else
		{
			
			return false;
		}

		
	}
	
	public boolean checkStockLedgerReportWithoutEnablingDonotUpdateStockChkBox() throws InterruptedException
	{
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger));
		stockLedger.click();
		
		//Thread.sleep(10000);
		new WebDriverWait(getDriver(), 300).until(ExpectedConditions.visibilityOf(accountSearchTextArea));

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));		
		 accountSearchTextArea.sendKeys("Item3");
		 accountSearchTextArea.sendKeys(Keys.ENTER);
		 	 
		 click(accountChkBox);
		 Thread.sleep(2000);
			
			click(sl_DateOptionDropdown);
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByValue("1");
			Thread.sleep(2000);
			
			
			click(sl_OkBtn);
			Thread.sleep(6000);
			
			
			String expRow1List = "[1, Item3Item3]";
			boolean actRow1List=ListComparisionWOOrder(report1stRowList,expRow1List);
			

			
			String expRow2List = "[2, "+getCurrentDate()+", MatRpt:1, 5.00, 5.00, 5.00, 25.00, 25.00, 5.00, 5.00]";
			boolean actRow2List=ListComparisionWOOrder(report2ndRowList,expRow2List);
			
			
			String expRow3List = "[3, Total, 5.00, 5.00, 5.00, 25.00, 25.00, 5.00, 5.00]";
			boolean actRow3List=ListComparisionWOOrder(report3rdRowList,expRow3List);
			
			
			if(actRow1List && actRow2List && actRow3List)
			{
				return true;
			}
			else
			{
				return false;
			}
			
			
		
	}
	
	@FindBy(xpath="//*[@id='221']")
	public static WebElement itemMenu;
		
		@FindBy(xpath="//*[@id='1105']")
		public static WebElement item;
	


	public boolean checkEnablingDonotUpdateStockChkBoxinItemProperties() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
	{
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemMenu));
		itemMenu.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item));
		item.click();
		
		Thread.sleep(6000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));		
		 accountSearchTextArea.sendKeys("Item3");
		 accountSearchTextArea.sendKeys(Keys.ENTER);
		 	 
		 click(accountChkBox);
		 Thread.sleep(2000);
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemPropBtn));
			itemPropBtn.click();
			
			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DonotUpdateStockChkBoxSelected));
			if(DonotUpdateStockChkBox.isSelected()==false)
			{
				
				DonotUpdateStockChkBoxSelected.click();
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemProp_OkBtn));
			itemProp_OkBtn.click();
			
			String expMessage="Saved Successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			Thread.sleep(1500);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));		
			 accountSearchTextArea.sendKeys("STOCK ITEM");
			 accountSearchTextArea.sendKeys(Keys.ENTER);
			 	 
			 click(accountChkBox);
			 Thread.sleep(2000);
			
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemPropBtn));
				itemPropBtn.click();
				
				Thread.sleep(8000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DonotUpdateStockChkBoxSelected));
				if(DonotUpdateStockChkBox.isSelected()==false)
				{
					
					DonotUpdateStockChkBoxSelected.click();
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DontMaintainStockbyBatch));
				if(Chk_DontMaintainStockbyBatch.isSelected()==false)
				{
					
					DontMaintainStockbyBatch.click();
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DontinputBin));
				if(Chk_DontinputBin.isSelected()==false)
				{
					
					DontinputBin.click();
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DontMaintainStocksByRMA));
				if(Chk_DontMaintainStocksByRMA.isSelected()==false)
				{
					
					DontMaintainStocksByRMA.click();
				}
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemProp_OkBtn));
				itemProp_OkBtn.click();
				
				String expMessage1="Saved Successfully";
				
				String actMessage1=checkValidationMessage(expMessage1);
				
				ProghornRestart("010");
			
			if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
			{
				return true;
			}
			else
			{
				return false;
			}
		
	}

	
	public boolean checkStockLedgerReportWithEnablingDonotUpdateStockChkBox() throws InterruptedException
	{
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger));
		stockLedger.click();
		
		Thread.sleep(10000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));		
		 accountSearchTextArea.sendKeys("Item3");
		 accountSearchTextArea.sendKeys(Keys.ENTER);
		 	 
		 click(accountChkBox);
		 Thread.sleep(2000);
			
			click(sl_DateOptionDropdown);
			Select s=new Select(sl_DateOptionDropdown);
			s.selectByValue("1");
			Thread.sleep(2000);
			
			
			click(sl_OkBtn);
			Thread.sleep(6000);
			
			
			boolean actReportTable1 = reportsTable.getText().isEmpty();
			System.out.println("Reports table is empty		"		+		actReportTable1);
			
			Thread.sleep(2500);
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger));
			stockLedger.click();
			Thread.sleep(2500);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));		
			 accountSearchTextArea.sendKeys("STOCK ITEM");
			 accountSearchTextArea.sendKeys(Keys.ENTER);
			 	 
			 click(accountChkBox);
			 Thread.sleep(2000);
				
				click(sl_DateOptionDropdown);
				Select s1=new Select(sl_DateOptionDropdown);
				s1.selectByValue("1");
				Thread.sleep(2000);
				
				
				click(sl_OkBtn);
				Thread.sleep(6000);
				
				
				boolean actReportTable2 = reportsTable.getText().isEmpty();
				System.out.println("Reports table is empty		"		+		actReportTable2);
			
			
			if(actReportTable1 && actReportTable2)
			{
				return true;
			}
			else
			{
				return false;
			}
			
			
		
	}
	
	
	
	public boolean checkLogoutItemPropertiesPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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
	

		public ItemPropertiesPage(WebDriver driver)
		{
			PageFactory.initElements(driver, this);
		}

		public boolean checkSalesInvoiceVATByEnablingDonotUpdateStockChkBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			
			getDriver().navigate().refresh();
			
			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			financialsMenu.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesMenu));
			salesMenu.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoicesVatBtn));
			salesInvoicesVatBtn.click();
			
			Thread.sleep(2000);
			
			newBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
			customerAccountTxt.click();
			Thread.sleep(1000);
			customerAccountTxt.sendKeys("Customer A");
			Thread.sleep(1000);
			customerAccountTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			departmentTxt.click();
			Thread.sleep(1000);
			departmentTxt.sendKeys("INDIA");
			Thread.sleep(1000);
			departmentTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeofSupplyTxt));
			placeofSupplyTxt.click();
			Thread.sleep(1000);
			placeofSupplyTxt.sendKeys("Abu Dhabi");
			Thread.sleep(1000);
			placeofSupplyTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			jurisdictionTxtt.sendKeys("Dubai");
			Thread.sleep(1000);
			jurisdictionTxtt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			select1stRow_1stColumn.click();
			Thread.sleep(1000);
			pvWareHouseTxt.sendKeys("HYDERABAD");
			Thread.sleep(1000);
			pvWareHouseTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			enter_ItemTxt.sendKeys("STOCK ITEM");
			Thread.sleep(1000);
			enter_ItemTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(TaxCode_input));
			TaxCode_input.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_SalesAccountTxt));
			enter_SalesAccountTxt.sendKeys("Sales - Computers");
			Thread.sleep(1000);
			enter_SalesAccountTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
			select1stRow_8thColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_AQTxt));
			enter_AQTxt.sendKeys("10");
			Thread.sleep(1000);
			enter_AQTxt.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_FQTxt));
			enter_FQTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			enter_RDTxt.sendKeys(Keys.TAB);
			Thread.sleep(1000);
			
			enter_Quantity.sendKeys(Keys.TAB);
			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			select1stRow_14thColumn.click();

			enter_Rate.sendKeys("10");
			enter_Rate.sendKeys(Keys.TAB);

			enter_Gross.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SIVAT_enter_Discount));
			SIVAT_enter_Discount.sendKeys(Keys.TAB);

			Thread.sleep(2000);

			String docno = documentNumberTxt.getAttribute("value");
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
			voucherSaveBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
			newReferenceTxt.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
			pickBtn.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
			Bill_OkBtn.click();

		
			String expMessage1 = "Voucher saved successfully";
			String expMessage2 = ": 1";
			String actMessage = checkValidationMessage(expMessage1);
			
			if(actMessage.startsWith(expMessage1)&&actMessage.endsWith(expMessage2))
			{
				return true;
			}
			else
			{
				return false;
			}
		}


}
