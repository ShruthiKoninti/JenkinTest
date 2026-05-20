package com.focus.Pages;

import java.awt.AWTException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.focus.Pages.LoginPage;
import com.focus.base.BaseEngine;
import com.focus.utilities.DriverUtility;

public class InfoPanelPage extends BaseEngine{
	
	@FindBy(xpath="//*[@id='id_transentry_infopanel_button_panel']/td[1]/i")
	public static WebElement openingStocks_InfoPanel;
	
	@FindBy(xpath="//*[@id='i_Trans']")
	public static WebElement InfoPanelAddBtn;
	
	@FindBy(xpath="(//*[@id='NewInfopanel']/a)[1]")
	public static WebElement InfoPanel_NewInfoPanel;
	
	@FindBy(xpath="//*[@id='DashletNameTrans']")
	public static WebElement InfoPanel_DashletName;
	
	@FindBy(xpath="//*[@id='Variable1Trans']")
	public static WebElement InfoPanel_VariableNameSelect;
	
	@FindBy(xpath="//*[@id='ddlVariablePeriod1Trans']")
	public static WebElement InfoPanel_PeriodSelect;
	
	@FindBy(xpath="(//a[contains(text(),'Setting')])[2]")
	public static WebElement InfoPanel_SettingsTab;
	
	@FindBy(xpath="//*[@id='div_filter']//div[4]//*[@id='4' and //a[text()='Warehouse']]")
	public static WebElement InfoPanel_FilterWarehouseChkBox;
	
	@FindBy(xpath="//*[@id='div_filter']//div[4]//*[@id='4' and //a[text()='Warehouse']]//..//span")
	public static WebElement InfoPanel_FilterWarehouseChkBoxSelected;
	
	@FindBy(xpath="//*[@id='div_filter']//div[3]//*[@id='2' and //label[text()='Item']]")
	public static WebElement InfoPanel_FilterItemChkBox;
	
	@FindBy(xpath="//*[@id='div_filter']//div[3]//*[@id='2' and //label[text()='Item']]//..//span")
	public static WebElement InfoPanel_FilterItemChkBoxSelected;
	
	@FindBy(xpath="//*[@id='div_refresh']//div[4]//*[@id='4' and //a[text()='Warehouse']]")
	public static WebElement InfoPanel_RefreshWarehouseChkBox;
	
	@FindBy(xpath="//*[@id='div_refresh']//div[4]//*[@id='4' and //a[text()='Warehouse']]//..//span")
	public static WebElement InfoPanel_RefreshWarehouseChkBoxSelected;
	
	@FindBy(xpath="//*[@id='div_refresh']//div[3]//*[@id='2' and //label[text()='Item']]")
	public static WebElement InfoPanel_RefreshItemChkBox;
	
	@FindBy(xpath="//*[@id='div_refresh']//div[3]//*[@id='2' and //label[text()='Item']]//..//span")
	public static WebElement InfoPanel_RefreshItemChkBoxSelected;
	
	@FindBy(xpath="(//*[@id='SaveDashlet_Icon'])[1]")
	public static WebElement InfoPanel_NewInfoPanelSaveBtn;
	
	
	@FindBy(xpath="//*[@id='Trans_Dash_Save']")
	public static WebElement InfoPanel_SaveBtn;
	
	
	@FindBy(xpath="(//*[contains(text(),'Stock Quantity')])[1]")
	public static WebElement InfoPanel_createdDashlet;
	
	
	
	@FindBy(xpath="//*[@id='CustumizeDivTrans']")
	public static WebElement InfoPanel_DraggableArea;
	
	
	@FindBy(xpath="//*[@id='CustumizeDivTrans']/div")
	public static WebElement InfoPanel_DashletDrag;
	
	@FindBy(xpath="//*[@title='Average Rate']")
	public static WebElement InfoPanel_createdDashletAvgRate;
	
	
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
	        
	        Thread.sleep(1000);

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

	        /*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
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
	
	@FindBy(xpath="//li[@title='Definition']/a")
	public static WebElement defitionTab;
	
	@FindBy(xpath="//*[@id='id_transactionentry_float_button']//tr[1]")
	public static WebElement floatBtn;
	
	@FindBy(xpath="//*[@id='id_transentry_infopanel_title']")
	public static WebElement infobarTitle;
	
	public static boolean checkEnableDoNotUseRealTimeRate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException, AWTException
	{	
		prongHornStopAtAdminLevel();
		Thread.sleep(3000);
		
		ClickUsingJs(SettingsMenu);
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Preferences));
		Preferences.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Inventory));
		Inventory.click();
		Thread.sleep(2000);
		
		if(CheckDoNotUseRealTimeRate.isSelected()==false)
		{
			ClickUsingJs(DoNotUseRealTimeRate);
			Thread.sleep(2000);
			getAlert().accept();
			
		}
		boolean actDonotUseRealTimeRateChkBOx=CheckDoNotUseRealTimeRate.isSelected();
		boolean expDonotUseRealTimeRateChkBOx=true;
		
		System.err.println("DonotUseRealTimeRateChkBOx Status: "+actDonotUseRealTimeRateChkBOx+"---------"+expDonotUseRealTimeRateChkBOx);
		
		
		Thread.sleep(1000);
		ClickUsingJs(updateBtn);

		getWaitForAlert();

		getAlert().accept();
		
		String expMessage2 = "Data saved Successfully";

		String actMessage2 = checkValidationMessage(expMessage2);

		System.err.println("********* Message  Text: " + actMessage2 + "  value expected  " + expMessage2);
		
		prongHornStartAtAdminLevel();
		
		if(actDonotUseRealTimeRateChkBOx == expDonotUseRealTimeRateChkBOx) 
		{
			return true;
		}
		
		else
		{
		return false;
		}
	}
	
	
	public boolean checkCreatingInfoPanelinOpeningStockVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();

		
		Thread.sleep(10000);
		
		//new WebDriverWait(getDriver(), 300).until(ExpectedConditions.visibilityOf(newBtn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		
		Thread.sleep(6000);
		
		if(infobarTitle.isDisplayed()==false)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(floatBtn));
			floatBtn.click();
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocks_InfoPanel));
		openingStocks_InfoPanel.click();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InfoPanelAddBtn));
		InfoPanelAddBtn.click();
		Thread.sleep(4000);
		
		//click(InfoPanel_NewInfoPanel);
		
		System.out.println(InfoPanel_NewInfoPanel.isDisplayed());
		
		click(InfoPanel_NewInfoPanel);
		Thread.sleep(10000);
		
		
		click(InfoPanel_DashletName);
		InfoPanel_DashletName.sendKeys("Stock Quantity");
		Thread.sleep(2000);
		InfoPanel_DashletName.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InfoPanel_VariableNameSelect));
		InfoPanel_VariableNameSelect.click();
		Thread.sleep(2000);
		
		InfoPanel_VariableNameSelect.sendKeys("Stock Balance");
		Thread.sleep(10000);
		
			
		/*Select s1=new Select(InfoPanel_PeriodSelect);
		s1.selectByVisibleText("As On Today");
		Thread.sleep(6000);*/
		
		click(InfoPanel_SettingsTab);
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InfoPanel_FilterWarehouseChkBoxSelected));
		
		if(InfoPanel_FilterWarehouseChkBox.isSelected()==false)
		{
			InfoPanel_FilterWarehouseChkBoxSelected.click();
		}
		
		Thread.sleep(2000);
		
		
		if(InfoPanel_RefreshWarehouseChkBox.isSelected()==false)
		{
			InfoPanel_RefreshWarehouseChkBoxSelected.click();
		}
		
		
		click(InfoPanel_NewInfoPanelSaveBtn);
		Thread.sleep(2000);
		
		String expMessage="Dashlet Saved Successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
	
		if(actMessage.equalsIgnoreCase(actMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	public boolean checkDragandDropofCreatedInfoPanel() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		Thread.sleep(4000);
		 Actions builder = new Actions(getDriver());
		 
		 Action dragAndDrop = builder.clickAndHold(InfoPanel_createdDashlet).build();
		 Thread.sleep(4000);
		  
		 dragAndDrop.perform();
					 
		 builder.dragAndDropBy(InfoPanel_DraggableArea, -100,-100).build().perform();
		 
		 Thread.sleep(2000);
		 
		
		 
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(InfoPanel_SaveBtn));
		 InfoPanel_SaveBtn.click();	

		 String expMessage = "Dashboard Layout Saved Successfully.";
		 String actMessage = checkValidationMessage(expMessage);
		 
		 Thread.sleep(2000);
		 
		 if(actMessage.equalsIgnoreCase(expMessage))
		 {
		 
			 return true;
		 }
		 else
		 {
			 return false;
		 }
		
	}
	@FindBy(xpath="//input[@id='id_body_37']")
	public static WebElement enter_Expirydate;
	
	public boolean checkSavingOpeningStockVoucher1WithBatchItem() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

	Thread.sleep(5000);
	

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	warehouseTxt.click();
	Thread.sleep(2000);
	warehouseTxt.sendKeys("HYDERABAD");
	Thread.sleep(2000);
	warehouseTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.sendKeys(Keys.SPACE);
	enter_ItemTxt.sendKeys("BR COGS ITEM");
	Thread.sleep(4000);

	enter_ItemTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
	enter_UnitTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
	enter_Quantity.sendKeys("1");
	enter_Quantity.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.sendKeys("5");
	enter_Rate.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
	enter_Gross.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Batch));
	enter_Batch.sendKeys("Batch#OSVR1");
	enter_Batch.sendKeys(Keys.TAB);

	DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
	Date date=new Date();

	Calendar cal=Calendar.getInstance();
	SimpleDateFormat currentDate = new SimpleDateFormat("dd/MM/yyyy");
	String presentDate = currentDate.format(cal.getTime());
	System.out.println("--------------Date Before Adding:: " + presentDate);

	cal.add(Calendar.DAY_OF_WEEK, 5); 

	String nextFiveDate=df.format(cal.getTime());

	System.out.println("--------------Date After Adding:: " + nextFiveDate);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
	enter_Expirydate.sendKeys(Keys.HOME,nextFiveDate);
	enter_Expirydate.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
	voucherSaveBtn.click();

	
	String expMessage1 = "Voucher saved successfully";
	
	String actMessage = checkValidationMessage(expMessage1);
	String expMessage2 = ": 1";
	
	
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
	
	
	@FindBy(xpath="//*[@id='DBdashlate0']//div//span")
	public static WebElement infopanel_QtyTxt;
	
	@FindBy(xpath="//*[@id='id_transentry_infopanel_dashletbtn_0']")
	public static WebElement infopanel_StockQtyExpandBtn;
	
	public boolean checkQuantityValueinInfoPanelinOpeningStocksVoucher() throws InterruptedException
	{
		Thread.sleep(4000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();

		
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		
		Thread.sleep(6000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.click();
		warehouseTxt.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits="Dozs";


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity="1.00";
		
		System.out.println(isElementPresent(By.xpath("//*[@id='id_transactionentry_float_button']/i")));
		
	/*	if(isElementPresent(By.xpath("//*[@id='id_transactionentry_float_button']/i")))
		{
			floatBtn.click();
			
		}*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(infopanel_StockQtyExpandBtn));
		infopanel_StockQtyExpandBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(infopanel_QtyTxt));
		String actQtyinInfoPanel=infopanel_QtyTxt.getText();
		String expQtyinInfoPanel="12.00";
		
		System.out.println("Units		"		+		actUnits		+	"	Expected	"	+	expUnits);
		
		System.out.println("Quantity	"		+		actQuantity		+	"	Expected	"	+	expQuantity);
		System.out.println("Quantity in Info Panel	in Pcs	"		+	actQtyinInfoPanel		+	"	Expected		"		+	expQtyinInfoPanel);
		
		if(actQtyinInfoPanel.equalsIgnoreCase(expQtyinInfoPanel))
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	
	@FindBy(xpath="//span[@class='icon-font6 no_padding_left_right icon-custamize']")
	public static WebElement openingStocks_InfoPanel1;
	
	public boolean checkCreatingInfoPanelforAvgRateinOpeningStocksVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocks_InfoPanel));
		openingStocks_InfoPanel.click();
		Thread.sleep(4000);
		
		
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InfoPanelAddBtn)); 
		  InfoPanelAddBtn.click(); 
		  Thread.sleep(2000);
		 
		
		click(InfoPanel_NewInfoPanel);
		Thread.sleep(6000);
		
		
		click(InfoPanel_DashletName);
		InfoPanel_DashletName.sendKeys("Average Rate");
		Thread.sleep(2000);
		InfoPanel_DashletName.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InfoPanel_VariableNameSelect));
		InfoPanel_VariableNameSelect.click();
		Thread.sleep(2000);
		
		InfoPanel_VariableNameSelect.sendKeys("Average Rate");
		Thread.sleep(8000);
		
		/*InfoPanel_VariableNameSelect.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);*/
		
		
		Select s1=new Select(InfoPanel_PeriodSelect);
		s1.selectByVisibleText("As On Today");
		Thread.sleep(2000);
		
		click(InfoPanel_SettingsTab);
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(InfoPanel_FilterItemChkBoxSelected));
		
		if(InfoPanel_FilterItemChkBox.isSelected()==false)
		{
			InfoPanel_FilterItemChkBoxSelected.click();
		}
		
		Thread.sleep(2000);
		
		
		if(InfoPanel_RefreshItemChkBox.isSelected()==false)
		{
			InfoPanel_RefreshItemChkBoxSelected.click();
		}
		
		
		click(InfoPanel_NewInfoPanelSaveBtn);
		Thread.sleep(2000);
		
		String expMessage="Dashlet Saved Successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
	
		if(actMessage.equalsIgnoreCase(actMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	public boolean checkDragandDropofCreatedInfoPanelofAvgRate() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

		
		Thread.sleep(4000);
		 Actions builder = new Actions(getDriver());
		 
		 Action dragAndDrop = builder.clickAndHold(InfoPanel_createdDashletAvgRate).build();
		 Thread.sleep(4000);
		  
		 dragAndDrop.perform();
					 
		 builder.dragAndDropBy(InfoPanel_DraggableArea, 50,50).build().perform();
		 
		 Thread.sleep(2000);
		 
		
		 
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(InfoPanel_SaveBtn));
		 InfoPanel_SaveBtn.click();	

		 String expMessage = "Dashboard Layout Saved Successfully.";
		 String actMessage = checkValidationMessage(expMessage);
		 
		 Thread.sleep(2000);
		 
		 if(actMessage.equalsIgnoreCase(expMessage))
		 {
		 
			 return true;
		 }
		 else
		 {
			 return false;
		 }
		
	
	}
	
	@FindBy(xpath="//tfoot[@id='id_transaction_bins_grid_foot']//td[13]")
	private static WebElement  binRowTotalToBeAllocated;
	
	
	public boolean checkSavingOpeningStockVoucherWithBinItemWithAutoAllocationOption() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		Thread.sleep(6000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_newBtn));
		new_newBtn.click();
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		warehouseTxt.click();
		warehouseTxt.sendKeys(Keys.SPACE);

		int warehouselist=warehouseHeaderComboList.size();

		for (int i = 0; i < warehouselist; i++) 
		{
			String warehouse=warehouseHeaderComboList.get(i).getText();

			if (warehouse.equalsIgnoreCase("HYDERABAD")) 
			{
				warehouseHeaderComboList.get(i).click();
				warehouseTxt.sendKeys(Keys.TAB);
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.SPACE);
		
		enter_ItemTxt.sendKeys("STD Rate Cogs Item");
		Thread.sleep(4000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys("10");
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys("12.69");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.SPACE);
		
		enter_ItemTxt.sendKeys("STD Rate Cogs Item");
		Thread.sleep(4000);
		enter_ItemTxt.sendKeys(Keys.TAB);

		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys("5");
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys("15");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": 2";
		
		Thread.sleep(2000);
		
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

	
	@FindBy(xpath="(//*[@id='id_transentry_infopanel_dashlet_container']//div[2]//span)[2]")
	public static WebElement infopanel_AvgRateTxt;
	
	@FindBy(xpath="//*[@id='id_transentry_infopanel_dashletbtn_1']")
	public static WebElement infopanel_AvgRateExpandBtn;

	public boolean checkAverageRateinInfoPanelinOpeningStockVoucher() throws InterruptedException
	{

		Thread.sleep(4000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
		inventoryTransactionsStocksMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksVoucher));
		openingStocksVoucher.click();

		
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
		previousBtn.click();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		enter_ItemTxt.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		float actQuantity=Float.valueOf(select1stRow_3rdColumn.getText()).floatValue();
		float expQuantity=0;

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		float actGross=Float.valueOf(select1stRow_5thColumn.getText()).floatValue();
		float expGross=0;
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		float actQuantity1=Float.valueOf(select2ndRow_3rdColumn.getText()).floatValue();
		float expQuantity1=0;

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		float actGross1=Float.valueOf(select2ndRow_5thColumn.getText()).floatValue();
		float expGross1=0;
		
		float totalQty=actQuantity+actQuantity1;
		float totalGross=actGross+actGross1;
		
		//System.out.println(isElementPresent(By.xpath("//*[@id='id_transactionentry_float_button']/i")));
		
		/*if(isElementPresent(By.xpath("//*[@id='id_transactionentry_float_button']/i")))
		{
			floatBtn.click();
			
		}*/
		
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(infopanel_AvgRateExpandBtn));
		infopanel_AvgRateExpandBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(infopanel_AvgRateTxt));
		String actAvgRateinInfoPanel=infopanel_AvgRateTxt.getText();
		float AvgRate=totalGross/(totalQty*12);
		DecimalFormat f = new DecimalFormat("##.00");
			
		String expAvgRateinInfoPanel=f.format(AvgRate);
		
		
		System.out.println("Total Quantity		"		+		totalQty*12);
		
		System.out.println("Total Gross			"		+		totalGross);
		System.out.println("Average Rate in Info Panel		"		+	actAvgRateinInfoPanel		+	"	Expected		"		+	expAvgRateinInfoPanel);
		
		if(actAvgRateinInfoPanel.equalsIgnoreCase(expAvgRateinInfoPanel))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	public boolean checkLogoutInfoPanelPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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

	
	public InfoPanelPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}

}
