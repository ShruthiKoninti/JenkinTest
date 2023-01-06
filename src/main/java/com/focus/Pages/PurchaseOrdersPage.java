package com.focus.Pages;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.server.handler.GetAllCookies;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class PurchaseOrdersPage extends BaseEngine 
{
	private static String xlfile = getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlSheetName = "PurchaseOrdersPage";
	
	@FindBy(xpath="//*[@id='btnMoreViews']")
	public static WebElement moreViews;
	
	public static boolean checkPurchasesOrderVoucherToDisplayLinksWithoutAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();

		Thread.sleep(2000);

		checkDeleteLinkStatus();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(1000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(moreViews));
		moreViews.click();
		Thread.sleep(1000);
		*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pendingMaterialRequistion));
		pendingMaterialRequistion.click();

		Thread.sleep(2000);

		String actPendingLinksList = Boolean.toString(voucherBodyGrid.getText().isEmpty());
		String expPendingLinksList = excelReader.getCellData(xlSheetName, 8, 6);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actPendingLinksList);

		System.out.println("Home page click on Pending Links Value Actual : " + actPendingLinksList  +   " Value Expected :  " + expPendingLinksList);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
		poVoucherHomeCloseBtn.click();

		if(actPendingLinksList.equalsIgnoreCase(expPendingLinksList))		
		{
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resFail);
			return false;
		}
	}



	@FindBy(xpath="//*[@id='ddlCompany']")
	private static WebElement companyList;




	public static boolean checkLogoutAndLoginWithUserAllOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 9, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 10, 5);
		
		String compname="Automation Company";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);
		
		 Select s1=new Select(companyList);
	        s1.selectByValue("36");
		

		lp.clickOnSignInBtn();
		
		Thread.sleep(2000);
		
		reLogin(unamelt, pawslt, compname);

		//checkRefershPopOnlogin();

		//checkPopUpWindow();

		Thread.sleep(6000);

		String actUserInfo=userNameUserDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameUserDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		companyLogo.click();


		String expUserInfo           =excelReader.getCellData(xlSheetName, 9, 6);
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actUserInfo);
		
		String expLoginCompanyName   =excelReader.getCellData(xlSheetName, 10, 6);
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, getLoginCompanyName);
		
		String actDashboard = Boolean.toString(dashboard.isDisplayed());
		String expDashboard = excelReader.getCellData(xlSheetName, 11, 6);
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actDashboard);

		System.out.println("**********************************************checkLogoutAndLoginWithUserAllOptions*********************************************");

		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard.equalsIgnoreCase(expDashboard)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 9, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 9, 8, resFail);
			return false;
		}
	}







	public static boolean checkAuthorizationInMaterialRequisitionWithUserAllOptionsST() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		/*getWebDriverWait().until(ExpectedConditions.visibilityOf(noOfAlertsDisplay));
		noOfAlertsDisplay.click();

		// Actual
		String  actAlertsMainHeader=alertsMainHeader.getText();

		boolean actNoOfAlertsDisplay=alertsMainHeader.isDisplayed();
		boolean expNoOfAlertsDisplay=true;

		System.out.println("Alerts is display"+ actAlertsMainHeader);

		String actAlertTransactionsTreeView=alertTransaction.getText();
		String expAlertTransactionsTreeView="  Transactions (1)";

		alertTransaction.click();

		String actBudgetAuthorization = transactionAuthorization.getText();
		String expBudgetAuthorization = " Material Requisition(s) Require Approval (1)";

		transactionAuthorization.click();

		String actBudgetText =transactionAuthorizationText.getText();
		String expBudgetText = "Voucher needs your authorization. Doc No: MatReq:1";
		
		Thread.sleep(2000);

		transactionAuthorizationText.click();

		Thread.sleep(3000);*/

		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
  		inventoryMenu.click();

  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
  		inventoryTransactionsMenu.click();

  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsStocksMenu));
  		inventoryTransactionsStocksMenu.click();

  		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(materialRequisitionVoucher));
  		materialRequisitionVoucher.click();

  		Thread.sleep(2000);

		getAction().doubleClick(grid_ChkBox1).build().perform();

		checkValidationMessage("Voucher loaded successfully");

		Thread.sleep(3000);

		/*System.out.println("************************************* checkAuthorizationInPurchaseOrderWithUserAllOptionsST  *********************************");
		System.out.println("noOfAlertsDisplay         : " + actNoOfAlertsDisplay         + " Value Expected  : " + expNoOfAlertsDisplay);
		System.out.println("noOfAlertsDisplay         : " + actAlertTransactionsTreeView + " Value Expected  : " + expAlertTransactionsTreeView);
		System.out.println("noOfAlertsDisplay         : " + actBudgetAuthorization       + " Value Expected  : " + expBudgetAuthorization);
		System.out.println("noOfAlertsDisplay         : " + actBudgetText                + " Value Expected  : " + expBudgetText);*/


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		/*MRpurchaseAccountTxt.click();*/

		String actPurchaseValue = MRpurchaseAccountTxt.getAttribute("value");
		String expPurchaseValue = excelReader.getCellData(xlSheetName, 12, 6);
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actPurchaseValue);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		/*departmentTxt.click();*/

		String actDepartmentValue = departmentTxt.getAttribute("value");
		String expDepartmentValue = excelReader.getCellData(xlSheetName, 13, 6);
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actDepartmentValue);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));

		String actR1C1Item =select1stRow_1stColumn.getText();
		String actR1C2Units=select1stRow_2ndColumn.getText();
		String actR1C5Qty  =select1stRow_3rdColumn.getText();
		String actR1C7Rate =select1stRow_4thColumn.getText();
		String actR1C8Gross=select1stRow_5thColumn.getText();

		String expR1C1Item =excelReader.getCellData(xlSheetName, 14, 6);
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actR1C1Item);
		
		String expR1C2Units=excelReader.getCellData(xlSheetName, 15, 6);
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actR1C2Units);
		
		String expR1C5Qty  =excelReader.getCellData(xlSheetName, 16, 6);
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actR1C5Qty);
		
		String expR1C7Rate =excelReader.getCellData(xlSheetName, 17, 6);
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actR1C7Rate);
		
		String expR1C8Gross=excelReader.getCellData(xlSheetName, 18, 6);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actR1C8Gross);

		System.out.println("Entry Page Save with Vendor Item Value Actual   : "+actPurchaseValue   +" Value Expected : "+expPurchaseValue);	
		System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actDepartmentValue  +" Value Expected : "+actDepartmentValue);
		System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C1Item  +" Value Expected : "+expR1C1Item);			
		System.out.println("Entry Page Save with Vendor Units Value Actual  : "+actR1C2Units  +" Value Expected : "+expR1C2Units);
		System.out.println("Entry Page Save with Vendor QTY Value Actual    : "+actR1C5Qty    +" Value Expected : "+expR1C5Qty);
		System.out.println("Entry Page Save with Vendor RATE Value Actual   : "+actR1C7Rate   +" Value Expected : "+expR1C7Rate);
		System.out.println("Entry Page Save with Vendor Gross Value Actual  : "+actR1C8Gross  +" Value Expected : "+expR1C8Gross);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(POauthorizeBtn));
		POauthorizeBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksNewHomeRow1AuthorizationStatus));

		String actStatus=openingStocksNewHomeRow1AuthorizationStatus.getText();
		String expStatus=excelReader.getCellData(xlSheetName, 19, 6);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actStatus);

		System.out.println("Authorization Status In home Page 	 :  "+actStatus +" Value Expected : "+expStatus);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
		poVoucherHomeCloseBtn.click();

		if(/*actNoOfAlertsDisplay==expNoOfAlertsDisplay && actAlertTransactionsTreeView.equalsIgnoreCase(expAlertTransactionsTreeView)
				&& actBudgetAuthorization.equalsIgnoreCase(expBudgetAuthorization) && actBudgetText.equalsIgnoreCase(expBudgetText)

				&&*/ actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units) && actR1C5Qty.equalsIgnoreCase(expR1C5Qty) 
				&& actR1C7Rate.equalsIgnoreCase(expR1C7Rate) && actR1C8Gross.equalsIgnoreCase(expR1C8Gross)

				&& actStatus.equalsIgnoreCase(expStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 8, resFail);
			return false;
		}
	}


@FindBy(xpath="//*[@id='id_mainlayoutmenu']/ul[2]/li[6]/div/i")
public static WebElement downArrow;





	public static boolean checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(downArrow));
		downArrow.click();
		Thread.sleep(2000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 20, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 21, 5);
		
		String compname="Automation Company";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);
		
		
		Select s1=new Select(companyList);
        s1.selectByValue("36");
	

		lp.clickOnSignInBtn();
		
		Thread.sleep(2000);
		
	
		
		reLogin(unamelt, pawslt, compname);

		//checkRefershPopOnlogin();

		//checkPopUpWindow();

		Thread.sleep(6000);

		String actUserInfo=userNameDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		companyLogo.click();


		String expUserInfo           =excelReader.getCellData(xlSheetName, 20, 6);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actUserInfo);
		
		String expLoginCompanyName   =excelReader.getCellData(xlSheetName, 21, 6);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, getLoginCompanyName);
		
		String actDashboard = Boolean.toString(dashboard.isDisplayed());
		String expDashboard = excelReader.getCellData(xlSheetName, 22, 6);
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actDashboard);

		System.out.println("**********************************************checkLogoutAndLoginWithSU*********************************************");

		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard.equalsIgnoreCase(expDashboard)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 20, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 20, 8, resFail);
			return false;
		}
	}







	public static boolean checkPurchasesOrderVoucherToDisplayLinksAfterAuthorizationAndSavePurchaseOrder() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();

		Thread.sleep(1000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(moreOptionsBtn));
		moreOptionsBtn.click();
		Thread.sleep(2000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pendingMaterialRequistion));
		pendingMaterialRequistion.click();

		Thread.sleep(2000);

		String actPendingLinksList = Boolean.toString(voucherBodyGrid.getText().isEmpty());
		String expPendingLinksList = excelReader.getCellData(xlSheetName, 23, 6);
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actPendingLinksList);
		
		System.out.println("Home page click on Pending Links Value Actual : " + actPendingLinksList  +   " Value Expected :  " + expPendingLinksList);


		String actHomePageVoucherNo     = openingStocksNewHomeRow1VoucherNo.getText();
		String actHomePageCreatedBy     = openingStocksNewHomeRow1CreatedBy.getText();
		String actHomePageModifiedBy    = openingStocksNewHomeRow1ModifiedBy.getText();
		String actHomePageSuspendBy     = openingStocksNewHomeRow1SuspendStatus.getText();
		String actHomePageAuthorizedBy  = openingStocksNewHomeRow1AuthorizationStatus.getText();

		String expHomePageVoucherNo     = excelReader.getCellData(xlSheetName, 24, 6);
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actHomePageVoucherNo);
		
		String expHomePageCreatedBy     = excelReader.getCellData(xlSheetName, 25, 6);
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actHomePageCreatedBy);
		
		String expHomePageModifiedBy    = excelReader.getCellData(xlSheetName, 26, 6);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actHomePageModifiedBy);
		
		String expHomePageSuspendBy     = excelReader.getCellData(xlSheetName, 27, 6);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actHomePageSuspendBy);
		
		String expHomePageAuthorizeBy   = excelReader.getCellData(xlSheetName, 28, 6);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actHomePageAuthorizedBy);

		System.out.println("Voucher No in Home Page Value Actual   :  " + actHomePageVoucherNo+    " Value Expected : " + expHomePageVoucherNo);
		System.out.println("Created By in Home Page Value Actual   :  " + actHomePageCreatedBy+    " Value Expected : " + expHomePageCreatedBy);
		System.out.println("Modified By in Home Page Value Actual  :  " + actHomePageModifiedBy+   " Value Expected : " + expHomePageModifiedBy);
		System.out.println("Suspend By in Home Page Value Actual   :  " + actHomePageSuspendBy+    " Value Expected : " + expHomePageSuspendBy);
		System.out.println("Authorize By in Home Page Value Actual :  " + actHomePageAuthorizedBy+ " Value Expected : " + expHomePageAuthorizeBy);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
		grid_ChkBox1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertBtn));
		convertBtn.click();

		Thread.sleep(2000);

		checkValidationMessage("Link data loaded");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		String actPurchaseAccount = MRpurchaseAccountTxt.getAttribute("value");
		String actDepartment      = departmentValuetxt.getAttribute("value");

		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 29, 6);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actPurchaseAccount);
		
		String expDepartment      = excelReader.getCellData(xlSheetName, 30, 6);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actDepartment);
		
		System.out.println("Entry Page Save with Vendor Item Value Actual   : " + actPurchaseAccount   + " Value Expected : " + expPurchaseAccount);	
		System.out.println("Entry Page Save with Vendor Units Value Actual  : " + actDepartment        + " Value Expected : " + expDepartment);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));

		String actR1C1Item =select1stRow_1stColumn.getText();
		String actR1C2Units=select1stRow_2ndColumn.getText();
		String actR1C5Qty  =select1stRow_3rdColumn.getText();
		String actR1C7Link =select1stRow_4thColumn.getText();
		String actR1C8Rate=select1stRow_5thColumn.getText();
		String actR1C9Gross=select1stRow_6thColumn.getText();

		String expR1C1Item =excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actR1C1Item);
		
		String expR1C2Units=excelReader.getCellData(xlSheetName, 32, 6);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actR1C2Units);
		
		String expR1C5Qty  =excelReader.getCellData(xlSheetName, 33, 6);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actR1C5Qty);
		
		String expR1C7Link =excelReader.getCellData(xlSheetName, 34, 6);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actR1C7Link);
		
		String expR1C8Rate =excelReader.getCellData(xlSheetName, 35, 6);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actR1C8Rate);
		
		String expR1C9Gross=excelReader.getCellData(xlSheetName, 36, 6);
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actR1C9Gross);

		System.out.println("Entry Page Save with Vendor Item Value Actual   : " + actR1C1Item   + " Value Expected : " + expR1C1Item);	
		System.out.println("Entry Page Save with Vendor Units Value Actual  : " + actR1C2Units  + " Value Expected : " + expR1C2Units);
		System.out.println("Entry Page Save with Vendor QTY Value Actual    : " + actR1C5Qty    + " Value Expected : " + expR1C5Qty);
		System.out.println("Entry Page Save with Vendor RATE Value Actual   : " + actR1C7Link   + " Value Expected : " + expR1C7Link);
		System.out.println("Entry Page Save with Vendor Gross Value Actual  : " + actR1C8Rate   + " Value Expected : " + expR1C8Rate);			
		System.out.println("Entry Page Save with Vendor Gross Value Actual  : " + actR1C9Gross  + " Value Expected : " + expR1C9Gross);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 23, 5));
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);


		if (actR1C5Qty.equalsIgnoreCase(expR1C5Qty)==false) 
		{
			select1stRow_3rdColumn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
			enter_Quantity.sendKeys(Keys.END);
			enter_Quantity.sendKeys(Keys.SHIFT, Keys.HOME);
			enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 24, 5));
			enter_Quantity.sendKeys(Keys.TAB);
			Thread.sleep(2000);
		}

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		//Thread.sleep(4000);

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(1);

		HashSet<String> expMsg = new HashSet();

		//expMsg.add(excelReader.getCellData(xlSheetName, 37, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 38, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 37, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 38, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		
		Thread.sleep(2000);

		if(actPendingLinksList.startsWith(expPendingLinksList)
				&& actHomePageVoucherNo.equalsIgnoreCase(expHomePageVoucherNo) && actHomePageCreatedBy.equalsIgnoreCase(expHomePageCreatedBy)
				&& actHomePageModifiedBy.equalsIgnoreCase(expHomePageModifiedBy) && actHomePageSuspendBy.equalsIgnoreCase(expHomePageSuspendBy)
				&& actHomePageAuthorizedBy.equalsIgnoreCase(expHomePageAuthorizeBy)

				&& actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) && actDepartment.equalsIgnoreCase(expDepartment)
				&& actR1C1Item.equalsIgnoreCase(expR1C1Item) && actR1C2Units.equalsIgnoreCase(expR1C2Units)
				&& actR1C5Qty.equalsIgnoreCase(expR1C5Qty) && actR1C7Link.equalsIgnoreCase(expR1C7Link)
				&& actR1C8Rate.equalsIgnoreCase(expR1C8Rate) && actR1C9Gross.equalsIgnoreCase(expR1C9Gross)

				&&actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))		
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			
			excelReader.setCellData(xlfile, xlSheetName, 23, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();

			excelReader.setCellData(xlfile, xlSheetName, 23, 8, resFail);
			return false;
		}
	}







	public static boolean checkLogoutAndLoginWithUserAllOptionsSTToAuthorizePurchaseOrder() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(downArrow));
		downArrow.click();
		Thread.sleep(2000);
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 39, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 40, 5);
		String compname="Automation company";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		lp.clickOnSignInBtn();
		Thread.sleep(2000);

		reLogin(unamelt, pawslt, compname);
		//checkRefershPopOnlogin();

		//checkPopUpWindow();

		Thread.sleep(5000);

		String actUserInfo=userNameUserDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameUserDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		companyLogo.click();


		String expUserInfo           =excelReader.getCellData(xlSheetName, 39, 6);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actUserInfo);
		
		String expLoginCompanyName   =excelReader.getCellData(xlSheetName, 40, 6);
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, getLoginCompanyName);
		
		String actDashboard = Boolean.toString(dashboard.isDisplayed());
		String expDashboard = excelReader.getCellData(xlSheetName, 41, 6);
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actDashboard);
		
		getDriver().navigate().refresh();

		Thread.sleep(2000);

		System.out.println("**********************************************checkLogoutAndLoginWithUserAllOptionsSTToAuthorizePurchaseOrder*********************************************");

		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard.equalsIgnoreCase(expDashboard)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 39, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 39, 8, resFail);
			return false;
		}
	}








	public static boolean checkAuthorizationInPurchaseOrderWithUserAllOptionsST() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		/*getWebDriverWait().until(ExpectedConditions.visibilityOf(noOfAlertsDisplay));
		noOfAlertsDisplay.click();

		getWebDriverWait().until(ExpectedConditions.visibilityOf(alertTransaction));
		alertTransaction.click();

		getWebDriverWait().until(ExpectedConditions.visibilityOf(transactionAuthorization));
		transactionAuthorization.click();

		getWebDriverWait().until(ExpectedConditions.visibilityOf(transactionAuthorizationText));
		transactionAuthorizationText.click();

		Thread.sleep(5000);

		checkValidationMessage("Voucher loaded Successfully");

		Thread.sleep(3000);*/
		
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
		getAction().doubleClick(grid_ChkBox1).build().perform();

		checkValidationMessage("Voucher loaded Successfully");
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		String actPurchaseAccount = MRpurchaseAccountTxt.getAttribute("value");
		String actVendorAccount   = vendorAccountTxt.getAttribute("value");
		String actDepartment      = departmentValuetxt.getAttribute("value");

		String expPurchaseAccount = excelReader.getCellData(xlSheetName, 42, 6);
		excelReader.setCellData(xlfile, xlSheetName, 42, 7, actPurchaseAccount);
		
		String expVendorAccount   = excelReader.getCellData(xlSheetName, 43, 6);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actVendorAccount);
		
		String expDepartment      = excelReader.getCellData(xlSheetName, 44, 6);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actDepartment);
		
		System.out.println("Entry Page Save with Vendor Item Value Actual   : " + actPurchaseAccount   + " Value Expected : " + expPurchaseAccount);	
		System.out.println("Entry Page Save with Vendor Account Value Actual: " + actVendorAccount     + " Value Expected : " + expVendorAccount);	
		System.out.println("Entry Page Save with Vendor Units Value Actual  : " + actDepartment        + " Value Expected : " + expDepartment);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));

		String actR1C1Item =select1stRow_1stColumn.getText();
		String actR1C2Units=select1stRow_2ndColumn.getText();
		String actR1C5Qty  =select1stRow_3rdColumn.getText();
		String actR1C7Link =select1stRow_4thColumn.getText();
		String actR1C8Rate=select1stRow_5thColumn.getText();
		String actR1C9Gross=select1stRow_6thColumn.getText();

		String expR1C1Item =excelReader.getCellData(xlSheetName, 45, 6);
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actR1C1Item);
		
		String expR1C2Units=excelReader.getCellData(xlSheetName, 46, 6);
		excelReader.setCellData(xlfile, xlSheetName, 46, 7, actR1C2Units);
		
		String expR1C5Qty  =excelReader.getCellData(xlSheetName, 47, 6);
		excelReader.setCellData(xlfile, xlSheetName, 47, 7, actR1C5Qty);
		
		String expR1C7Link =excelReader.getCellData(xlSheetName, 48, 6);
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actR1C7Link);
		
		String expR1C8Rate=excelReader.getCellData(xlSheetName, 49, 6);
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actR1C8Rate);
		
		String expR1C9Gross=excelReader.getCellData(xlSheetName, 50, 6);
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actR1C9Gross);
		
		System.out.println("Entry Page Save with Vendor Item Value Actual   : " + actR1C1Item   + " Value Expected : " + expR1C1Item);	
		System.out.println("Entry Page Save with Vendor Units Value Actual  : " + actR1C2Units  + " Value Expected : " + expR1C2Units);
		System.out.println("Entry Page Save with Vendor QTY Value Actual    : " + actR1C5Qty    + " Value Expected : " + expR1C5Qty);
		System.out.println("Entry Page Save with Vendor RATE Value Actual   : " + actR1C7Link   + " Value Expected : " + expR1C7Link);
		System.out.println("Entry Page Save with Vendor Gross Value Actual  : " + actR1C8Rate   + " Value Expected : " + expR1C8Rate);			
		System.out.println("Entry Page Save with Vendor Gross Value Actual  : " + actR1C9Gross  + " Value Expected : " + expR1C9Gross);

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(POauthorizeBtn));
		POauthorizeBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksNewHomeRow1AuthorizationStatus));

		String actStatus=openingStocksNewHomeRow1AuthorizationStatus.getText();
		String expStatus=excelReader.getCellData(xlSheetName, 51, 6);
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actStatus);

		if (actStatus.equalsIgnoreCase(expStatus)==false) 
		{
			getDriver().navigate().refresh();

			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
			inventoryTransactionsMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
			inventoryTransactionsPurchasesMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
			purchasesOrdersVoucher.click();

			Thread.sleep(3000);

			if (actStatus.equalsIgnoreCase(expStatus)==false) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
				grid_ChkBox1.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(authorizeBtn));
				authorizeBtn.click();
				
				Thread.sleep(2000);
			}
			actStatus=openingStocksNewHomeRow1AuthorizationStatus.getText();
		}

		System.out.println("Authorization Status In home Page 	 :  "+actStatus +" Value Expected : "+expStatus);

		if(actPurchaseAccount.equalsIgnoreCase(expPurchaseAccount) && actVendorAccount.equalsIgnoreCase(expVendorAccount) 
				&& actDepartment.equalsIgnoreCase(expDepartment) && actR1C1Item.equalsIgnoreCase(expR1C1Item) 
				&& actR1C2Units.equalsIgnoreCase(expR1C2Units) && actR1C5Qty.equalsIgnoreCase(expR1C5Qty) 
				&& actR1C7Link.equalsIgnoreCase(expR1C7Link) && actR1C8Rate.equalsIgnoreCase(expR1C8Rate) 
				&& actR1C9Gross.equalsIgnoreCase(expR1C9Gross)

				&& actStatus.equalsIgnoreCase(expStatus))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resFail);
			return false;
		}
	}








	public static boolean checkLogoutAndLoginWithSUForPurchasesVoucherVAT() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 52, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 53, 5);
		String compname="Automation Company";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		lp.clickOnSignInBtn();
		Thread.sleep(2000);

		reLogin(unamelt, pawslt, compname);
		//checkRefershPopOnlogin();

		//checkPopUpWindow();

		Thread.sleep(6000);

		String actUserInfo=userNameDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		companyLogo.click();


		String expUserInfo           =excelReader.getCellData(xlSheetName, 52, 6);
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actUserInfo);
		
		String expLoginCompanyName   =excelReader.getCellData(xlSheetName, 53, 6);
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, getLoginCompanyName);
		
		String actDashboard = Boolean.toString(dashboard.isDisplayed());
		String expDashboard = excelReader.getCellData(xlSheetName, 54, 6);
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actDashboard);

		System.out.println("**********************************************checkLogoutAndLoginWithSUForPurchasesVoucherVAT*********************************************");

		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard.equalsIgnoreCase(expDashboard)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 52, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 52, 8, resFail);
			return false;
		}
	}







	public boolean checkVerifingDetailsOfSavedPurchaseOrdersVouchersInHomepage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		Calendar cal=Calendar.getInstance();
		SimpleDateFormat todayDate = new SimpleDateFormat("dd/MM/yyyy");
		String currentDate = todayDate.format(cal.getTime());
		System.out.println("--------------Date Before Adding:: " + currentDate);

		int count = voucherHomeBodyList.size();

		ArrayList<String>  actVouchersList= new ArrayList<String>(); 

		for (int i = 0; i < count; i++) 
		{
			String data = voucherHomeBodyList.get(i).getText();

			if (data.isEmpty()==false && i!=9 && i!=10) 
			{
				System.out.println(data);

				actVouchersList.add(data);
			}
		}

		ArrayList<String>  expVouchersList= new ArrayList<String>(); 

		expVouchersList.add(excelReader.getCellData(xlSheetName, 55, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 56, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 57, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 58, 6));
		expVouchersList.add(currentDate);
		expVouchersList.add(currentDate);
		expVouchersList.add(excelReader.getCellData(xlSheetName, 59, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 60, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 61, 6));
		expVouchersList.add(excelReader.getCellData(xlSheetName, 62, 6));
		
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actVouchersList.toString());

		System.out.println("**********************checkVerifingDetailsOfSavedOpeningStocksVouchersInHomepage*********************");

		System.out.println(actVouchersList);
		System.out.println(expVouchersList);

		if (actVouchersList.equals(expVouchersList)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 55, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 55, 8, resFail);
			return false;
		}
	}








	public static boolean checkSavingPurchasesOrderVoucher2WithTwoRows() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRpurchaseAccountTxt));
		MRpurchaseAccountTxt.click();
		MRpurchaseAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 63, 5));
		Thread.sleep(3000);
		MRpurchaseAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.click();
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 64, 5));
		Thread.sleep(3000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentValuetxt));
		departmentValuetxt.click();
		departmentValuetxt.sendKeys(excelReader.getCellData(xlSheetName, 65, 5));
		Thread.sleep(3000);
		departmentValuetxt.sendKeys(Keys.TAB);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 66, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		select1stRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 67, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		select1stRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 68, 5));
		enter_Rate.sendKeys(Keys.TAB);




		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.END);
		enter_ItemTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 69, 5));
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);	


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		select2ndRow_3rdColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 70, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
		select2ndRow_4thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 71, 5));
		enter_Rate.sendKeys(Keys.TAB);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		
		//Thread.sleep(3000);

		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 2");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 63, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 64, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);

		if(actMessage.equals(expMessage))		*/
		String expMessage1 = excelReader.getCellData(xlSheetName, 63, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 64, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		{
			excelReader.setCellData(xlfile, xlSheetName, 63, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 63, 8, resFail);
			return false;
		}
	}







	public static boolean checkLogoutAndLoginWithUserAllOptionsSTToAuthorizePurchaseOrderVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 72, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 73, 5);
		String compname="Automation Company";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		lp.clickOnSignInBtn();
		Thread.sleep(2000);
		
		reLogin(unamelt, pawslt, compname);

		//checkRefershPopOnlogin();

		//checkPopUpWindow();

		Thread.sleep(5000);

		String actUserInfo=userNameUserDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameUserDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		companyLogo.click();


		String expUserInfo           =excelReader.getCellData(xlSheetName, 72, 6);
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actUserInfo);
		
		String expLoginCompanyName   =excelReader.getCellData(xlSheetName, 73, 6);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, getLoginCompanyName);
		
		String actDashboard = Boolean.toString(dashboard.isDisplayed());
		String expDashboard = excelReader.getCellData(xlSheetName, 74, 6);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actDashboard);

		getDriver().navigate().refresh();

		Thread.sleep(2000);

		System.out.println("**********************************************checkLogoutAndLoginWithUserAllOptionsSTToAuthorizePurchaseOrderVoucher2*********************************************");

		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard.equalsIgnoreCase(expDashboard)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 72, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 72, 8, resFail);
			return false;
		}
	}



	




	public static boolean checkAuthorizationInPurchaseOrderVoucher2WithUserAllOptionsST() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		/*getWebDriverWait().until(ExpectedConditions.visibilityOf(noOfAlertsDisplay));
		noOfAlertsDisplay.click();

		getWebDriverWait().until(ExpectedConditions.visibilityOf(alertTransaction));
		alertTransaction.click();

		getWebDriverWait().until(ExpectedConditions.visibilityOf(transactionAuthorization));
		transactionAuthorization.click();

		getWebDriverWait().until(ExpectedConditions.visibilityOf(transactionAuthorizationText));
		transactionAuthorizationText.click();*/
		
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
		getAction().doubleClick(grid_ChkBox1).build().perform();

		checkValidationMessage("Voucher loaded Successfully");
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(POauthorizeBtn));
		POauthorizeBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksNewHomeRow1AuthorizationStatus));

		String actStatus=openingStocksNewHomeRow1AuthorizationStatus.getText();
		String expStatus=excelReader.getCellData(xlSheetName, 75, 6);
		excelReader.setCellData(xlfile, xlSheetName, 75, 7, actStatus);

		System.out.println("Authorization Status In home Page 	 :  "+actStatus +" Value Expected : "+expStatus);

		if(actStatus.equalsIgnoreCase(expStatus))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 75, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 75, 8, resFail);
			return false;
		}
	}







	public static boolean checkLogoutAndLoginWithSUAfterAuthorizingPurchaseOrderVoucher2() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt=excelReader.getCellData(xlSheetName, 76, 5);

		String pawslt=excelReader.getCellData(xlSheetName, 77, 5);
		String compname="Automation Company";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		lp.clickOnSignInBtn();
		Thread.sleep(2000);
		reLogin(unamelt, pawslt, compname);

		//checkRefershPopOnlogin();

		//checkPopUpWindow();

		Thread.sleep(6000);

		String actUserInfo=userNameDisplay.getText();

		System.out.println("User Info  : "+actUserInfo);

		System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  "+ getLoginCompanyName);
		companyLogo.click();


		String expUserInfo           =excelReader.getCellData(xlSheetName, 76, 6);
		String expLoginCompanyName   =excelReader.getCellData(xlSheetName, 77, 6);

		String actDashboard = Boolean.toString(dashboard.isDisplayed());
		String expDashboard = excelReader.getCellData(xlSheetName, 78, 6);

		System.out.println("***************checkLogoutAndLoginWithSUAfterAuthorizingPurchaseOrderVoucher2*************");

		System.out.println("User Info               : "+actUserInfo         +" Value Expected "+expUserInfo);
		System.out.println("Login Company           : "+getLoginCompanyName +" Value Expected "+expLoginCompanyName);
		System.out.println("Dashboard is Displaying : "+actDashboard        +" Value Expected "+expDashboard );

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName) && actDashboard.equalsIgnoreCase(expDashboard)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 76, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 76, 8, resFail);
			return false;
		}
	}









	public static boolean checkDataInPurchaseVoucherVATVoucher1ConvertedLink() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
		homePannelOpenBtn.click();
		
		Thread.sleep(1000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(moreViews));
		moreViews.click();
		Thread.sleep(2000);*/
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pendingPurchasesOrdersLink));
		pendingPurchasesOrdersLink.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int vcount = grid_VoucherNoList.size();

		for (int i = 0; i < vcount; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 79, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
					break;
				}
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(convertBtn));
		convertBtn.click();
		
		Thread.sleep(2000);

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 79, 6);
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 80, 6);
		excelReader.setCellData(xlfile, xlSheetName, 80, 7, actVendor);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 81, 6);
		excelReader.setCellData(xlfile, xlSheetName, 81, 7, actCurrency);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 82, 6);
		excelReader.setCellData(xlfile, xlSheetName, 82, 7, actDepartment);
		
		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 83, 6);
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actR1Item);

		String actR1PurchaseAccount = select1stRow_4thColumn.getText();
		String expR1PurchaseAccount = excelReader.getCellData(xlSheetName, 84, 6);
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, actR1PurchaseAccount);
		
		String actR1Units = select1stRow_5thColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 85, 6);
		excelReader.setCellData(xlfile, xlSheetName, 85, 7, actR1Units);

		String actR1Quantity = select1stRow_9thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 86, 6);
		excelReader.setCellData(xlfile, xlSheetName, 86, 7, actR1Quantity);
		
		String actR1LPurchase = select1stRow_10thColumn.getText();
		String expR1LPurchase = excelReader.getCellData(xlSheetName, 87, 6);
		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actR1LPurchase);
		
		String actR1Rate = select1stRow_11thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 88, 6);
		excelReader.setCellData(xlfile, xlSheetName, 88, 7, actR1Rate);
		
		String actR1Gross = select1stRow_12thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 89, 6);
		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actR1Gross);


		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 90, 6);
		excelReader.setCellData(xlfile, xlSheetName, 90, 7, actR2Item);
		
		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 91, 6);
		excelReader.setCellData(xlfile, xlSheetName, 91, 7, actR2PurchaseAccount);
		
		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 92, 6);
		excelReader.setCellData(xlfile, xlSheetName, 92, 7, actR2Units);

		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 93, 6);
		excelReader.setCellData(xlfile, xlSheetName, 93, 7, actR2Quantity);

		String actR2LPurchase = select2ndRow_10thColumn.getText();
		String expR2LPurchase = excelReader.getCellData(xlSheetName, 94, 6);
		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actR2LPurchase);
		
		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 95, 6);
		excelReader.setCellData(xlfile, xlSheetName, 95, 7, actR2Rate);

		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 96, 6);
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actR2Gross);
		
		System.out.println("*******************************checkSavingPurchaseVoucherVATVoucher1ConvertLink**********************************");

		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);

		System.out.println("*******************************ROW1**********************************");

		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("PurchaseAccount : "+actR1PurchaseAccount	+"  Value Expected  "+expR1PurchaseAccount);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("LPurchase       : "+actR1LPurchase			+"  Value Expected  "+expR1LPurchase);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);

		System.out.println("*******************************ROW2**********************************");

		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("PurchaseAccount : "+actR2PurchaseAccount	+"  Value Expected  "+expR2PurchaseAccount);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("LPurchase       : "+actR1LPurchase			+"  Value Expected  "+expR1LPurchase);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);

		if(actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
				&& actDepartment.equalsIgnoreCase(expDepartment)  
				&& actR1Item.equalsIgnoreCase(expR1Item) && actR1PurchaseAccount.equalsIgnoreCase(expR1PurchaseAccount) 
				&& actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1LPurchase.equalsIgnoreCase(expR1LPurchase)
				&& actR1Rate.equalsIgnoreCase(expR1Rate) && actR1Gross.equalsIgnoreCase(expR1Gross)

				&& actR2Item.equalsIgnoreCase(expR2Item) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
				&& actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2LPurchase.equalsIgnoreCase(expR2LPurchase)
				&& actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			getWaitForAlert();
			getAlert().accept();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 79, 8, resPass);
			return true;
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			getWaitForAlert();
			getAlert().accept();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 79, 8, resFail);
			return false;
		}
	}









	public boolean checkCloseRow2LinkInPurchaseOrderHavingTwoLinkDocuments() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int vcount = grid_VoucherNoList.size();

		for (int i = 0; i < vcount; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 97, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					getAction().doubleClick(grid_CheckBoxList.get(i)).build().perform();
					break;
				}
			}
		}

		Thread.sleep(2000);
		checkValidationMessage("Screen Opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(secondRowIndex));
		getAction().contextClick(secondRowIndex).build().perform();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectRowBtn));
		selectRowBtn.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleBtn));
		toggleBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeLinksOption));
		closeLinksOption.click();
		
		Thread.sleep(2000);

		String expValidationMessage = excelReader.getCellData(xlSheetName, 97, 6);
		String actValidationMessage = checkValidationMessage(expValidationMessage);
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actValidationMessage);
		
		Thread.sleep(2000);

		if (actValidationMessage.equalsIgnoreCase(expValidationMessage))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 97, 8, resPass);
			return true;
		}
		else                                                                    
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
			new_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 97, 8, resFail);
			return false;
		}
	}









	public static boolean checkSavingPurchaseVoucherVATVoucher1WithPurchaseOrderLink() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
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
		Thread.sleep(2000);

		checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 98, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 99, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 100, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 101, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 102, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 103, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 104, 5));
		Thread.sleep(2000);
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
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 105, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(5000);

		int POCount = LPurchaseOrderPopupList.size();

		ArrayList<String> PurchaseOrder = new ArrayList<String>();

		for (int i = 0; i < POCount; i++) 
		{
			if (i==0 || i==6 || i==8 || i==9 /*|| i==10*/ || i==11 || i==17 || i==19 || i==20 /*|| i==21*/) 
			{
				String data = LPurchaseOrderPopupList.get(i).getText();
				PurchaseOrder.add(data);
			}
		}

		String actRow1SuspendColumn = LpurchaseOrderPopupRow1SuspendCol.getText();
		boolean actRow1Suspend = false;
		boolean expRow1Suspend = true;

		if (actRow1SuspendColumn.isEmpty() || actRow1SuspendColumn.equalsIgnoreCase("0.00")) 
		{
			actRow1Suspend=true;
		}


		String actRow2SuspendColumn = LpurchaseOrderPopupRow2SuspendCol.getText();
		boolean actRow2Suspend = false;
		boolean expRow2Suspend = true;

		if (actRow2SuspendColumn.isEmpty() || actRow2SuspendColumn.equalsIgnoreCase("0.00")) 
		{
			actRow2Suspend=true;
		}

		String actRow1PurOrdList = PurchaseOrder.toString();
		String expRow1PurOrdList = excelReader.getCellData(xlSheetName, 98, 6);
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actRow1PurOrdList);
		
		System.err.println(actRow1PurOrdList);
		System.err.println(expRow1PurOrdList);

		int pvoccount = LPurchaseOrderPopupVoucherNos.size();

		for (int i = 0; i < pvoccount; i++) 
		{
			String data = LPurchaseOrderPopupVoucherNos.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 106, 5))) 
			{
				if (LPurchaseOrderPopupCheckBoxs.get(i).isSelected()==false) 
				{
					LPurchaseOrderPopupCheckBoxs.get(i).click();
				}
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LPurchaseOrderPopupOkBtn));
		LPurchaseOrderPopupOkBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
		select1stRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 107, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_10thColumn));
		String actRow1LPurOrd = select1stRow_10thColumn.getText();
		String expRow1LPurOrd = excelReader.getCellData(xlSheetName, 101, 6);
		excelReader.setCellData(xlfile, xlSheetName, 101, 7, actRow1LPurOrd);




		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		select2ndRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(Keys.END);
		enter_WarehouseTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 108, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 109, 5));
		Thread.sleep(2000);
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
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 110, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		int POCount1 = LPurchaseOrderPopupList.size();

		ArrayList<String> PurchaseOrder1 = new ArrayList<String>();

		for (int i = 0; i < POCount1; i++) 
		{
			if (i==0 || i==6 || i==8 || i==9 /*|| i==10*/ || i==11 || i==17 || i==19 || i==20 /*|| i==21*/) 
			{
				String data = LPurchaseOrderPopupList.get(i).getText();
				PurchaseOrder1.add(data);
			}
		}

		String actRow1SuspendColumnR2 = LpurchaseOrderPopupRow1SuspendCol.getText();
		boolean actRow1SuspendR2 = false;
		boolean expRow1SuspendR2 = true;

		if (actRow1SuspendColumnR2.isEmpty() || actRow1SuspendColumnR2.equalsIgnoreCase("0.00")) 
		{
			actRow1SuspendR2=true;
		}


		String actRow2SuspendColumnR2 = LpurchaseOrderPopupRow2SuspendCol.getText();
		boolean actRow2SuspendR2 = false;
		boolean expRow2SuspendR2 = true;

		if (actRow2SuspendColumnR2.isEmpty() || actRow2SuspendColumnR2.equalsIgnoreCase("0.00")) 
		{
			actRow2SuspendR2=true;
		}

		String actRow2PurOrdList = PurchaseOrder1.toString();
		String expRow2PurOrdList = excelReader.getCellData(xlSheetName, 102, 6);
		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actRow2PurOrdList);

		System.err.println(actRow2PurOrdList);
		System.err.println(expRow2PurOrdList);

		int pvoccount1 = LPurchaseOrderPopupVoucherNos.size();

		for (int i = 0; i < pvoccount1; i++) 
		{
			String data = LPurchaseOrderPopupVoucherNos.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 111, 5))) 
			{
				if (LPurchaseOrderPopupCheckBoxs.get(i).isSelected()==false) 
				{
					LPurchaseOrderPopupCheckBoxs.get(i).click();
				}
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LPurchaseOrderPopupOkBtn));
		LPurchaseOrderPopupOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_11thColumn));
		select2ndRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 112, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_10thColumn));
		String actRow2LPurOrd = select2ndRow_10thColumn.getText();
		String expRow2LPurOrd = excelReader.getCellData(xlSheetName, 105, 6);
		excelReader.setCellData(xlfile, xlSheetName, 105, 7, actRow2LPurOrd);





		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		select3rdRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(Keys.END);
		enter_WarehouseTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 113, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 114, 5));
		Thread.sleep(2000);
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
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 115, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		int POCount2 = LPurchaseOrderPopupList.size();

		ArrayList<String> PurchaseOrder2 = new ArrayList<String>();

		for (int i = 0; i < POCount2; i++) 
		{
			if (i==0 || i==6 || i==8 || i==9 /*|| i==10*/) 
			{
				String data = LPurchaseOrderPopupList.get(i).getText();
				PurchaseOrder2.add(data);
			}
		}

		String actRow1SuspendColumnR3 = LpurchaseOrderPopupRow1SuspendCol.getText();
		boolean actRow1SuspendR3 = false;
		boolean expRow1SuspendR3 = true;

		if (actRow1SuspendColumnR3.isEmpty() || actRow1SuspendColumnR3.equalsIgnoreCase("0.00")) 
		{
			actRow1SuspendR3=true;
		}


		String actRow3PurOrdList = PurchaseOrder2.toString();
		String expRow3PurOrdList = excelReader.getCellData(xlSheetName, 106, 6);
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actRow3PurOrdList);

		System.err.println(actRow3PurOrdList);
		System.err.println(expRow3PurOrdList);

		int pvoccount2 = LPurchaseOrderPopupVoucherNos.size();

		for (int i = 0; i < pvoccount2; i++) 
		{
			String data = LPurchaseOrderPopupVoucherNos.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 116, 5))) 
			{
				if (LPurchaseOrderPopupCheckBoxs.get(i).isSelected()==false) 
				{
					LPurchaseOrderPopupCheckBoxs.get(i).click();
				}
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LPurchaseOrderPopupCloseBtn));
		LPurchaseOrderPopupCloseBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_11thColumn));
		select3rdRow_11thColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys(excelReader.getCellData(xlSheetName, 117, 5));
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.click();
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_10thColumn));
		String actRow3LPurOrd = Boolean.toString(select3rdRow_10thColumn.getText().isEmpty());
		String expRow3LPurOrd = excelReader.getCellData(xlSheetName, 108, 6);
		excelReader.setCellData(xlfile, xlSheetName, 108, 7, actRow3LPurOrd);

		System.err.println(actRow1LPurOrd+"  Value Expected  "+expRow1LPurOrd);
		System.err.println(actRow2LPurOrd+"  Value Expected  "+expRow2LPurOrd);
		System.err.println(actRow3LPurOrd+"  Value Expected  "+expRow3LPurOrd);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(Keys.END);
		enter_WarehouseTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 118, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(secondRowIndex));
		//secondRowIndex.click();
		
		getAction().contextClick(secondRowIndex).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteRowBtn));
		deleteRowBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(secondRowIndex));
		//secondRowIndex.click();
		
		getAction().contextClick(secondRowIndex).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteRowBtn));
		deleteRowBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newReferenceTxt));
		newReferenceTxt.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();

		//Thread.sleep(3000);
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : SU/IND/TEXT1");
		expMessage.add("Saving in background.");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 109, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 110, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		
		String expMessage1 = excelReader.getCellData(xlSheetName, 109, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 110, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))


		if(actRow1PurOrdList.equalsIgnoreCase(expRow1PurOrdList) && actRow1LPurOrd.equalsIgnoreCase(expRow1LPurOrd)
				&& actRow1Suspend==expRow1Suspend
				&& actRow2PurOrdList.equalsIgnoreCase(expRow2PurOrdList) && actRow2LPurOrd.equalsIgnoreCase(expRow2LPurOrd)
				&& actRow1SuspendR2==expRow1SuspendR2
				&& actRow3PurOrdList.equalsIgnoreCase(expRow3PurOrdList) && actRow3LPurOrd.equalsIgnoreCase(expRow3LPurOrd)
				&& actRow1SuspendR3==expRow1SuspendR3
				&& actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))
		{
			excelReader.setCellData(xlfile, xlSheetName, 98, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 98, 8, resFail);
			return false;
		}
	}








	public static boolean checkLPurchaseOrderPopupInPurchaseVoucherVATVoucher2() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		vendorAccountTxt.sendKeys(excelReader.getCellData(xlSheetName, 119, 5));	
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderDueDate));
		voucherHeaderDueDate.click();
		voucherHeaderDueDate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		voucherHeaderCurrency.sendKeys(excelReader.getCellData(xlSheetName, 120, 5));
		Thread.sleep(2000);
		voucherHeaderCurrency.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		departmentTxt.sendKeys(excelReader.getCellData(xlSheetName, 121, 5));	
		Thread.sleep(2000);
		departmentTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		placeOFSupplyTxt.click();
		placeOFSupplyTxt.sendKeys(excelReader.getCellData(xlSheetName, 122, 5));
		Thread.sleep(2000);
		placeOFSupplyTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		jurisdictionTxt.click();
		jurisdictionTxt.sendKeys(excelReader.getCellData(xlSheetName, 123, 5));
		Thread.sleep(2000);
		jurisdictionTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		select1stRow_1stColumn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_WarehouseTxt));
		enter_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 124, 5));
		Thread.sleep(2000);
		enter_WarehouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 125, 5));
		Thread.sleep(2000);
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
		enter_Quantity.sendKeys(excelReader.getCellData(xlSheetName, 126, 5));
		enter_Quantity.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		int POCount = LPurchaseOrderPopupList.size();

		ArrayList<String> PurchaseOrder = new ArrayList<String>();

		for (int i = 0; i < POCount; i++) 
		{
			if (i==0 || i==6 || i==8 || i==9 /*|| i==10*/ || i==11 || i==17 || i==19 || i==20 /*|| i==21*/) 
			{
				String data = LPurchaseOrderPopupList.get(i).getText();
				PurchaseOrder.add(data);
			}
		}

		String actRow1SuspendColumn = LpurchaseOrderPopupRow1SuspendCol.getText();
		boolean actRow1Suspend = false;
		boolean expRow1Suspend = true;

		if (actRow1SuspendColumn.isEmpty() || actRow1SuspendColumn.equalsIgnoreCase("0.00")) 
		{
			actRow1Suspend=true;
		}


		String actRow2SuspendColumn = LpurchaseOrderPopupRow2SuspendCol.getText();
		boolean actRow2Suspend = false;
		boolean expRow2Suspend = true;

		if (actRow2SuspendColumn.isEmpty() || actRow2SuspendColumn.equalsIgnoreCase("0.00")) 
		{
			actRow2Suspend=true;
		}


		String actPurOrdList = PurchaseOrder.toString();
		String expPurOrdList = excelReader.getCellData(xlSheetName, 119, 6);
		excelReader.setCellData(xlfile, xlSheetName, 119, 7, actPurOrdList);

		System.err.println(actPurOrdList);
		System.err.println(expPurOrdList);

		int pvoccount = LPurchaseOrderPopupVoucherNos.size();

		for (int i = 0; i < pvoccount; i++) 
		{
			String data = LPurchaseOrderPopupVoucherNos.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 127, 5))) 
			{
				if (LPurchaseOrderPopupCheckBoxs.get(i).isSelected()==false) 
				{
					LPurchaseOrderPopupCheckBoxs.get(i).click();
				}
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LPurchaseOrderPopupOkBtn));
		LPurchaseOrderPopupOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();

		Thread.sleep(2000);
		
		getWaitForAlert();
		
		Thread.sleep(2000);
	
		getAlert().accept();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));
		grid_ChkBox1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHomeDeleteBtn));
		voucherHomeDeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 122, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 122, 7, actMessage);

		if(actPurOrdList.equalsIgnoreCase(expPurOrdList) && actRow1Suspend==expRow1Suspend
				&& actRow2Suspend==expRow2Suspend && actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 119, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 119, 8, resFail);
			return false;
		}
	}





	
	public static boolean checkSuspendingPurchasesOrdersVoucher2FromEntryPage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int vcount = grid_VoucherNoList.size();

		for (int i = 0; i < vcount; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 128, 5))) 
			{
				getAction().doubleClick(grid_VoucherNoList.get(i)).build().perform();
				break;
			}
		}

		checkValidationMessage("Voucher loaded successfully");
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_SuspendBtn));
		new_SuspendBtn.click();

		String expSuspendMessage1=excelReader.getCellData(xlSheetName, 128, 6);
		String expSuspendMessage2=excelReader.getCellData(xlSheetName, 129, 6);

		String actSuspendMessage=checkValidationMessage(expSuspendMessage1);
		
		excelReader.setCellData(xlfile, xlSheetName, 128, 7, actSuspendMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(2000);
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 130, 6);
		excelReader.setCellData(xlfile, xlSheetName, 130, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actSuspendMessage.startsWith(expSuspendMessage1) && actSuspendMessage.endsWith(expSuspendMessage2)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 128, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 128, 8, resFail);
			return false;
		}
	}
	
	

	
	
	
	
	public boolean checkResavingPurchasesOrdersVoucher2AndSuspendingFromHomepage() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));

		int voucherGridDocNoCount = voucherGridDocNo.size();

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 131, 5);

			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				getAction().doubleClick(voucherGridDocNo.get(i)).build().perform();
				break;
			}
		}

		Thread.sleep(2000);
		
		checkValidationMessage("Voucher loaded successfully");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		//Thread.sleep(3000);
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Saving in background.");
		expMessage.add("Voucher saved successfully : 2");*/
		
		/*String actMessage = SavingInBackground(2);

		HashSet<String> expMsg = new HashSet();

		expMsg.add(excelReader.getCellData(xlSheetName, 131, 6));
		expMsg.add(excelReader.getCellData(xlSheetName, 132, 6));
		
		String expMessage = expMsg.toString();
		
		excelReader.setCellData(xlfile, xlSheetName, 131, 7, actMessage);
		
		

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);*/
		

		String expMessage1 = excelReader.getCellData(xlSheetName, 109, 6);
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = excelReader.getCellData(xlSheetName, 110, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 109, 7, actMessage);
	
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		//if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2))

		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		
		Thread.sleep(2000);

		for(int i=0;i<voucherGridDocNoCount;i++)
		{
			String data =excelReader.getCellData(xlSheetName, 132, 5);

			if(voucherGridDocNo.get(i).getText().equalsIgnoreCase(data))
			{
				if (voucherGridIndexChkBox.get(i).isSelected()==false) 
				{
					voucherGridIndexChkBox.get(i).click();
					break;
				}
			}
		}
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(suspendBtn));
		suspendBtn.click();
		
		String expSuspendMessage = excelReader.getCellData(xlSheetName, 133, 6);
		String actSuspendMessage = checkValidationMessage(expSuspendMessage);
		excelReader.setCellData(xlfile, xlSheetName, 133, 7, actSuspendMessage);
		
		if (validationConfirmationMessage.getText().isEmpty()==false) 
		{
			checkValidationMessage("");
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHomeRow1SuspendedStatus));
		
		String actSuspendStatus = voucherHomeRow1SuspendedStatus.getText();
		String expSuspendStatus = excelReader.getCellData(xlSheetName, 134, 6);
		excelReader.setCellData(xlfile, xlSheetName, 134, 7, actSuspendStatus);
		
		System.out.println("Suspended Status : "+actSuspendStatus+"  Value Expected  "+expSuspendStatus);

		if(actMessage.startsWith(expMessage1)&& actMessage.endsWith(expMessage2) && actSuspendMessage.equalsIgnoreCase(expSuspendMessage)
				&& actSuspendStatus.equalsIgnoreCase(expSuspendStatus))
		{
			excelReader.setCellData(xlfile, xlSheetName, 131, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 131, 8, resFail);
			return false;
		}
	}
	
	
	
	


	public boolean checkDeletingVoucher2PurchaseOrder() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsMenu));
		inventoryTransactionsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryTransactionsPurchasesMenu));
		inventoryTransactionsPurchasesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesOrdersVoucher));
		purchasesOrdersVoucher.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int vcount = grid_VoucherNoList.size();

		for (int i = 0; i < vcount; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 135, 5))) 
			{
				if (grid_CheckBoxList.get(i).isSelected()==false) 
				{
					grid_CheckBoxList.get(i).click();
					break;
				}
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHomeDeleteBtn));
		voucherHomeDeleteBtn.click();

		getWaitForAlert();
		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 135, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 135, 7, actMessage);
		
		if (validationConfirmationMessage.getText().isEmpty()==false) 
		{
			checkValidationMessage("");
		}

		if (actMessage.equalsIgnoreCase(expMessage))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 135, 8, resPass);
			return true;
		}
		else                                                                    
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(poVoucherHomeCloseBtn));
			poVoucherHomeCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 135, 8, resFail);
			return false;
		}
	}

	
	
	
	
	
	@FindBy(xpath = "//*[@id='errmsgDiv']")
    public static WebElement loginPageErrorMesg;
    
    public static void reLogin(String unamelt, String pawslt,String compname)
    {
    	LoginPage lp=new LoginPage(getDriver());
            try {
                    
                    if(loginPageErrorMesg.isDisplayed())
                    {
                            
                            Thread.sleep(1999);
                            username.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
                            
                            lp.enterUserName(unamelt);

                            Thread.sleep(2000);
                            
                            password.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);

                            lp.enterPassword(pawslt);
                            
                            Select oSelect = new Select(companyDropDownList);
                            oSelect.selectByValue("36");

                           /* List <WebElement> elementCount = oSelect.getOptions();

                            int cqSize = elementCount.size();

                            System.out.println("CompanyDropdownList Count :"+cqSize);

                            int i;

                            for(i=0; i<elementCount.size(); i++) 
                            {

                                    elementCount.get(i).getText();

                                    String optionName = elementCount.get(i).getText();
                                    if(optionName.toUpperCase().startsWith(compname.toUpperCase()))
                                    {
                                            System.out.println("q"+elementCount.get(i).getText());
                                            elementCount.get(i).click();
                                    }

                            }*/

                           lp.clickOnSignInBtn();
                            
                            
                    }
                    
            } catch (Exception e) {
                    // TODO: handle exception
            }
            
            
    }



	public PurchaseOrdersPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
