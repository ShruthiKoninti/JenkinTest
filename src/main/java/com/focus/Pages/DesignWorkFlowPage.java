package com.focus.Pages;

import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class DesignWorkFlowPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "DesignWorkFlowPage";
	
	public static String checkValidationMessage(String ExpMessage) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		try
		{
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
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





	



	public boolean checkSavingTestWorkFlow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		ClickUsingJs(settingsmenuBtn);
		Thread.sleep(2000);
		ClickUsingJs(DesignWorkFlowBut);
		
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workflowNameTxt));
		workflowNameTxt.click();
		Thread.sleep(1500);
		workflowNameTxt.sendKeys("TestFlow");
		Thread.sleep(1500);

		workflowNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(requestForQuoteDragAndDrop));
		requestForQuoteDragAndDrop.click();
		Thread.sleep(2000);

		// getAction().dragAndDropBy(purchasesOrdersDragAndDrop, 350,
		// -250).build().perform();
		getAction().dragAndDropBy(requestForQuoteDragAndDrop, 350, -100).build().perform();

		//click(purchasesVouchersExpandBtn);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesQuotationsDragAndDrop));
		purchasesQuotationsDragAndDrop.click();

		getAction().dragAndDropBy(purchasesQuotationsDragAndDrop, 650, 0).build().perform();

		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FirstDraggedElement));

		/*boolean actFirstDraggedElement = FirstDraggedElement.isDisplayed();
		boolean actdraggedPurchasesVouchersN = draggedPurchasesVouchersN1.isDisplayed();

		System.err.println(actFirstDraggedElement);
		System.err.println(actdraggedPurchasesVouchersN);

		boolean expFirstDraggedElement = true;
		boolean expdraggedPurchasesVouchersN = true;*/

		getAction().contextClick(FirstDraggedElement).build().perform();

		Thread.sleep(10000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(draggedlinkBtn));

		boolean actdraggedlinkBtn = draggedlinkBtn.isDisplayed();
		boolean actdraggeddeleteBtn = draggeddeleteBtn.isDisplayed();

		System.err.println("actdraggedlinkBtn is :" + actdraggedlinkBtn);
		System.err.println("actdraggeddeleteBtn is :" + actdraggeddeleteBtn);

		boolean expdraggedlinkBtn = true;
		boolean expdraggeddeleteBtn = true;

		click(draggedlinkBtn);
		System.err.println("hello");

		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(linkNode0Right));
		linkNode0Right.click();
		
		
		Thread.sleep(2000);
		
		linkNode1Left.click();

		Thread.sleep(2000);

		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("document.querySelector('#line0 > path').setAttribute('stroke-width', '21')");

		boolean actlink0TO1 = link0TO1.isDisplayed();
		boolean explink0TO1 = true;

		Thread.sleep(2000);
		System.err.println("actlink0TO1 is :" + actlink0TO1);

		// getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(linkNode1Left));
		// linkNode1Left.click();

		//Thread.sleep(2000);
		//moveToElement(link5TO6);

		/*
		 * Thread.sleep(2000); moveToElement(link0TO1); ScrollIntoView(link0TO1);
		 * Thread.sleep(2000); click(link0TO1);
		 */
		 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(link0TO1));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(link0TO1));
		link0TO1.click();
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DefinitionTab));
		DefinitionTab.click();

		Thread.sleep(2000);

		Select s = new Select(DefinitionLinkValueDropdown);

		s.selectByVisibleText("Quantity");

		int definitionTabChkBoxNameListCount = definitionTabChkBoxNameList.size();

		for (int i = 1; i < definitionTabChkBoxNameListCount; i++) {
			String data = definitionTabChkBoxNameList.get(i).getText();

			if (data.equalsIgnoreCase("ITEM")) {
				definitionTabChkBoxList.get(i).click();
			}
		}

		// getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBtn));
		// OkBtn.click();
		ClickUsingJs(OkBtn);*/
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DefinitionLinkValueDropdown));
		Select s=new Select(DefinitionLinkValueDropdown);
		s.selectByVisibleText(/* excelReader.getCellData(xlSheetName, 10, 5) */"Quantity");
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(materialRequiitionAndPquotationsitemchkbox));
		materialRequiitionAndPquotationsitemchkbox.click();
		Thread.sleep(2000);
		
		//getAction().moveToElement(OkBtn).build().perform();
		
		//((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", OkBtn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBtn));
		ClickUsingJs(OkBtn);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwSaveBtn));
		dwSaveBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwCloseBtn));
		dwCloseBtn.click();

		System.out.println("********************************checkSavingTestWorkFlow***********************************");

		String expMessage=excelReader.getCellData(xlSheetName, 8, 6);

		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
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



	
	
	@FindBy(xpath="//*[@id='idNode_0']")
	public static WebElement base;


	public boolean checkEditingSavedTestWorkFlow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		ClickUsingJs(settingsmenuBtn);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DesignWorkFlowBut));
		DesignWorkFlowBut.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workflowNameTxt));
		workflowNameTxt.click();
		Thread.sleep(1000);
		clearValueFromTextBox(workflowNameTxt);
		Thread.sleep(1000);
		workflowNameTxt.sendKeys(/*excelReader.getCellData(xlSheetName, 11, 5)*/"TestFlow");
		Thread.sleep(1000);
		workflowNameTxt.sendKeys(Keys.TAB);
		Thread.sleep(1200);
		base.click();
		
		//ClickUsingJs(Link0T01Click);
		/*ScrollIntoView(link0TO1);
		
		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("document.querySelector('#line0 > path').setAttribute('stroke-width', '8')");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(link0TO1));
		getAction().moveToElement(link0TO1).pause(2000).click().build().perform();*/

		Thread.sleep(2000);

		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("document.querySelector('#line0 > path').setAttribute('stroke-width', '25')");
		Thread.sleep(1500);
		
		moveToElement(link0TO1);

		 //getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(link0TO1));
		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(link0TO1));
		//link0TO1.click();
		Thread.sleep(2000);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DefinitionLinkValueDropdown));
		Select s=new Select(DefinitionLinkValueDropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 12, 5));
		
		Thread.sleep(1200);
		
		System.err.println("customer account chkbox is :" + CheckcustomerAccountChkbox.isSelected());

		if(CheckcustomerAccountChkbox.isSelected()==false)
		{
		customerAccountChkbox.click();
		}
		Thread.sleep(2000);
		
		ClickUsingJs(OkBtn);
		
		//((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", OkBtn);

		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBtn));
		//OkBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwSaveBtn));
		dwSaveBtn.click();

		System.out.println("*******************************checkEditingSavedTestWorkFlow**************************");

		String expMessage=excelReader.getCellData(xlSheetName, 11, 6);

		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resFail);
			return false;
		}
	}






	public boolean checkDeletingTestWorkFlow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workflowNameTxt));
		workflowNameTxt.click();
		workflowNameTxt.sendKeys(Keys.END);
		workflowNameTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		workflowNameTxt.sendKeys(/*excelReader.getCellData(xlSheetName, 13, 5)*/"TestFlow");
		Thread.sleep(2000);
		workflowNameTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwDeleteBtn));
		dwDeleteBtn.click();

		Thread.sleep(1000);

		getWaitForAlert();
		getAlert().accept();

		System.out.println("****************************checkDeletingTestWorkFlow***************************");

		String expMessage=excelReader.getCellData(xlSheetName, 13, 6);

		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 13, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 13, 8, resFail);
			return false;
		}
	}







	public boolean checkEditingPurchasesFlow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		ClickUsingJs(settingsmenuBtn);

		int count = SettingsSubMenusList.size();

		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();

			System.err.println(data);

			ScrollIntoView(SettingsSubMenusList.get(i));

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 14, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workflowNameTxt));
		workflowNameTxt.click();
		workflowNameTxt.sendKeys(Keys.END);
		workflowNameTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		workflowNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
		Thread.sleep(2000);
		workflowNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersExpandBtn));
		purchaseVouchersExpandBtn.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesVouchersVATDragAndDrop));
		purchasesVouchersVATDragAndDrop.click();

		//getAction().dragAndDrop(purchasesVouchersVATDragAndDrop, draggedMaterialRequisition).build().perform();

		getAction().dragAndDropBy(purchasesVouchersVATDragAndDrop, 800, -150).build().perform();
		
		Thread.sleep(2000);

		getAction().contextClick(draggedPurchasesOrders).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(draggedlinkBtn));
		draggedlinkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(linkNode5Left));
		linkNode5Left.click();

		Thread.sleep(2000);

		getAction().contextClick(draggedPurchasesVouchersVAT).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(draggedlinkBtn));
		draggedlinkBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(linkNode2Left));
		linkNode2Left.click();
		
		Thread.sleep(2000);
		
		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("document.querySelector('#line4 > path').setAttribute('stroke-width', '25')");

        Thread.sleep(1500);
		
		moveToElement(link4TO5);
		
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DefinitionLinkValueDropdown));
		Select s=new Select(DefinitionLinkValueDropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 16, 5));

		Thread.sleep(2000);
		
		int definitionTabChkBoxNameListCount = definitionTabChkBoxNameList.size();

		for(int i=0;i<definitionTabChkBoxNameListCount;i++)
		{
			String data = definitionTabChkBoxNameList.get(i).getText();
			System.err.println("***"+data+"***");
			
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 17, 5)))
			{
				definitionTabChkBoxList.get(i).click();
				System.err.println("**Item is Selected**");
				break;
			}
		}

		boolean selected = false;

		for(int i=0;i<definitionTabChkBoxNameListCount;i++)
		{
			String data = definitionTabChkBoxNameList.get(i).getText();
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 18, 5)))
			{
				selected=definitionTabChkBoxList.get(i).isSelected();

				break;

			}
		}

		Thread.sleep(2000);

		
		ClickUsingJs(OkBtn);

		Thread.sleep(3000);
		
		JavascriptExecutor js1 = (JavascriptExecutor) getDriver();
		js1.executeScript("document.querySelector('#line5 > path').setAttribute('stroke-width', '25')");
		
        Thread.sleep(1500);
		
		moveToElement(link5TO6);
		

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DefinitionLinkValueDropdown));
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 19, 5));


		int definitionTabChkBoxNameListCount1 = definitionTabChkBoxNameList.size();

		for(int i=1;i<definitionTabChkBoxNameListCount1;i++)
		{
			String data = definitionTabChkBoxNameList.get(i).getText();
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 20, 5)))
			{
				definitionTabChkBoxList.get(i).click();
				break;
			}
		}

		boolean selected1 = false;

		for(int i=1;i<definitionTabChkBoxNameListCount1;i++)
		{
			String data = definitionTabChkBoxNameList.get(i).getText();
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 21, 5)))
			{
				selected1=definitionTabChkBoxList.get(i).isSelected();

				break;

			}
		}

		Thread.sleep(2000);
		
		ClickUsingJs(OkBtn);

		Thread.sleep(2000);
		
		ClickUsingJs(dwSaveBtn);

		System.out.println("************************checkSavingPurchasesFlow*************************************");

		String expMessage=excelReader.getCellData(xlSheetName, 14, 6);

		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actMessage);


		if(actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Test Pass : Error Message is as Expected");
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwCloseBtn));
			dwCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 14, 8, resPass);
			return true;
		}else
		{
			System.out.println("Test Fail : Error Message is NOT as Expected");
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwCloseBtn));
			dwCloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 14, 8, resFail);
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

	       /* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	        companyLogo.click();*/

	        String getCompanyTxt1=Company_Name.getText();
	        String getLoginCompanyName1=getCompanyTxt1.substring(0, 31);
	        System.out.println("company name  :  "+ getLoginCompanyName1);
	        //companyLogo.click();

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


	public boolean checkLogoutDesignWorkFlowPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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


















	public DesignWorkFlowPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}


}
