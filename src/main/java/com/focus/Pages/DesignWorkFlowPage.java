package com.focus.Pages;

import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();

		int count = SettingsSubMenusList.size();

		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();

			System.err.println(data);

			ScrollIntoView(SettingsSubMenusList.get(i));

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workflowNameTxt));
		workflowNameTxt.click();
		workflowNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		Thread.sleep(1000);
		workflowNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(requestForQuoteDragAndDrop));
		requestForQuoteDragAndDrop.click();

		getAction().dragAndDropBy(requestForQuoteDragAndDrop, 350, -250).build().perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesQuotationsDragAndDrop));
		purchasesQuotationsDragAndDrop.click();

		getAction().dragAndDropBy(purchasesQuotationsDragAndDrop, 650, 28).build().perform();
		
		//getAction().dragAndDropBy(SecondDraggedElement, 350, 0).build().perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(draggedPurchasesQuotations));

		getAction().contextClick(draggedPurchasesQuotations).build().perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(draggedlinkBtn));
		draggedlinkBtn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(linkNode1Left));
		linkNode1Left.click();

	
		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("document.querySelector('#line0 > path').setAttribute('stroke-width', '8')"); 

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(link0TO1));
		getAction().moveToElement(link0TO1).pause(2000).click().build().perform();
		//link0TO1.click();
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DefinitionLinkValueDropdown));
		Select s=new Select(DefinitionLinkValueDropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 10, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(materialRequiitionAndPquotationsitemchkbox));
		materialRequiitionAndPquotationsitemchkbox.click();
		Thread.sleep(2000);
		
		//getAction().moveToElement(OkBtn).build().perform();
		
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", OkBtn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBtn));
		OkBtn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwSaveBtn));
		dwSaveBtn.click();

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






	public boolean checkEditingSavedTestWorkFlow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workflowNameTxt));
		workflowNameTxt.click();
		workflowNameTxt.sendKeys(Keys.END);
		workflowNameTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		workflowNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		workflowNameTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		
		
		//ScrollIntoView(link0TO1);
		
		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("document.querySelector('#line0 > path').setAttribute('stroke-width', '8')");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(link0TO1));
		getAction().moveToElement(link0TO1).pause(2000).click().build().perform();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DefinitionLinkValueDropdown));
		Select s=new Select(DefinitionLinkValueDropdown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 12, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountChkbox));
		customerAccountChkbox.click();
		
		Thread.sleep(2000);
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", OkBtn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBtn));
		OkBtn.click();
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
		workflowNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
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
		settingsmenuBtn.click();

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

		getAction().dragAndDropBy(purchasesVouchersVATDragAndDrop, 800, -250).build().perform();

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
		js.executeScript("document.querySelector('#line4 > path').setAttribute('stroke-width', '8')");

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(link4TO5));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(link4TO5));
		link4TO5.click();
		
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

		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", OkBtn);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBtn));
		OkBtn.click();

		Thread.sleep(3000);
		
		JavascriptExecutor js1 = (JavascriptExecutor) getDriver();
		js1.executeScript("document.querySelector('#line5 > path').setAttribute('stroke-width', '8')");
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(link5TO6));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(link5TO6));
		link5TO6.click();

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
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBtn));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBtn));
		OkBtn.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwSaveBtn));
		dwSaveBtn.click();

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
















	public DesignWorkFlowPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
