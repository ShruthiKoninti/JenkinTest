package com.focus.Pages;
import com.focus.base.BaseEngine;

import java.awt.AWTException;
import java.io.File;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;


public class SmokeDashboardPage extends BaseEngine
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "SmokeDashboardPage";

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




	public static boolean checkDashboard() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
		String getDashboard=dashboard.getText();

		System.out.println(getDashboard);

		String actualnewDashBoard = Boolean.toString(newDashBoard.isDisplayed());
		String actualeditDashBoard = Boolean.toString(editDashBoard.isDisplayed());
		String actualdashBoardCustomizeBtn = Boolean.toString(dashBoardCustomizeBtn.isDisplayed());


		String expnewDashBoard = excelReader.getCellData(xlSheetName, 8, 6);
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actualnewDashBoard);

		String expeditDashBoard = excelReader.getCellData(xlSheetName, 9, 6);
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actualeditDashBoard);

		String expdashBoardCustomizeBtn = excelReader.getCellData(xlSheetName, 10, 6);
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actualdashBoardCustomizeBtn);

		if(actualnewDashBoard.equalsIgnoreCase(expnewDashBoard) && actualeditDashBoard.equalsIgnoreCase(expeditDashBoard) && actualdashBoardCustomizeBtn.equalsIgnoreCase(expdashBoardCustomizeBtn))
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



	public static boolean checkSaveBtnWithInputDashboardName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
		dashboard.click();

		Thread.sleep(1000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(newDashBoard));
		newDashBoard.click();

		Thread.sleep(1000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardName));
		dashBoardName.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		Thread.sleep(2000);
		dashBoardName.sendKeys(Keys.TAB);

		checkValidationMessage("");

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardSaveBtn));
		dashBoardSaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 11, 6);

		String actMessage = checkValidationMessage(expMessage);

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





	public static boolean checkSavedDashboardFromSelectDashboard() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getDriver().navigate().refresh();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardFirst));

		String actSavedDashboard = dashBoardFirst.getText();
		String expSavedDashboard = excelReader.getCellData(xlSheetName, 12, 6);
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actSavedDashboard);

		if(actSavedDashboard.equalsIgnoreCase(expSavedDashboard))
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








	public static boolean checkUpdateDashboardByClickOnSaveBtnInDashboard() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardFirst));
		dashBoardFirst.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(editDashBoard));
		editDashBoard.click();
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardName));
		dashBoardName.click();
		Thread.sleep(2000);
		dashBoardName.sendKeys(Keys.END);
		Thread.sleep(2000);
		dashBoardName.sendKeys(Keys.SHIFT, Keys.HOME);
		Thread.sleep(2000);
		dashBoardName.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		Thread.sleep(2000);
		dashBoardName.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardSetAsDefaultCheckbox));
		dashBoardSetAsDefaultCheckbox.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(userAllOptionsBtn));
		userAllOptionsBtn.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardRightSideArrowIcon));
		dashBoardRightSideArrowIcon.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardSaveBtn));
		dashBoardSaveBtn.click();
		Thread.sleep(2000);

		String expMessage = excelReader.getCellData(xlSheetName, 13, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, expMessage);

		

		/*getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCancelBtn));
		dashBoardCancelBtn.click();*/

		getDriver().navigate().refresh();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardFirst));
		String actDashboard=dashBoardFirst.getText();
		String expDashboard=excelReader.getCellData(xlSheetName, 14, 6);
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actDashboard);

		System.out.println("*****Dashboard******  "+actDashboard+"  Value Expected  "+expDashboard);

		if(actMessage.equalsIgnoreCase(expMessage) && actDashboard.equalsIgnoreCase(expDashboard))
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


	@FindBy(xpath="//*[@id='ddlCompany']")
	private static WebElement companyList;
	
   

	public static boolean checkLogoutandLoginWithNewUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		/*getWebDriverWait().until(ExpectedConditions.elementToBeClickable(signoutBtn));
					 signoutBtn.click();*/

		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		Thread.sleep(2000);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();

		Thread.sleep(2000);
		
	/*	LoginPage.reLogin(excelReader.getCellData(xlSheetName, 15, 5),excelReader.getCellData(xlSheetName, 16, 5),"Automation Company");
		Thread.sleep(5000);*/
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 15, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.click();
		password.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		Thread.sleep(2000);
		
		 Select s1=new Select(companyList);
	        s1.selectByValue("36");
	        Thread.sleep(2000);
	  

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("12345");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();

		Thread.sleep(2000);
		

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardFirst));

		Thread.sleep(2000);

		dashBoardFirst.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(editDashBoard));
		editDashBoard.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardSetAsDefaultCheckbox));

		if (dashBoardSetAsDefaultCheckbox.isSelected()==false) 
		{
			dashBoardSetAsDefaultCheckbox.click();
		}

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardSaveBtn));
		dashBoardSaveBtn.click();
		Thread.sleep(2000);

		String expMessage = excelReader.getCellData(xlSheetName, 15, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actMessage);

	

		/*getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCancelBtn));
		dashBoardCancelBtn.click();*/

		getDriver().navigate().refresh();

		Thread.sleep(3000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardFirst));
		String actDashboard=dashBoardFirst.getText();
		String expDashboard=excelReader.getCellData(xlSheetName, 16, 6);
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actDashboard);

		System.out.println("Dashboard Name : "+actDashboard+"  Value Expected  "+expDashboard);

		if(actDashboard.equalsIgnoreCase(expDashboard))	
		{
			excelReader.setCellData(xlfile, xlSheetName, 15, 8, resPass);
			return true;
		} 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 15, 8, resFail);
			return false;
		}
	}





	public static boolean checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp=new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 17, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();

		Thread.sleep(5000);
		
		lp.reLogin(excelReader.getCellData(xlSheetName, 17, 5),excelReader.getCellData(xlSheetName, 18, 5),"Automation Company");
		Thread.sleep(5000);

		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 17, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		Thread.sleep(2000);
		
		 Select s1=new Select(companyList);
	        s1.selectByValue("36");
	        Thread.sleep(2000);
	  

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		
		
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		userNameDisplay.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 17, 6);
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actUserInfo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 18, 6);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actLoginCompanyName);
		companyLogo.click();

		System.out.println("UserInfo               : ."+actUserInfo          +".  value expected  ."+expUserInfo+".");
		System.out.println("LoginCompanyName       : ."+actLoginCompanyName  +".  value expected  ."+expLoginCompanyName+".");

		if(actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName))  
		{
			excelReader.setCellData(xlfile, xlSheetName, 17, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 17, 8, resFail);
			return false;
		}
	}







	public static boolean checkCustomizationWithSelectingDashboard() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardFirst));
		dashBoardFirst.click();

		Thread.sleep(1000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeBtn));
		dashBoardCustomizeBtn.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeSearch));

		Thread.sleep(2000);


		String actualdashBoardCustomizeSearch        = Boolean.toString(dashBoardCustomizeSearch.isDisplayed());
		String actualdashBoardCustomizeGraph         = Boolean.toString(dashBoardCustomizeGraph.isDisplayed());
		//String actualdashBoardCustomizeReport        = Boolean.toString(dashBoardCustomizeSearch.isDisplayed());
		String actualdashBoardCustomizeInfoPanel     = Boolean.toString(dashBoardCustomizeInfoPanel.isDisplayed());
		String actualdashBoardCustomizeInfoWorkFlow  = Boolean.toString(dashBoardCustomizeInfoWorkFlow.isDisplayed());


		String expdashBoardCustomizeSearch        = excelReader.getCellData(xlSheetName, 19, 6);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actualdashBoardCustomizeSearch);

		String expdashBoardCustomizeGraph         = excelReader.getCellData(xlSheetName, 20, 6);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actualdashBoardCustomizeGraph);

	/*	String expdashBoardCustomizeReport        = excelReader.getCellData(xlSheetName, 21, 6);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actualdashBoardCustomizeReport);
*/
		String expdashBoardCustomizeInfoPanel     = excelReader.getCellData(xlSheetName, 22, 6);
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actualdashBoardCustomizeInfoPanel);

		String expdashBoardCustomizeInfoWorkFlow  = excelReader.getCellData(xlSheetName, 23, 6);
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actualdashBoardCustomizeInfoWorkFlow);




		if(actualdashBoardCustomizeSearch.equalsIgnoreCase(expdashBoardCustomizeSearch) && actualdashBoardCustomizeGraph.equalsIgnoreCase(expdashBoardCustomizeGraph) && 
				/*actualdashBoardCustomizeReport.equalsIgnoreCase(expdashBoardCustomizeReport) &&*/ actualdashBoardCustomizeInfoPanel.equalsIgnoreCase(expdashBoardCustomizeInfoPanel) &&
				actualdashBoardCustomizeInfoWorkFlow.equalsIgnoreCase(expdashBoardCustomizeInfoWorkFlow))
		{
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 19, 8, resFail);
			return false;
		}
	} 






	public static boolean checkGraphOptionsInCustomizeDashboard() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath()); 

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraph));
		dashBoardCustomizeGraph.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphNewGraph));


		String actualdashBoardCustomizeGraphNewGraph   =Boolean.toString(dashBoardCustomizeGraphNewGraph.isDisplayed());
		String expdashBoardCustomizeGraphNewGraph      =excelReader.getCellData(xlSheetName, 24, 6);
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actualdashBoardCustomizeGraphNewGraph);

		if(actualdashBoardCustomizeGraphNewGraph.equalsIgnoreCase(expdashBoardCustomizeGraphNewGraph))
		{
			excelReader.setCellData(xlfile, xlSheetName, 24, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 24, 8, resFail);
			return false;
		}
	}





	public static boolean checkSaveInTheDefineGraphInDefineGraph() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		excelReader=new ExcelReader(POJOUtility.getExcelPath()); 

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphNewGraph));
		dashBoardCustomizeGraphNewGraph.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefTab));
		dashBoardCustomizeGraphDefTab.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefName));
		dashBoardCustomizeGraphDefName.sendKeys(excelReader.getCellData(xlSheetName, 25, 5));

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefSeriesName));
		dashBoardCustomizeGraphDefSeriesName.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));

		Thread.sleep(3000);

		dashBoardCustomizeGraphDefSeriesName.sendKeys(Keys.TAB);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefSave));
		dashBoardCustomizeGraphDefSave.click();

		String expMessage = excelReader.getCellData(xlSheetName, 25, 6);
		String actMessage = checkValidationMessage(expMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 25, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 25, 8, resFail);
			return false;
		}
	}







	public static boolean checkDragAndDropGraphAfterDeletingInCreatedashlet() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());  

		Actions builder = new Actions(getDriver());

		Action dragAndDrop = builder.clickAndHold(dashBoardGraphNewFromDrag).build();

		dragAndDrop.perform();

		builder.dragAndDropBy(dashBoardGraphNewToDrag, -400,200).build().perform();

		Thread.sleep(2000);
		dashBoardGraphNewToDrag.click();
		
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeDashSave));
		dashBoardCustomizeDashSave.click();	
		Thread.sleep(2000);

		String expMessage = excelReader.getCellData(xlSheetName, 27, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actMessage);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdDashletInHomepage));
		String actDashletCreated = createdDashletInHomepage.getText();
		String expDashletCreated = excelReader.getCellData(xlSheetName, 28, 6);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actDashletCreated);

		System.out.println("***Dashlet Created*** "+actDashletCreated+"  Value Expected  "+expDashletCreated);

		if(actMessage.equalsIgnoreCase(expMessage) && actDashletCreated.equalsIgnoreCase(expDashletCreated))
		{ 
			excelReader.setCellData(xlfile, xlSheetName, 27, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 27, 8, resFail);
			return false;
		} 
	}






	public static boolean checkUpdateGraphCreatedInDashletsInDefineGraph() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeBtn));
		dashBoardCustomizeBtn.click();

		Thread.sleep(5000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeNewGraphCreated));
		dashBoardCustomizeNewGraphCreated.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeEditGraphCreated));
		dashBoardCustomizeEditGraphCreated.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefName));
		dashBoardCustomizeGraphDefName.click();
		dashBoardCustomizeGraphDefName.clear();
		dashBoardCustomizeGraphDefName.sendKeys(excelReader.getCellData(xlSheetName, 29, 5));

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefSave));
		dashBoardCustomizeGraphDefSave.click();

		String expMessage = excelReader.getCellData(xlSheetName, 29, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actMessage);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeDashSave));
		dashBoardCustomizeDashSave.click();	

		String expMessage1 = excelReader.getCellData(xlSheetName, 30, 6);
		String actMessage1 = checkValidationMessage(expMessage1);
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actMessage1);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdDashletInHomepage));
		String actDashletCreated = createdDashletInHomepage.getText();
		String expDashletCreated = excelReader.getCellData(xlSheetName, 31, 6);
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actDashletCreated);

		System.out.println("***Dashlet Created*** "+actDashletCreated+"  Value Expected  "+expDashletCreated);

		getDriver().navigate().refresh();

		Thread.sleep(2000);

		if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) && actDashletCreated.equalsIgnoreCase(expDashletCreated))
		{
			excelReader.setCellData(xlfile, xlSheetName, 29, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 29, 8, resFail);
			return false;
		}
	}




	public static boolean checkDashletGraphCreatedInDashboardAndDeleteDashlet() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{

		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeBtn));

		dashBoardCustomizeBtn.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashlet));

		getAction().contextClick(dashlet).perform();
		dashboardInfoPanelDeleteOption.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboardDeleteOptionMessageOkBtn));
		dashboardDeleteOptionMessageOkBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 32, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
		{

			excelReader.setCellData(xlfile, xlSheetName, 32, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 32, 8, resFail);
			return false;
		}
	}









	public static boolean checkYesBtnInDeleteConfirmationMessagGraphCreatedInDashletsInDefineGraph() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getDriver().navigate().refresh();
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraph));
		dashBoardCustomizeGraph.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeNewGraphCreated));

		dashBoardCustomizeNewGraphCreated.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeDeleteGraphCreated));

		dashBoardCustomizeDeleteGraphCreated.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeDeleteDasletYesBtn));

		dashBoardCustomizeDeleteDasletYesBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 33, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actMessage);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeDashCancel));
		dashBoardCustomizeDashCancel.click();

		getDriver().navigate().refresh();

		Thread.sleep(2000);

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 33, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 33, 8, resFail);
			return false;
		}
	}






	public static boolean checkAddingGraphDashlet() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath()); 

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeBtn));
		dashBoardCustomizeBtn.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraph));
		dashBoardCustomizeGraph.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphNewGraph));

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphNewGraph));
		dashBoardCustomizeGraphNewGraph.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefTab));
		dashBoardCustomizeGraphDefTab.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefName));
		dashBoardCustomizeGraphDefName.sendKeys(excelReader.getCellData(xlSheetName, 34, 5));

		Select showas = new Select(dashBoardCustomizeGraphDefShowAs);
		showas.selectByVisibleText(excelReader.getCellData(xlSheetName, 35, 5));

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefSeriesName));
		dashBoardCustomizeGraphDefSeriesName.sendKeys(excelReader.getCellData(xlSheetName, 36, 5));

		Thread.sleep(3000);

		dashBoardCustomizeGraphDefSeriesName.sendKeys(Keys.TAB);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefSave));
		dashBoardCustomizeGraphDefSave.click();

		String expMessage = excelReader.getCellData(xlSheetName, 34, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actMessage);


		Actions builder = new Actions(getDriver());

		Action dragAndDrop = builder.clickAndHold(dashBoardGraphNewFromDrag).build();

		dragAndDrop.perform();

		builder.dragAndDropBy(dashBoardGraphNewToDrag, -400,-185).build().perform();

		Thread.sleep(3000);


		getAction().dragAndDropBy(dashlateDrag, 250, 200).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeReport));
		dashBoardCustomizeReport.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeReportNewReport));
		dashBoardCustomizeReportNewReport.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeReportReportName));
		Select reportDrpdwn = new Select(dashBoardCustomizeReportReportName);
		reportDrpdwn.selectByValue(excelReader.getCellData(xlSheetName, 37, 5));

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeReportDateOptionDrpdwn));
		Select dateOption = new Select(dashBoardCustomizeReportDateOptionDrpdwn);
		dateOption.selectByVisibleText(excelReader.getCellData(xlSheetName, 38, 5));

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboardCustomizeReportSelectAllItemsChkbox));
		dashboardCustomizeReportSelectAllItemsChkbox.click();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeGraphDefSave));
		dashBoardCustomizeGraphDefSave.click();

		String actMessage1 = checkValidationMessage(expMessage);

		Thread.sleep(2000);

		Actions builder1 = new Actions(getDriver());

		Action dragAndDropReport = builder1.clickAndHold(newReportNameDisplayInCreateDashlet).build();

		builder1.perform();

		builder1.dragAndDropBy(newReportNameDisplayInCreateDashlet, -200, -200).build().perform();

		getAction().sendKeys(Keys.HOME).build().perform();

		Thread.sleep(1000);

		builder1.dragAndDropBy(createdDashlet1, -200, -40).build().perform();

		getAction().dragAndDropBy(dashlate1Drag, 250, 200).build().perform();

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeDashSave));
		dashBoardCustomizeDashSave.click();	

		String expMessage2 = excelReader.getCellData(xlSheetName, 35, 6);
		String actMessage2 = checkValidationMessage(expMessage2);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actMessage2);

		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdDashletInHomepage));
		String actDashletCreated = createdDashletInHomepage.getText();
		String expDashletCreated = excelReader.getCellData(xlSheetName, 36, 6);
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actDashletCreated);

		System.out.println("***Dashlet Created*** "+actDashletCreated+"  Value Expected  "+expDashletCreated);

		if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage) && actMessage2.equalsIgnoreCase(expMessage2) 
				&& actDashletCreated.equalsIgnoreCase(expDashletCreated))
		{
			excelReader.setCellData(xlfile, xlSheetName, 34, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 34, 8, resFail);
			return false;
		}
	}







	public static boolean checkDashBoardDeleted() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	{

		excelReader=new ExcelReader(POJOUtility.getExcelPath()); 

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardFirst));

		dashBoardFirst.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteDashBoard));

		deleteDashBoard.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardConfirmationMessage));
		String actDashboardMsg=dashBoardConfirmationMessage.getText();
		String getDashboardMsg=excelReader.getCellData(xlSheetName, 39, 6);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actDashboardMsg);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardConfirmationMessageOkBtn));
		dashBoardConfirmationMessageOkBtn.click();

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(validationDashboardMessageAtBotton));

		String getDeleteDashboardMessage=validationDashboardMessageAtBotton.getText();
		String expDeleteDashboardMessage=excelReader.getCellData(xlSheetName, 40, 6);
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, getDeleteDashboardMessage);

		System.out.println("DashboardMsg             : "+actDashboardMsg+"  Value Expected  "+getDashboardMsg);
		System.out.println("DeleteDashboardMessage   : "+getDeleteDashboardMessage+"  Value Expected  "+expDeleteDashboardMessage);

		if(actDashboardMsg.equalsIgnoreCase(getDashboardMsg) && getDeleteDashboardMessage.equalsIgnoreCase(expDeleteDashboardMessage))
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



	
	/////
	
	public boolean checkOtherDashletOptionsinDashboard() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeBtn));

		dashBoardCustomizeBtn.click();

		Thread.sleep(2000);
		
		
		Actions builder = new Actions(getDriver());

		Action dragAndDrop = builder.clickAndHold(dashBoardGraphNewFromDrag).build();

		dragAndDrop.perform();

		builder.dragAndDropBy(dashBoardGraphNewToDrag, -400,200).build().perform();

		Thread.sleep(2000);
		dashBoardGraphNewToDrag.click();
		
		Thread.sleep(2000);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashBoardCustomizeDashSave));
		dashBoardCustomizeDashSave.click();	
		Thread.sleep(2000);

		String expMessage = "Dashlet saved Successfully";
		String actMessage = checkValidationMessage(expMessage);
		

		Thread.sleep(2000);

		
		
		
		return true;
	}



	static String actBackupName;

	public static boolean TakeBackupForTransactions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException 
	{


		excelReader=new ExcelReader(POJOUtility.getExcelPath()); 

		File Efile=new File(getBaseDir()+"\\autoIt\\ExportFiles\\SanityBackupUptoMasters.fbak");

		if(Efile.exists())
		{
			Efile.delete();
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
		dataMangementMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(backup));
		backup.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(backupFileNameTxt));
		backupFileNameTxt.click();
		backupFileNameTxt.sendKeys(Keys.END);
		backupFileNameTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		backupFileNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 41, 5));

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(backupBtn));
		backupBtn.click();

		Thread.sleep(15000);

		actBackupName = checkDownloadedFileName(getDriver());
		String expBackupName = excelReader.getCellData(xlSheetName, 41, 6);
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actBackupName);

		Thread.sleep(2000);

		System.err.println("Backup Downloaded : "+actBackupName+"  Value Expected  "+expBackupName);

		if(actBackupName.equalsIgnoreCase(expBackupName))
		{
			excelReader.setCellData(xlfile, xlSheetName, 41, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 41, 8, resFail);
			return false;
		}
	}


	public SmokeDashboardPage(WebDriver driver)
	{	

		PageFactory.initElements(driver, this);

	}






}


