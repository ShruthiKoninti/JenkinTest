package com.focus.Pages;

import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class SmokeSecurityPage extends BaseEngine
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "SmokeSecurityPage";
	
	public static String checkValidationMessage(String ExpMessage) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		LoginPage lp=new LoginPage(getDriver());
		try
		{
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
			String actErrorMessage=errorMessage.getText();
			String expErrorMessage=ExpMessage;

			if(errorMessage.getText().isEmpty()==false)
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
				errorMessageCloseBtn.click();

				System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);

				return actErrorMessage;
			}
			else
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






	public static boolean checkPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));	
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyTitle));

		String actpasswordPolicyTitle=passwordPolicyTitle.getText();

		String exppasswordPolicyTitle = excelReader.getCellData(xlSheetName, 8, 6);	

		System.out.println("passwordPolicyTitle : "+actpasswordPolicyTitle+" Value Expected : "+exppasswordPolicyTitle);
		
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actpasswordPolicyTitle);


		if(actpasswordPolicyTitle.equalsIgnoreCase(exppasswordPolicyTitle))
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




	public static boolean checkPasswordPolicyTestSave() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyAddBtn));
		passwordPolicyAddBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(policyName));
		policyName.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLength));
		passwordLength.sendKeys(Keys.END);
		passwordLength.sendKeys(Keys.SHIFT, Keys.HOME);
		Thread.sleep(1000);
		passwordLength.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));

		SaveBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 9, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
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







	public static boolean checkPasswordPolicyLists() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		int passwordListCount = passwordList.size();
		System.out.println(passwordListCount);

		ArrayList<String> actpasswordList = new ArrayList<String>();


		for(int i=0;i< passwordListCount;i++)
		{
			String data = passwordList.get(i).getText();
			actpasswordList.add(data);
		}


		String actpasswordPolicyNameInList=actpasswordList.toString();

		String exppasswordPolicyNameInList =excelReader.getCellData(xlSheetName, 11, 6);

		System.out.println(actpasswordPolicyNameInList);
		System.out.println(exppasswordPolicyNameInList);
		
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actpasswordPolicyNameInList);

		if(actpasswordPolicyNameInList.equalsIgnoreCase(exppasswordPolicyNameInList))
		{
			System.out.println("Pass : Password policy Test is displaying");
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Fail : Password policy Test is not displaying");
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resFail);
			return false;
		}

	}







	public static boolean checkGetTheTestPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) 
		{
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 12, 5))) 
			{
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		String actpolicyName=policyName.getAttribute("value");
		String exppolicyName = excelReader.getCellData(xlSheetName, 12, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actpolicyName);

		String actpasswordLength=passwordLength.getAttribute("value");
		String exppasswordLength = excelReader.getCellData(xlSheetName, 13, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actpasswordLength);

		Thread.sleep(2000);

		Select complexity = new Select(passwordComplexity);

		List<WebElement> options=complexity.getOptions();

		WebElement comboBox = passwordComplexity;
		Select selectedValue = new Select(comboBox);
		String actpasswordComplexity = selectedValue.getFirstSelectedOption().getText();
		String exppasswordComplexity = excelReader.getCellData(xlSheetName, 14, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actpasswordComplexity);

		System.out.println("actpolicyName : "+actpolicyName);
		System.out.println("actpasswordLength : "+actpasswordLength);
		System.out.println("actpasswordComplexity :"+actpasswordComplexity);
		

		System.out.println("exppolicyName : "+exppolicyName);
		System.out.println("exppasswordLength : "+exppasswordLength);
		System.out.println("exppasswordComplexity :"+exppasswordComplexity);
		
		if(actpolicyName.equalsIgnoreCase(exppolicyName) && actpasswordLength.equalsIgnoreCase(exppasswordLength) 
				&& actpasswordComplexity.equalsIgnoreCase(exppasswordComplexity))
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







	public static boolean checkUpdateTestPasswordPolicy() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) 
		{
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 15, 5))) 
			{
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLength));
		passwordLength.click();
		passwordLength.sendKeys(Keys.END);
		passwordLength.sendKeys(Keys.SHIFT,Keys.HOME);
		passwordLength.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));		
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage=excelReader.getCellData(xlSheetName, 15, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
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





	public static boolean checkGetTheUpdatedPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) 
		{
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 17, 5))) 
			{
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordComplexity));

		//WebElement comboBox = passwordComplexity;
		Select selectedValue = new Select(passwordComplexity);
		String actpasswordComplexity = selectedValue.getFirstSelectedOption().getText();

		String exppasswordComplexity = excelReader.getCellData(xlSheetName, 17, 6);
		System.out.println("actpasswordComplexity  :  "+actpasswordComplexity);
		
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actpasswordComplexity);


		if(actpasswordComplexity.equalsIgnoreCase(exppasswordComplexity))
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






	public static boolean checkAlertOnDeletePasswordPolicy() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyAddBtn));
		passwordPolicyAddBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(policyName));
		policyName.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLength));
		passwordLength.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String actMessageSave=checkValidationMessage("Password policy created successfully");

		Thread.sleep(3000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) 
		{
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 20, 5))) 
			{
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyDeletebtn));
		passwordPolicyDeletebtn.click();

		Thread.sleep(2000);
		getWaitForAlert();
		

		String getExpectedCompMsg=excelReader.getCellData(xlSheetName, 18, 6);

		String actCompanySaveMsg=getDriver().switchTo().alert().getText();

		System.out.println("actCompanySaveMsg : "+actCompanySaveMsg);
		
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actCompanySaveMsg);


		getDriver().switchTo().alert().accept();
		Thread.sleep(2000);

		String expMessage=excelReader.getCellData(xlSheetName, 19, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actMessage);

		if(actCompanySaveMsg.equalsIgnoreCase(getExpectedCompMsg) && actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 18, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 18, 8, resFail);
			return false;
		}

	}











	public boolean CheckSavingAllProfileSTwithAlloptionsEnabled() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		createProfileProfileNameCombo.click();
		createProfileProfileNameCombo.sendKeys(excelReader.getCellData(xlSheetName, 21, 5));
		Thread.sleep(2000);
		createProfileProfileNameCombo.sendKeys(Keys.TAB);
		
		Thread.sleep(1000);

		int profileMenuPanelListCount = profileMenuPanelList.size();

		for(int i=0;i<profileMenuPanelListCount;i++)
		{
			profileMenuPanelList.get(i).click();

			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileSelectAllIcon));
			createProfileSelectAllIcon.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		System.out.println("***********************************************CheckSavingAllProfileSTwithAlloptionsEnabled************************************************");

		String expMessage=excelReader.getCellData(xlSheetName, 21, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actMessage);


		System.out.println("Saving Message of Profile : " +actMessage+" Value Expected : "+expMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileCloseIcon));
		createProfileCloseIcon.click();

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 21, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 21, 8, resFail);
			return false;
		}
	}








	public boolean checkCreateProfileLoadFromAllProfileAndValidatePurchasesVouchers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileLoadFromIcon));
		createProfileLoadFromIcon.click();

		String expMessage=excelReader.getCellData(xlSheetName, 22, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actMessage);


		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		Thread.sleep(2000);
		createProfileProfileNameCombo.click();
		Thread.sleep(2000);
		createProfileProfileNameCombo.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileLoadFromIcon));
		createProfileLoadFromIcon.click();

		Thread.sleep(2000);

		int profilesCount=createProfileLoadFromPopupProfilesList.size();

		System.out.println("Number of Profiles : "+profilesCount);

		for (int i = 0; i < profilesCount; i++) 
		{
			String actloadProfile=createProfileLoadFromPopupProfilesList.get(i).getText();

			System.out.println("loadProfile  : "+actloadProfile);

			if (actloadProfile.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 23, 5)))
			{
				createProfileLoadFromPopupProfilesList.get(i).click();

				break;
			}
			else
			{
				System.out.println(" Profile NOT Found ");
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadPopupOkBtn));
		loadPopupOkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileFinancalsExpandBtn));
		createProfileFinancalsExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsExpandBtn));
		financialsTransactionsExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchasesExpandBtn));
		financialsTransactionsPurchasesExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchasesVoucherOption));
		financialsTransactionsPurchasesVoucherOption.click();

		int restrictionTabChkBoxListCount = restrictionTabChkBoxList.size();

		boolean selected = false;

		for(int i=0;i<restrictionTabChkBoxListCount;i++)
		{
			selected = restrictionTabChkBoxList.get(i).isSelected();

			if(selected==false)
			{
				String notSelect = restrictionTabChkBoxNameList.get(i).getText();
				System.err.println("Un Selected CheckBox : "+notSelect);
			}
		}

		String actRestrictionTabCheckboxSelected = Boolean.toString(selected);

		String expRestrictionTabCheckboxSelected = excelReader.getCellData(xlSheetName, 23, 6);


		System.out.println("RestrictionTabCheckboxSelected   : "+actRestrictionTabCheckboxSelected+"  value expected  "+expRestrictionTabCheckboxSelected);
		
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actRestrictionTabCheckboxSelected);

		if (actMessage.equalsIgnoreCase(expMessage) && actRestrictionTabCheckboxSelected.equalsIgnoreCase(expRestrictionTabCheckboxSelected)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 22, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 22, 8, resFail);
			return false;
		}
	}









	public boolean checkCreateProfileSavingValidatingAndDeletingLoadedProfile() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 24, 6);

		String actMessage=checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actMessage);

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		createProfileProfileNameCombo.click();
		createProfileProfileNameCombo.sendKeys(Keys.SPACE);

		int actSize=createProfileProfileNameList.size();

		for (int i = 0; i < actSize; i++) 
		{
			String actProfileName=createProfileProfileNameList.get(i).getText();
			if (actProfileName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 24, 5)))
			{
				createProfileProfileNameList.get(i).click();

				createProfileProfileNameCombo.sendKeys(Keys.TAB);

				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileFinancalsExpandBtn));
		createProfileFinancalsExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsExpandBtn));
		financialsTransactionsExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchasesExpandBtn));
		financialsTransactionsPurchasesExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchasesVoucherOption));
		financialsTransactionsPurchasesVoucherOption.click();

		int restrictionTabChkBoxListCount = restrictionTabChkBoxList.size();

		boolean selected = false;

		for(int i=0;i<restrictionTabChkBoxListCount;i++)
		{
			selected = restrictionTabChkBoxList.get(i).isSelected();

			if(selected==false)
			{
				String notSelect = restrictionTabChkBoxNameList.get(i).getText();
				System.err.println("Un Selected CheckBox : "+notSelect);
			}
		}


		String actRestrictionTabCheckboxSelected = Boolean.toString(selected);

		String expRestrictionTabCheckboxSelected = excelReader.getCellData(xlSheetName, 25, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actRestrictionTabCheckboxSelected);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileDeleteIcon));
		createProfileDeleteIcon.click();

		getWaitForAlert();

		String actAlert=getAlert().getText();

		String expAlert=excelReader.getCellData(xlSheetName, 26, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actAlert);

		getAlert().accept();

		String expMessage1=excelReader.getCellData(xlSheetName, 27, 6);

		String actMessage1=checkValidationMessage(expMessage1);
		
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actMessage1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileDeleteIcon));
		createProfileProfileNameCombo.click();
		createProfileProfileNameCombo.sendKeys(Keys.SPACE);

		boolean Deleted=true;

		int sizeAfterDelete=createProfileProfileNameList.size();

		for (int i = 0; i < sizeAfterDelete; i++) 
		{
			String profileName=createProfileProfileNameList.get(i).getText();

			if (profileName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 25, 5))) 
			{
				System.out.println(" Profile Not Deleted ");

				Deleted=false;
			}
		}

		createProfileProfileNameCombo.sendKeys(Keys.TAB);

		String actProfileDeleted = Boolean.toString(Deleted);

		String expProfileDeleted =	excelReader.getCellData(xlSheetName, 28, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actProfileDeleted);

		System.out.println(actProfileDeleted);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileCloseIcon));
		createProfileCloseIcon.click();

		System.out.println("RestrictionTabCheckboxSelected     : *"+actRestrictionTabCheckboxSelected        +"*  value expected  "+expRestrictionTabCheckboxSelected);
		System.out.println("Alert Text                         : *"+actAlert                                 +"*  value expected  "+expAlert);
		System.out.println("Validation Message                 : *"+actMessage1                              +"*  value expected  "+expMessage1);
		System.out.println("Deleted Profile                    : *"+actProfileDeleted                        +"*  value expected  "+expProfileDeleted);

		if (actMessage.equalsIgnoreCase(expMessage) && actRestrictionTabCheckboxSelected.equalsIgnoreCase(expRestrictionTabCheckboxSelected) && actAlert.equalsIgnoreCase(expAlert) && 
				actMessage1.equalsIgnoreCase(expMessage1) && actProfileDeleted.equalsIgnoreCase(expProfileDeleted))
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











	





	public boolean  checkSaveingRoleWithAllfieldsInAllTabs() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityCreateRoleMenu));
		homeSecurityCreateRoleMenu.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();

		createRoleRoleNameCombo.sendKeys("Testing");
		createRoleRoleNameCombo.sendKeys(Keys.TAB);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRolePasswordPolicyDropdown));
		Select s=new Select(createRolePasswordPolicyDropdown);
		s.selectByVisibleText("Simple Policy");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(avaliableAllProfile));
		avaliableAllProfile.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleMOveFiledsFromLeftSideToRightSide));
		createRoleMOveFiledsFromLeftSideToRightSide.click();

		Thread.sleep(2000);

		int actAssignedProfilesListCount=assignedProfileList.size();

		Set<String> AssignedProfilesList=new HashSet<String>();

		for (int i = 0; i < actAssignedProfilesListCount; i++) 
		{
			String profile=assignedProfileList.get(i).getText();
			Thread.sleep(2000);
			AssignedProfilesList.add(profile);
		}

		String actAssignedProfilesList = AssignedProfilesList.toString();

		String expAssignedProfilesList = "[All Profile ST]";


		//Additions Tab
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(additionTab));
		additionTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleAddTabFinancialsMenuExpand));
		createRoleAddTabFinancialsMenuExpand.click();

		Thread.sleep(2000);

		int actFinancialMenusCount=additionTabFinancialsMenuList.size();

		Set<String> FinancialMenusList=new HashSet<String>();

		for (int i = 0; i < actFinancialMenusCount; i++) 
		{
			String profile=additionTabFinancialsMenuList.get(i).getText();

			FinancialMenusList.add(profile);
		}

		String actFinancialMenusList = FinancialMenusList.toString();

		String expFinancialMenusList = "[Transactions, Final Accounts, Credit Management, Reports, Receivable and Payable Analysis, Budget, Online Payments]";


		System.out.println(actFinancialMenusList);
		System.out.println(expFinancialMenusList);


		for (int i = 0; i < actFinancialMenusCount; i++) 
		{
			String text=additionTabFinancialsMenuList.get(i).getText();

			if (text.equalsIgnoreCase("Transactions"))
			{
				additionTabFinancialsMenuList.get(i).click();
			}
		}

		Thread.sleep(2000);

		int actTransactionsMenusCount=additionTabFinancialsTransactionMenusList.size();

		Set<String> TransactionsMenusList=new HashSet<String>();

		for (int i = 0; i < actTransactionsMenusCount; i++) 
		{
			String profile=additionTabFinancialsTransactionMenusList.get(i).getText();

			TransactionsMenusList.add(profile);
		}

		String actTransactionsMenusList = TransactionsMenusList.toString();

		String expTransactionsMenusList = "[Cash and Bank, Sales, Journals, Purchases, Auto Postings]";

		System.out.println(actTransactionsMenusList);
		System.out.println(expTransactionsMenusList);


		for (int i = 0; i < actTransactionsMenusCount; i++) 
		{
			String text=additionTabFinancialsTransactionMenusList.get(i).getText();

			if (text.equalsIgnoreCase("Purchases")) 
			{
				additionTabFinancialsTransactionMenusList.get(i).click();
			}
		}

		Thread.sleep(2000);

		int actPurchasesMenusCount=additionTabFinancialsTransactionsPurchasesMenusList.size();

		Set<String> PurchasesMenusList=new HashSet<String>();

		for (int i = 0; i < actPurchasesMenusCount; i++) 
		{
			String profile=additionTabFinancialsTransactionsPurchasesMenusList.get(i).getText();

			PurchasesMenusList.add(profile);
		}

		String actPurchasesMenusList = PurchasesMenusList.toString();

		String expPurchasesMenusList = "[Purchases Returns, Purchases Vouchers]";


		System.out.println(actPurchasesMenusList);
		System.out.println(expPurchasesMenusList);



		for (int i = 0; i < actPurchasesMenusCount; i++) 
		{
			String text=additionTabFinancialsTransactionsPurchasesMenusList.get(i).getText();

			if (text.equalsIgnoreCase("Purchases Vouchers")) 
			{
				additionTabFinancialsTransactionsPurchasesMenusList.get(i).click();
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleAddTabSelectAllChkBox));
		createRoleAddTabSelectAllChkBox.click();

		Thread.sleep(2000);


		//Exclusions Tab
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exclusionsTab));
		exclusionsTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleExclTabInventoryMenu));
		createRoleExclTabInventoryMenu.click();

		Thread.sleep(2000);

		int actInventoryMenusCount=exclusionsTabInventoryMenusList.size();

		Set<String> InventoryMenusList=new HashSet<String>();

		for (int i = 0; i < actInventoryMenusCount; i++) 
		{
			String menu=exclusionsTabInventoryMenusList.get(i).getText();

			InventoryMenusList.add(menu);
		}

		String actInventoryMenusList = InventoryMenusList.toString();

		String expInventoryMenusList = "[Transactions, Order Management, Reports]";


		System.out.println(actInventoryMenusList);
		System.out.println(expInventoryMenusList);


		for (int i = 0; i < actInventoryMenusCount; i++) 
		{
			String menu=exclusionsTabInventoryMenusList.get(i).getText();

			if (menu.equalsIgnoreCase("Transactions")) 
			{
				exclusionsTabInventoryMenusList.get(i).click();
			}
		}

		Thread.sleep(2000);

		int actInventoryTransactionsMenusCount=exclusionsTabInventoryTransactionsMenusList.size();

		Set<String> InventoryTransactionsMenusList=new HashSet<String>();

		for (int i = 0; i < actInventoryTransactionsMenusCount; i++) 
		{
			String menu=exclusionsTabInventoryTransactionsMenusList.get(i).getText();

			InventoryTransactionsMenusList.add(menu);
		}

		String actInventoryTransactionsMenusList = InventoryTransactionsMenusList.toString();

		String expInventoryTransactionsMenusList = "[Sales, Stocks, Hold/Unhold Stock, Stock Reconciliation, Purchases, Stock Allocation]";


		System.out.println(actInventoryTransactionsMenusList);
		System.out.println(expInventoryTransactionsMenusList);


		for (int i = 0; i < actInventoryTransactionsMenusCount; i++) 
		{
			String menu=exclusionsTabInventoryTransactionsMenusList.get(i).getText();

			if (menu.equalsIgnoreCase("Stock Allocation")) 
			{
				exclusionsTabInventoryTransactionsMenusList.get(i).click();
			}
		}

		Thread.sleep(2000);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleExclTabUnSelectAllChkBox));
		createRoleExclTabUnSelectAllChkBox.click();


		//Restriction For Entry Tab

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRestrictionForEntryTab));
		createRoleRestrictionForEntryTab.click();

		Thread.sleep(2000);

		int actRestrictionForEntryMastersListCount=createRoleRestrictionForEntryTabMastersList.size();

		for (int i = 0; i < actRestrictionForEntryMastersListCount; i++) 
		{
			String masters=createRoleRestrictionForEntryTabMastersList.get(i).getText();

			if (masters.equalsIgnoreCase("Account")) 
			{
				createRoleRestrictionForEntryTabMastersList.get(i).click();  
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRestrictionForEntryTabExclChkBox));
		createRoleRestrictionForEntryTabExclChkBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableSelectRow1Col1));
		restrictionForEntryTabTableSelectRow1Col1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableEnterMasterTxt));
		String v = "v";
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(v);

		int actSize=restrictionForEntryTabTableMasterComboList.size();

		for (int i = 0; i < actSize; i++) 
		{
			String actMasters=restrictionForEntryTabTableMasterComboList.get(i).getText();

			if (actMasters.equalsIgnoreCase("Vendor B")) 
			{
				restrictionForEntryTabTableMasterComboList.get(i).click();
				restrictionForEntryTabTableEnterMasterTxt.sendKeys(Keys.TAB);
				break;
			}
		}

		Thread.sleep(2000);

		restrictionForEntryTabRow1EntryChkBox.click();
		restrictionForEntryTabRow1ReportChkBox.click();
		restrictionForEntryTabRow1ViewChkBox.click();




		for (int i = 0; i < actRestrictionForEntryMastersListCount; i++) 
		{
			String masters=createRoleRestrictionForEntryTabMastersList.get(i).getText();

			if (masters.equalsIgnoreCase("Account")) 
			{
				createRoleRestrictionForEntryTabMastersList.get(i).click();  
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableSelectRow2Col1));
		restrictionForEntryTabTableSelectRow2Col1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableEnterMasterTxt));
		String c  = "Customer B";
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(c);

		int actSiz=restrictionForEntryTabTableMasterComboList.size();

		for (int i = 0; i < actSiz; i++) 
		{
			String actMasters=restrictionForEntryTabTableMasterComboList.get(i).getText();

			if (actMasters.equalsIgnoreCase("Customer B")) 
			{
				restrictionForEntryTabTableMasterComboList.get(i).click();
				restrictionForEntryTabTableEnterMasterTxt.sendKeys(Keys.TAB);
				break;
			}
		}

		Thread.sleep(2000);

		restrictionForEntryTabRow2EntryChkBox.click();
		restrictionForEntryTabRow2ReportChkBox.click();
		restrictionForEntryTabRow2ViewChkBox.click();



		//Restriction For Trees
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionforTreesTab));
		restrictionforTreesTab.click();

		Thread.sleep(2000);

		int actRestrictionForTreesMastersListCount=createRoleRestrictionForTreesTabMastersList.size();

		for (int i = 0; i < actRestrictionForTreesMastersListCount; i++) 
		{
			String masters=createRoleRestrictionForTreesTabMastersList.get(i).getText();

			if (masters.equalsIgnoreCase("Account")) 
			{
				createRoleRestrictionForTreesTabMastersList.get(i).click();  
			}
		}


		restrictionForTreesTabDefaultChkBox.click();

		// Transaction Rights Tab
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionRightsTab));
		transactionRightsTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(allowBillWiseOnAccountChkBox));
		allowBillWiseOnAccountChkBox.click();

		Select budgetLimitWarning= new Select(budgetLimitWarningDropdown);
		budgetLimitWarning.selectByVisibleText("Allow");

		Select creditLimitWarning= new Select(creditLimitWarningDropdown);
		creditLimitWarning.selectByVisibleText("Alert");

		Select negativeCashCheck= new Select(negativeCashCheckDropdown);
		//negativeCashCheck.selectByVisibleText("Allow");
		negativeCashCheck.selectByVisibleText("Alert");


		Select negativeStockCheck= new Select(negativeStockCheckDropdown);
		negativeStockCheck.selectByVisibleText("Allow");

		cantaddfutureTransChkbox.click();

		cannotAddTransThatAreMoreThanTxt.click();
		cannotAddTransThatAreMoreThanTxt.clear();
		cannotAddTransThatAreMoreThanTxt.sendKeys("5");

		cannotEditPreviousMonthEntriesAfterTxt.click();
		cannotEditPreviousMonthEntriesAfterTxt.clear();
		cannotEditPreviousMonthEntriesAfterTxt.sendKeys("15");

		CantEditTransthatAreremorethanTxt.click();
		CantEditTransthatAreremorethanTxt.clear();
		CantEditTransthatAreremorethanTxt.sendKeys("50");

		CantAddPrevMnthEntriesaftTxt.click();
		CantAddPrevMnthEntriesaftTxt.clear();
		CantAddPrevMnthEntriesaftTxt.sendKeys("1");

		CantPrintAfterValueTxt.click();
		CantPrintAfterValueTxt.clear();
		CantPrintAfterValueTxt.sendKeys("12");

		Select CantPrintAfterValue= new Select(cantPrintAfterValueDropdown);
		CantPrintAfterValue.selectByVisibleText("Hours");

		CantRePrintAfterValueTxt.click();
		CantRePrintAfterValueTxt.clear();
		CantRePrintAfterValueTxt.sendKeys("1");

		Select CantRePrintAfterValue= new Select(CantRePrintAfterDropdown);
		CantRePrintAfterValue.selectByVisibleText("Days");

		doNotAllowMasterCustomization.click();



		//AI Rights
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AIRightsTab));
		AIRightsTab.click();

		Thread.sleep(2000);

		AIRightsEnableVUIChkBox.click();

		AIRightsEnableAIChkBox.click();

		NoOfVoiceCommandsPermittedTxt.click();
		NoOfVoiceCommandsPermittedTxt.clear();
		NoOfVoiceCommandsPermittedTxt.sendKeys("3");

		NoOfTextCommandsPermittedTxt.click();
		NoOfTextCommandsPermittedTxt.clear();
		NoOfTextCommandsPermittedTxt.sendKeys("5");

		AIRightsSalesChkBox.click();
		AIRightsSalesOrderChkBox.click();
		AIRightsStockChkBox.click();
		AIRightsPurchasesChkBox.click();
		AIRightsPurchasesOrdersChkBox.click();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
		createRoleSaveicon.click();

		String expMessage="Role saved Successfully";

		String actMessage=checkValidationMessage(expMessage);


		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleCloseIcon));
		//createRoleCloseIcon.click();

		System.out.println("actAssignedProfilesList : "+actAssignedProfilesList);
		System.out.println("expAssignedProfilesList : "+expAssignedProfilesList);



		System.out.println("actFinancialMenusList : "+actFinancialMenusList);
		System.out.println("expFinancialMenusList : "+expFinancialMenusList);

		System.out.println("actTransactionsMenusList : "+actTransactionsMenusList);
		System.out.println("expTransactionsMenusList : "+expTransactionsMenusList);

		System.out.println("actPurchasesMenusList : "+actPurchasesMenusList);
		System.out.println("expPurchasesMenusList : "+expPurchasesMenusList);

		System.out.println("actInventoryMenusList : "+actInventoryMenusList);
		System.out.println("expInventoryMenusList : "+expInventoryMenusList);

		System.out.println("actInventoryTransactionsMenusList : "+actInventoryTransactionsMenusList);
		System.out.println("expInventoryTransactionsMenusList : "+expInventoryTransactionsMenusList);


		System.out.println("Error Message : "+actMessage+"  value expected  "+expMessage);

		if (actAssignedProfilesList.equalsIgnoreCase(expAssignedProfilesList) &&   actFinancialMenusList.equalsIgnoreCase(expFinancialMenusList) &&
				actTransactionsMenusList.equalsIgnoreCase(expTransactionsMenusList) && actPurchasesMenusList.equalsIgnoreCase(expPurchasesMenusList) && 
				actInventoryMenusList.equalsIgnoreCase(expInventoryMenusList) && actInventoryTransactionsMenusList.equalsIgnoreCase(expInventoryTransactionsMenusList) && actMessage.equalsIgnoreCase(expMessage)) 
		{

			return true;
		} 
		else 
		{
			return false;
		}
	}










	public boolean  checkEditAndUpdateSavedRole() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		createRoleRoleNameCombo.clear();
		createRoleRoleNameCombo.sendKeys("Testing");
		Thread.sleep(2000);
		createRoleRoleNameCombo.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assignedProfilesTab));
		assignedProfilesTab.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(avaliableAllProfileSt));
		avaliableAllProfileSt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleMOveFiledsFromLeftSideToRightSide));
		createRoleMOveFiledsFromLeftSideToRightSide.click();


		//Additions Tab
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(additionTab));
		additionTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addFininicalExpandBtn));
		addFininicalExpandBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addFinTransationExpandBtn));
		addFinTransationExpandBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addFinTransationpurchaseExpandBtn));
		addFinTransationpurchaseExpandBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addFinTransPurPurchasesVouchers));
		addFinTransPurPurchasesVouchers.click();


		int actPurchasesMenusCount=additionTabFinancialsTransactionsPurchasesMenusList.size();
		int expPurchasesMenusCount=2;

		System.out.println("additionTabFinancialsTransactionsPurchasesMenusList : "+actPurchasesMenusCount);
		System.out.println("additionTabFinancialsTransactionsPurchasesMenusList : "+expPurchasesMenusCount);



		boolean actCRAddAlwaysSuspendOnSavingChkBox				=CRAddAlwaysSuspendOnSavingChkBox.isDisplayed();
		boolean actCRAddHideMenuChkBox							=CRAddHideMenuChkBox.isDisplayed();



		boolean expCRAddAlwaysSuspendOnSavingChkBox				=true;
		boolean expCRAddHideMenuChkBox							=true;

		System.out.println("CRAddAlwaysSuspendOnSavingChkBox           : "+actCRAddAlwaysSuspendOnSavingChkBox          +"  value expected  "+expCRAddAlwaysSuspendOnSavingChkBox);
		System.out.println("CRAddHideMenuChkBox                        : "+actCRAddHideMenuChkBox                       +"  value expected  "+expCRAddHideMenuChkBox);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleAddTabSelectAllChkBox));
		createRoleAddTabSelectAllChkBox.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionRightsTab));
		transactionRightsTab.click();

		cannotAddTransThatAreMoreThanTxt.click();
		cannotAddTransThatAreMoreThanTxt.clear();
		cannotAddTransThatAreMoreThanTxt.sendKeys("1");

		cannotEditPreviousMonthEntriesAfterTxt.click();	
		cannotEditPreviousMonthEntriesAfterTxt.clear();
		cannotEditPreviousMonthEntriesAfterTxt.sendKeys("7");

		CantEditTransthatAreremorethanTxt.click();
		CantEditTransthatAreremorethanTxt.clear();
		CantEditTransthatAreremorethanTxt.sendKeys("20");

		CantAddPrevMnthEntriesaftTxt.click();
		CantAddPrevMnthEntriesaftTxt.clear();
		CantAddPrevMnthEntriesaftTxt.sendKeys("0");

		CantPrintAfterValueTxt.click();
		CantPrintAfterValueTxt.clear();
		CantPrintAfterValueTxt.sendKeys("1");

		Select CantPrintAfterValue= new Select(cantPrintAfterValueDropdown);
		CantPrintAfterValue.selectByVisibleText("Days");

		CantRePrintAfterValueTxt.click();
		CantRePrintAfterValueTxt.clear();
		CantRePrintAfterValueTxt.sendKeys("5");

		Select CantRePrintAfterValue= new Select(CantRePrintAfterDropdown);
		CantRePrintAfterValue.selectByVisibleText("Hours");

		doNotAllowMasterCustomization.click();


		//AI Rights
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AIRightsTab));
		AIRightsTab.click();

		Thread.sleep(2000);


		AIRightsEnableVUIChkBox.click();

		AIRightsEnableAIChkBox.click();

		doNotApplyTheTagResTrictionToTheResultsChkBox.click();

		NoOfVoiceCommandsPermittedTxt.click();
		NoOfVoiceCommandsPermittedTxt.clear();
		NoOfVoiceCommandsPermittedTxt.sendKeys("1");

		NoOfTextCommandsPermittedTxt.click();
		NoOfTextCommandsPermittedTxt.clear();
		NoOfTextCommandsPermittedTxt.sendKeys("1");

		AIRightsSalesOrderChkBox.click();
		AIRightsPurchasesOrdersChkBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
		createRoleSaveicon.click();

		String expMessage = "Role updated Successfully";
		String actMessage = checkValidationMessage(expMessage);

		Thread.sleep(2000);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		} 
		else 
		{
			return false;
		}
	}










	public boolean  checkCompareAndDeleteUpdatedRole() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.sendKeys("RoleToDelete");
		Thread.sleep(2000);
		createRoleRoleNameCombo.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRolePasswordPolicyDropdown));
		createRolePasswordPolicyDropdown.click();
		createRolePasswordPolicyDropdown.sendKeys("Simple");
		createRolePasswordPolicyDropdown.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(avaliableAllProfileSt));
		avaliableAllProfileSt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleMOveFiledsFromLeftSideToRightSide));
		createRoleMOveFiledsFromLeftSideToRightSide.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
		createRoleSaveicon.click();

		checkValidationMessage("Role saved Successfully");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		createRoleRoleNameCombo.sendKeys(Keys.END);
		createRoleRoleNameCombo.sendKeys(Keys.SHIFT,Keys.END);
		createRoleRoleNameCombo.sendKeys("RoleToDelete");
		Thread.sleep(1000);
		createRoleRoleNameCombo.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CreateRoleDeleteButtonImage));
		getAction().moveToElement(CreateRoleDeleteButtonImage).click().build().perform();

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CreateRoleDeleteButtonImage));
		CreateRoleDeleteButtonImage.click();
*/
		getWaitForAlert();

		String actAlertText = getAlert().getText();
		String expAlertText = "are you sure that you want to delete Role";

		getAlert().accept();
		Thread.sleep(2000);

		String expMessage = "Role Successfully deleted..";
		String actMessage = checkValidationMessage(expMessage);


		System.out.println("AlertText  : "+actAlertText+"  value expected  "+expAlertText);
		System.out.println("Error Message  : "+actMessage+"  value expected  "+expMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleCloseIcon));
		createRoleCloseIcon.click();

		if ( actAlertText.equalsIgnoreCase(expAlertText) && actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		} 
		else 
		{
			return false;
		}
	}





	


	public static boolean checkCreateUserWithAllMandatoryFileds() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
				Thread.sleep(3000);
				getDriver().navigate().refresh();
				Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
		createUserIcon.click();
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
		createUserLoginNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 29, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		createUserERPRoleDropDown.sendKeys(excelReader.getCellData(xlSheetName, 30, 5));

		int count = createUserERPRolesDropdownOptions.size()-1;
		System.out.println("Count "+ count);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 31, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 32, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		createUserNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 33, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
		createUserLoginAbbrivationTxt.sendKeys(excelReader.getCellData(xlSheetName, 34, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityQuestionDropDown));
		Select s = new Select(createUserSecurityQuestionDropDown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 35, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));	
		createUserSecurityAnswerTxt.sendKeys(excelReader.getCellData(xlSheetName, 36, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
		int counnt = createUserLanguageDropdownList.size()-1;
		String actLanguageDropdownCount = Integer.toString(counnt);
		String expLanguageDropdownCount = excelReader.getCellData(xlSheetName, 29, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actLanguageDropdownCount);

		Set LanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < counnt; i++) 
		{
			String data = createUserLanguageDropdownList.get(i).getText();
			LanguageDropdownList.add(data);
		}

		String actLanguageDropdownList = LanguageDropdownList.toString();
		String expLanguageDropdownList = excelReader.getCellData(xlSheetName, 30, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actLanguageDropdownList);

		System.out.println("LanguageDropdownCount  : "+actLanguageDropdownCount  +"  value expected  "+expLanguageDropdownCount);
		System.out.println("LanguageDropdownList   : "+actLanguageDropdownList   +"  value expected  "+expLanguageDropdownList);

		Select language = new Select(getLanguage);
		language.selectByIndex(1);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
		int cont = createUserAltLanguageDropdownList.size()-1;
		String actAltLanguageDropdownCount = Integer.toString(count);
		String expAltLanguageDropdownCount =excelReader.getCellData(xlSheetName, 31, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actAltLanguageDropdownCount);

		Set AltLanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < cont; i++) 
		{
			String data = createUserAltLanguageDropdownList.get(i).getText();
			AltLanguageDropdownList.add(data);
		}

		String actAltLanguageDropdownList = AltLanguageDropdownList.toString();
		String expAltLanguageDropdownList = excelReader.getCellData(xlSheetName, 32, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actAltLanguageDropdownList);

		Select altlanguage = new Select(getAltLanguage);
		altlanguage.selectByIndex(2);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userCreationTimeZoneDrpdwn));
		Select timeZone = new Select(userCreationTimeZoneDrpdwn);
		timeZone.selectByVisibleText(excelReader.getCellData(xlSheetName, 37, 5));
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
		createUserAdditionalInfoTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserUserTypeDrpDwn));
		Select user = new Select(createUserUserTypeDrpDwn);
		user.selectByVisibleText(excelReader.getCellData(xlSheetName, 38, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailTxt));
		createUserEmailTxt.click();
		createUserEmailTxt.clear();
		createUserEmailTxt.sendKeys(excelReader.getCellData(xlSheetName, 39, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 33, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);


		if (actMessage.equalsIgnoreCase(expMessage)) 
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











	public static boolean checkLogoutAndLoginWithCreatedUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp=new LoginPage(getDriver());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 40, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 41, 5));
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		
		Thread.sleep(2000);
		lp.reLogin(excelReader.getCellData(xlSheetName, 40, 5),excelReader.getCellData(xlSheetName, 41, 5),"Automation Company ‎");
		

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDemoDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 40, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actUserInfo);


		Thread.sleep(2000);;

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 41, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actLoginCompanyName);
		companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ."+actUserInfo          +".  value expected  ."+expUserInfo+".");
		System.out.println("LoginCompanyName       : ."+actLoginCompanyName  +".  value expected  ."+expLoginCompanyName+".");

		if(actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName))  
		{
			System.out.println("***Test Pass: Login Successfull***");
			 excelReader.setCellData(xlfile, xlSheetName, 40, 8, resPass);
			return true;
		}
		else
		{

			System.out.println("***Test Fail: Login Not Successfull***");
			 excelReader.setCellData(xlfile, xlSheetName, 40, 8, resFail);
			return false;
		}
	}










	public static boolean checkUpdatingTheCreatedUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount=createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) 
		{
			String actUser=createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 44, 5))) 
			{
				createUserSavedUsersList.get(i).click();

				Thread.sleep(3000);

				break;
			}
		}
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.click();
		createUserPasswordTxt.sendKeys(Keys.END);
		createUserPasswordTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		createUserPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 45, 5));
		createUserPasswordTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.click();
		createUserConfirmPasswordTxt.sendKeys(Keys.END);
		createUserConfirmPasswordTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		createUserConfirmPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 46, 5));
		createUserConfirmPasswordTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		createUserNameTxt.click();
		createUserNameTxt.sendKeys(Keys.END);
		createUserNameTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		createUserNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 47, 5));
		
		Thread.sleep(2000);

		/*for (int i = 0; i < 5; i++) 
		{
			if (validationConfirmationMessage.getText().isEmpty()==false) 
			{
				checkValidationMessage("");
			}
		}
		
		Thread.sleep(2000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		getAction().moveToElement(SaveBtn).click().build().perform();

		String expMessage = excelReader.getCellData(xlSheetName, 44, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 44, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 44, 8, resFail);
			return false;
		}
	}	










	public boolean checkDeletingUpdatedUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount=createUserSavedUsersList.size();

		for (int i = 0; i < actUsersCount; i++) 
		{
			String actUser=createUserSavedUsersList.get(i).getText();
			System.out.println(actUser);

			if (actUser.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 48, 5)))
			{
				createUserSavedUsersList.get(i).click();

				Thread.sleep(2000);
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileDeleteIcon));
		createProfileDeleteIcon.click();

		getWaitForAlert();

		getAlert().accept();

		String expMessage = excelReader.getCellData(xlSheetName, 48, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 48, 8, resPass);

			return true;
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 48, 8, resFail);
			return false;
		}
	}











	public boolean checInputMandatoryFieldsInCreateUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
		createUserIcon.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
		createUserLoginNameTxt.click();
		createUserLoginNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 49, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		createUserERPRoleDropDown.click();
		createUserERPRoleDropDown.sendKeys(excelReader.getCellData(xlSheetName, 50, 5));

		int count = createUserERPRolesDropdownOptions.size()-1;
		System.out.println("Count "+ count);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 51, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 52, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		createUserNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 53, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
		createUserLoginAbbrivationTxt.sendKeys(excelReader.getCellData(xlSheetName, 54, 5));

		createUserSecurityQuestionDropDown.click();
		Select s = new Select(createUserSecurityQuestionDropDown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 55, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));	
		createUserSecurityAnswerTxt.sendKeys(excelReader.getCellData(xlSheetName, 56, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
		int counnt = createUserLanguageDropdownList.size()-1;
		String actLanguageDropdownCount = Integer.toString(counnt);
		String expLanguageDropdownCount = excelReader.getCellData(xlSheetName, 49, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actLanguageDropdownCount);

		Set LanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < counnt; i++) 
		{
			String data = createUserLanguageDropdownList.get(i).getText();
			LanguageDropdownList.add(data);
		}

		String actLanguageDropdownList = LanguageDropdownList.toString();
		String expLanguageDropdownList = excelReader.getCellData(xlSheetName, 50, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actLanguageDropdownList);

		System.out.println("LanguageDropdownCount  : "+actLanguageDropdownCount  +"  value expected  "+expLanguageDropdownCount);
		System.out.println("LanguageDropdownList   : "+actLanguageDropdownList   +"  value expected  "+expLanguageDropdownList);

		Select language = new Select(getLanguage);
		language.selectByVisibleText(excelReader.getCellData(xlSheetName, 57, 5));
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
		int cont = createUserAltLanguageDropdownList.size()-1;
		String actAltLanguageDropdownCount = Integer.toString(count);
		String expAltLanguageDropdownCount =excelReader.getCellData(xlSheetName, 51, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actAltLanguageDropdownCount);

		Set AltLanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < cont; i++) 
		{
			String data = createUserAltLanguageDropdownList.get(i).getText();
			AltLanguageDropdownList.add(data);
		}

		String actAltLanguageDropdownList = AltLanguageDropdownList.toString();
		String expAltLanguageDropdownList = excelReader.getCellData(xlSheetName, 52, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actAltLanguageDropdownList);

		Select altlanguage = new Select(getAltLanguage);
		altlanguage.selectByVisibleText(excelReader.getCellData(xlSheetName, 58, 5));
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userCreationTimeZoneDrpdwn));
		Select timeZone = new Select(userCreationTimeZoneDrpdwn);
		timeZone.selectByVisibleText(excelReader.getCellData(xlSheetName, 59, 5));
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
		createUserAdditionalInfoTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserUserTypeDrpDwn));
		Select user = new Select(createUserUserTypeDrpDwn);
		user.selectByVisibleText(excelReader.getCellData(xlSheetName, 60, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailTxt));
		createUserEmailTxt.click();
		createUserEmailTxt.clear();
		createUserEmailTxt.sendKeys(excelReader.getCellData(xlSheetName, 61, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 53, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 49, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 49, 8, resFail);
			return false;
		}
	}








	public static boolean checkCreateUserWithRestrictionsForAccount() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
		createUserIcon.click();
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
		createUserLoginNameTxt.click();
		createUserLoginNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 62, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		createUserERPRoleDropDown.click();
		createUserERPRoleDropDown.sendKeys(excelReader.getCellData(xlSheetName, 63, 5));

		int count = createUserERPRolesDropdownOptions.size()-1;
		System.out.println("Count "+ count);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 64, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 65, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		createUserNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 66, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
		createUserLoginAbbrivationTxt.sendKeys(excelReader.getCellData(xlSheetName, 67, 5));

		createUserSecurityQuestionDropDown.click();
		Select s = new Select(createUserSecurityQuestionDropDown);
		s.selectByVisibleText(excelReader.getCellData(xlSheetName, 68, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));	
		createUserSecurityAnswerTxt.sendKeys(excelReader.getCellData(xlSheetName, 69, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
		int counnt = createUserLanguageDropdownList.size()-1;
		String actLanguageDropdownCount = Integer.toString(counnt);
		String expLanguageDropdownCount = excelReader.getCellData(xlSheetName, 62, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 62, 7, actLanguageDropdownCount);

		Set LanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < counnt; i++) 
		{
			String data = createUserLanguageDropdownList.get(i).getText();
			LanguageDropdownList.add(data);
		}

		String actLanguageDropdownList = LanguageDropdownList.toString();
		String expLanguageDropdownList = excelReader.getCellData(xlSheetName, 63, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actLanguageDropdownList);

		System.out.println("LanguageDropdownCount  : "+actLanguageDropdownCount  +"  value expected  "+expLanguageDropdownCount);
		System.out.println("LanguageDropdownList   : "+actLanguageDropdownList   +"  value expected  "+expLanguageDropdownList);

		Select language = new Select(getLanguage);
		language.selectByVisibleText(excelReader.getCellData(xlSheetName, 70, 5));
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
		int cont = createUserAltLanguageDropdownList.size()-1;
		String actAltLanguageDropdownCount = Integer.toString(count);
		String expAltLanguageDropdownCount = excelReader.getCellData(xlSheetName, 64, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 64, 7, actAltLanguageDropdownCount);

		Set AltLanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < cont; i++) 
		{
			String data = createUserAltLanguageDropdownList.get(i).getText();
			AltLanguageDropdownList.add(data);
		}

		String actAltLanguageDropdownList = AltLanguageDropdownList.toString();
		String expAltLanguageDropdownList = excelReader.getCellData(xlSheetName, 65, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actAltLanguageDropdownList);

		Select altlanguage = new Select(getAltLanguage);
		altlanguage.selectByVisibleText(excelReader.getCellData(xlSheetName, 71, 5));
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userCreationTimeZoneDrpdwn));
		Select timeZone = new Select(userCreationTimeZoneDrpdwn);
		timeZone.selectByVisibleText(excelReader.getCellData(xlSheetName, 72, 5));
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
		createUserAdditionalInfoTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserUserTypeDrpDwn));
		Select user = new Select(createUserUserTypeDrpDwn);
		user.selectByVisibleText(excelReader.getCellData(xlSheetName, 73, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailTxt));
		createUserEmailTxt.click();
		createUserEmailTxt.clear();
		createUserEmailTxt.sendKeys(excelReader.getCellData(xlSheetName, 74, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserRestrictionforentryTab));
		createUserRestrictionforentryTab.click();

		Thread.sleep(2000);

		int actRestrictionForEntryMastersListCount=createRoleRestrictionForEntryTabMastersList.size();

		for (int i = 0; i < actRestrictionForEntryMastersListCount; i++) 
		{
			String masters=createRoleRestrictionForEntryTabMastersList.get(i).getText();

			if (masters.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 75, 5))) 
			{
				createRoleRestrictionForEntryTabMastersList.get(i).click();  
				System.err.println("**Account is Selected**");
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRestrictionForEntryTabExclChkBox));
		createRoleRestrictionForEntryTabExclChkBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableSelectRow1Col1));
		restrictionForEntryTabTableSelectRow1Col1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableEnterMasterTxt));
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(excelReader.getCellData(xlSheetName, 76, 5));
		Thread.sleep(2000);
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(Keys.TAB);

		restrictionForEntryTabRow1EntryChkBox.click();
		restrictionForEntryTabRow1ReportChkBox.click();
		restrictionForEntryTabRow1ViewChkBox.click();


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableSelectRow2Col1));
		restrictionForEntryTabTableSelectRow2Col1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableEnterMasterTxt));
		
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(excelReader.getCellData(xlSheetName, 77, 5));
		Thread.sleep(2000);
      
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(Keys.TAB); 
		
		restrictionForEntryTabRow2EntryChkBox.click();
		restrictionForEntryTabRow2ReportChkBox.click();
		restrictionForEntryTabRow2ViewChkBox.click();



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 66, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 66, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);


		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 62, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 62, 8, resFail);
			return false;
		}
	}










	public boolean checkLogoutAndLoginWithNewlySavedUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp=new LoginPage(getDriver());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 78, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 79, 5));
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);
		lp.reLogin(excelReader.getCellData(xlSheetName, 78, 5),excelReader.getCellData(xlSheetName, 79, 5),"Automation Company");

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameTestDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 78, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 78, 7, actUserInfo);

		Thread.sleep(2000);;

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 79, 6);
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actLoginCompanyName);
		companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ."+actUserInfo          +".  value expected  ."+expUserInfo+".");
		System.out.println("LoginCompanyName       : ."+actLoginCompanyName  +".  value expected  ."+expLoginCompanyName+".");

		if(actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName))  
		{
			excelReader.setCellData(xlfile, xlSheetName, 78, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 78, 8, resFail);
			return false;
		}
	}








	public boolean checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp=new LoginPage(getDriver());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 42, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 43, 5));

		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();
		
		Thread.sleep(2000);
		lp.reLogin(excelReader.getCellData(xlSheetName, 42, 5),excelReader.getCellData(xlSheetName, 43, 5),"Automation Company");

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 42, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 42, 7, actUserInfo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));

		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 43, 6);
		companyLogo.click();

		System.out.println("UserInfo               : ."+actUserInfo          +".  value expected  ."+expUserInfo+".");
		System.out.println("LoginCompanyName       : ."+actLoginCompanyName  +".  value expected  ."+expLoginCompanyName+".");
		
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actLoginCompanyName);

		if(actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName))  
		{
			System.out.println("***Test Pass: Login Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resPass);
			return true;
		}
		else
		{

			System.out.println("***Test Fail: Login Not Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resFail);
			return false;
		}
	}











	public boolean checkEditingAndUpdatingSavedUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{ 
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount=createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) 
		{
			String actUser=createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 80, 5))) 
			{
				createUserSavedUsersList.get(i).click();

				Thread.sleep(4000);

				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 81, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 82, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		createUserNameTxt.clear();
		createUserNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 83, 5));

		for (int i = 0; i < 5; i++) 
		{
			if (validationConfirmationMessage.getText().isEmpty()==false) 
			{
				checkValidationMessage("");
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 80, 6);
		String actMessage = checkValidationMessage(expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 80, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			 excelReader.setCellData(xlfile, xlSheetName, 80, 8, resPass);
			return true;
		}
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			 excelReader.setCellData(xlfile, xlSheetName, 80, 8, resFail);
			return false;
		}
	} 






	public boolean checkLogoutAndLoginWithUpdatedUserWithPreviousPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp=new LoginPage(getDriver());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		userNameDisplayLogo.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 84, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 85, 5));

		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();
		lp.reLogin(excelReader.getCellData(xlSheetName, 84, 5),excelReader.getCellData(xlSheetName, 85, 5),"Automation Company");
		

		String actMessage = invalidPasswordMsgs.getText();
		String expMessage = excelReader.getCellData(xlSheetName, 84, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, actMessage);

		System.out.println("mandatoryMsgs : "+actMessage+"  value expected  "+expMessage);

		if(actMessage.equalsIgnoreCase(expMessage))  
		{
			excelReader.setCellData(xlfile, xlSheetName, 84, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 84, 8, resFail);
			return false;
		}
	}








	public boolean checkLoginWithUpdatedPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp=new LoginPage(getDriver());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 86, 5));

		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();
		
		Thread.sleep(2000);
		lp.reLogin(excelReader.getCellData(xlSheetName, 84, 5),excelReader.getCellData(xlSheetName, 86, 5),"Automation Company");

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameTestDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 86, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 86, 7, actUserInfo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));

		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 87, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actLoginCompanyName);

		companyLogo.click();

		System.out.println("UserInfo               : "+actUserInfo            +"  value expected  "+expUserInfo);
		System.out.println("LoginCompanyName       : "+actLoginCompanyName    +"  value expected  "+expLoginCompanyName);

		if(actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName))  
		{
			System.out.println("***Test Pass: Login Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 86, 8, resPass);

			return true;
		}
		else
		{

			System.out.println("***Test Fail: Login Not Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 86, 8, resFail);

			return false;
		}
	}









	public static boolean checkPurchasesVoucherForRestrictedAccountsInUserThroughRole() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesExpandBtn));
		purchasesExpandBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersBtn));
		purchaseVouchersBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen Opened");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountComboExpand));
		vendorAccountComboExpand.click();

		int vendorCount = vendorAccountList.size();
		boolean vendor = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < vendorCount; i++)
		{
			data = vendorAccountList.get(i).getText();
			accountslist.add(data);
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 88, 5)))
			{
				vendor=false;
				break;
			}
		}

		System.out.println("Vendor Accounts: "  +accountslist);
		
		String actual   = Boolean.toString(vendor);
		String expected = excelReader.getCellData(xlSheetName, 88, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 88, 7, actual);

		if (actual.equalsIgnoreCase(expected)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 88, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 88, 8, resFail);
			return false;
		}
	}










	public static boolean checkSalesInvoicesVATForRestrictedAccountsInUserThroughRole() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesMenu));
		salesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoicesVatBtn));
		salesInvoicesVatBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen Opened");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountComboExpand));
		customerAccountComboExpand.click();


		int customerCount = customerAccountList.size();
		boolean customer = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < customerCount; i++)
		{
			data = customerAccountList.get(i).getText();
			accountslist.add(data);
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 89, 5)))
			{
				customer=false;
				break;
			}
		}

		System.out.println("Customer  Accounts: "  +accountslist);
		
		
		String actual   = Boolean.toString(customer);
		String expected = excelReader.getCellData(xlSheetName, 89, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actual);

		if (actual.equalsIgnoreCase(expected))
		{
			excelReader.setCellData(xlfile, xlSheetName, 89, 8, resPass);
			return true;

		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 89, 8, resFail);
			return false;
		}

	}








	public static boolean checkLedgerForRestrictedAccountsInUserThroughRole() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsReportsMenu));
		financialsReportsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ledger));
		ledger.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));		
		accountSearchTextArea.sendKeys(excelReader.getCellData(xlSheetName, 90, 5));
		accountSearchTextArea.sendKeys(Keys.ENTER);


		int accountsCount = accountsListInLedger.size();
		boolean vendor = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < accountsCount; i++)
		{
			data = accountsListInLedger.get(i).getText();
			accountslist.add(data);
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 91, 5)))
			{
				vendor=false;
				break;
			}
		}


		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));	
		accountSearchTextArea.clear();
		Thread.sleep(2000);
		accountSearchTextArea.sendKeys(excelReader.getCellData(xlSheetName, 92, 5));
		accountSearchTextArea.sendKeys(Keys.ENTER);

		boolean customer = true;
		String data1 = null;

		for (int i = 0; i < accountsCount; i++)
		{
			data1 = accountsListInLedger.get(i).getText();
			accountslist.add(data);
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 93, 5)))
			{
				customer=false;
				break;
			}
		}
		
		
		String actual   = Boolean.toString(vendor);
		String expected = excelReader.getCellData(xlSheetName, 90, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 90, 7, actual);
		
		String actual1   = Boolean.toString(customer);
		String expected1 = excelReader.getCellData(xlSheetName, 91, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 91, 7, actual1);
		
		

		if (actual.equalsIgnoreCase(expected) && actual1.equalsIgnoreCase(expected1)) 
		{

			excelReader.setCellData(xlfile, xlSheetName, 90, 8, resPass);
			return true;
		} 
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 90, 8, resFail);
			return false;
		}
	}








	public static boolean checkLogoutAndLoginWithCreatedRestrictedUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp=new LoginPage(getDriver());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 94, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 95, 5));
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();

		Thread.sleep(2000);
		lp.reLogin(excelReader.getCellData(xlSheetName, 94, 5),excelReader.getCellData(xlSheetName, 95, 5),"Automation Company");
		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameUserRestrictionDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 94, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actUserInfo);

		Thread.sleep(2000);;

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 95, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 95, 7, actLoginCompanyName);
		companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ."+actUserInfo          +".  value expected  ."+expUserInfo+".");
		System.out.println("LoginCompanyName       : ."+actLoginCompanyName  +".  value expected  ."+expLoginCompanyName+".");

		if(actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName))  
		{
			System.out.println("***Test Pass: Login Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 94, 8, resPass);
			return true;
		}
		else
		{

			System.out.println("***Test Fail: Login Not Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 94, 8, resFail);
			return false;
		}
	}









	public static boolean checkPurchasesVocherForRestrictedAccountThroughUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesExpandBtn));
		purchasesExpandBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersBtn));
		purchaseVouchersBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen Opened");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountComboExpand));
		vendorAccountComboExpand.click();


		int vendorCount = vendorAccountList.size();
		boolean vendor = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < vendorCount; i++)
		{
			data = vendorAccountList.get(i).getText();
			accountslist.add(data);
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 96, 5)))
			{
				vendor=false;
				break;
			}
		}

		System.out.println("Vendor Accounts: "  +accountslist);

		String actual   = Boolean.toString(vendor);
		String expected = excelReader.getCellData(xlSheetName, 96, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actual);

		if (actual.equalsIgnoreCase(expected)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 96, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 96, 8, resFail);
			return false;
		}
	}











	public static boolean checkSalesInvoiceVATForRestrictedAccountsThroughUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesMenu));
		salesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoicesVatBtn));
		salesInvoicesVatBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		checkValidationMessage("Screen Opened");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountComboExpand));
		customerAccountComboExpand.click();


		int customerCount = customerAccountList.size();
		boolean customer = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < customerCount; i++)
		{
			data = customerAccountList.get(i).getText();
			accountslist.add(data);
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 97, 5)))
			{
				customer=false;
				break;
			}
		}

		System.out.println("Customer  Accounts: "  +accountslist);

		String actual   = Boolean.toString(customer);
		String expected = excelReader.getCellData(xlSheetName, 97, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actual);

		if (actual.equalsIgnoreCase(expected))
		{
			excelReader.setCellData(xlfile, xlSheetName, 97, 8, resPass);
			return true;

		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 97, 8, resFail);
			return false;
		}
	}










	public static boolean checkLedgerForRestrictedAccountThroughUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsReportsMenu));
		financialsReportsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ledger));
		ledger.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));		
		accountSearchTextArea.sendKeys(excelReader.getCellData(xlSheetName, 98, 5));
		accountSearchTextArea.sendKeys(Keys.ENTER);


		int accountsCount = accountsListInLedger.size();
		boolean vendor = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < accountsCount; i++)
		{
			data = accountsListInLedger.get(i).getText();
			accountslist.add(data);
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 99, 5)))
			{
				vendor=false;
				break;
			}
		}

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));	
		accountSearchTextArea.clear();
		Thread.sleep(2000);
		accountSearchTextArea.sendKeys(excelReader.getCellData(xlSheetName, 100, 5));
		accountSearchTextArea.sendKeys(Keys.ENTER);



		boolean customer = true;
		String data1 = null;

		for (int i = 0; i < accountsCount; i++)
		{
			data1 = accountsListInLedger.get(i).getText();
			accountslist.add(data);
			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 101, 5)))
			{
				customer=false;
				break;
			}
		}

		String actual   = Boolean.toString(vendor);
		String expected = excelReader.getCellData(xlSheetName, 98, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actual);
		
		String actual1   = Boolean.toString(customer);
		String expected1 = excelReader.getCellData(xlSheetName, 99, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 99, 7, actual1);
		
		

		if (actual.equalsIgnoreCase(expected) && actual1.equalsIgnoreCase(expected1)) 
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








	public boolean checkLogoutAndLoginWithSUToUpdatedUserDeleteUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp=new LoginPage(getDriver());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 102, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 103, 5));

		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();

		Thread.sleep(2000);
		lp.reLogin(excelReader.getCellData(xlSheetName, 102, 5),excelReader.getCellData(xlSheetName, 103, 5),"Automation Company");
		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 102, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actUserInfo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));

		companyLogo.click();

		String getCompanyTxt = companyName.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 103, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 103, 7, actLoginCompanyName);

		companyLogo.click();

		System.out.println("UserInfo               : "+actUserInfo            +"  value expected  "+expUserInfo);
		System.out.println("LoginCompanyName       : "+actLoginCompanyName    +"  value expected  "+expLoginCompanyName);

		if(actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName))  
		{
			System.out.println("***Test Pass: Login Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 102, 8, resPass);

			return true;
		}
		else
		{
			System.out.println("***Test Fail: Login Not Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 102, 8, resFail);

			return false;
		}
	}


///////////////////////////////////////////////
	public boolean checkChangePasswordatUserLevelwithNewPassword() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOption));
		changePasswordOption.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOldPasswordTxt));
		changePasswordOldPasswordTxt.sendKeys("su");
		changePasswordOldPasswordTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordNewPasswordTxt));
		changePasswordNewPasswordTxt.sendKeys("1234");
		changePasswordNewPasswordTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordConfirmTxt));
		changePasswordConfirmTxt.sendKeys("1234");
		changePasswordConfirmTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordSaveBtn));
		changePasswordSaveBtn.click();
		
		String expMessage="Password Saved Successfully";
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
	
	public boolean checkLoginangLogoutSUwithChangePassword() throws InterruptedException
	{
		
LoginPage lp=new LoginPage(getDriver());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("su");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("1234");

		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();
		
		Thread.sleep(2000);
		lp.reLogin("su","1234","Automation Company");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = "";
		if(actUserInfo.equalsIgnoreCase(expUserInfo))
		{
			System.out.println("Test Pass with Change Password");
			return true;
			
		}
		else
				
		{
			return false;
		}
		

	}

////////
	public boolean checkChangePasswordatUserLevelwithOldPassword() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOption));
		changePasswordOption.click();
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOldPasswordTxt));
		changePasswordOldPasswordTxt.sendKeys("su");
		changePasswordOldPasswordTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordNewPasswordTxt));
		changePasswordNewPasswordTxt.sendKeys("su");
		changePasswordNewPasswordTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordConfirmTxt));
		changePasswordConfirmTxt.sendKeys("su");
		changePasswordConfirmTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordSaveBtn));
		changePasswordSaveBtn.click();
		
		String expMessage="Password Saved Successfully";
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
	
	public boolean checkLoginangLogoutSUwithChangePasswordUpdate() throws InterruptedException
	{
		
LoginPage lp=new LoginPage(getDriver());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("su");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("su");

		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();
		
		Thread.sleep(2000);
		lp.reLogin("su","su","Automation Company");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = "SU";
		if(actUserInfo.equalsIgnoreCase(expUserInfo))
		{
			System.out.println("Test Pass with Change Password");
			return true;
			
		}
		else
				
		{
			return false;
		}
		

	}


























	public SmokeSecurityPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}


}	

