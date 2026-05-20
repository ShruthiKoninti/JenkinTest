package com.focus.Pages;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.xalan.xsltc.compiler.sym;
import org.openqa.selenium.By;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Set;
import java.util.regex.Pattern;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

import io.github.bonigarcia.wdm.WebDriverManager;

public class SmokeSecurityPage extends BaseEngine {
	private static String resPass = "Pass";
	private static String resFail = "Fail";
	private static ExcelReader excelReader;
	private static String xlfile = getBaseDir() + "\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "SmokeSecurityPage";

	public static String checkValidationMessage(String ExpMessage)
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		LoginPage lp = new LoginPage(getDriver());
		try {
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
			String actErrorMessage = errorMessage.getText();
			String expErrorMessage = ExpMessage;

			if (errorMessage.getText().isEmpty() == false) {
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
				errorMessageCloseBtn.click();

				System.out.println("ValidationMessage  :  " + actErrorMessage + " Value Expected : " + expErrorMessage);

				return actErrorMessage;
			} else {

				System.out.println("ValidationMessage  :  " + actErrorMessage + " Value Expected : " + expErrorMessage);

				return actErrorMessage;
			}
		} catch (Exception e) {
			System.err.println("Error Message NOT Found or NOT Clickable");
			System.err.println(e.getMessage());

			String Exception = e.getMessage();

			return Exception;
		}
	}

	public static boolean checkPasswordPolicy()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyTitle));

		String actpasswordPolicyTitle = passwordPolicyTitle.getText();

		String exppasswordPolicyTitle = excelReader.getCellData(xlSheetName, 8, 6);

		System.out.println(
				"passwordPolicyTitle : " + actpasswordPolicyTitle + " Value Expected : " + exppasswordPolicyTitle);

		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actpasswordPolicyTitle);

		if (actpasswordPolicyTitle.equalsIgnoreCase(exppasswordPolicyTitle)) {
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resPass);

			return true;
		} else {

			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resFail);
			return false;
		}
	}

	public static boolean checkPasswordPolicyTestSave()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyAddBtn));
		getAction().moveToElement(passwordPolicyAddBtn).build().perform();
		ScrollToElement(passwordPolicyAddBtn);

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

		String expMessage = excelReader.getCellData(xlSheetName, 9, 6);

		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			excelReader.setCellData(xlfile, xlSheetName, 9, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 9, 8, resFail);

			return false;
		}
	}

	public static boolean checkPasswordPolicyLists()
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		int passwordListCount = passwordList.size();
		System.out.println(passwordListCount);

		ArrayList<String> actpasswordList = new ArrayList<String>();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();
			actpasswordList.add(data);
		}

		String actpasswordPolicyNameInList = actpasswordList.toString();

		String exppasswordPolicyNameInList = excelReader.getCellData(xlSheetName, 11, 6);

		System.out.println(actpasswordPolicyNameInList);
		System.out.println(exppasswordPolicyNameInList);

		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actpasswordPolicyNameInList);

		if (actpasswordPolicyNameInList.equalsIgnoreCase(exppasswordPolicyNameInList)) {
			System.out.println("Pass : Password policy Test is displaying");
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resPass);
			return true;
		} else {
			System.out.println("Fail : Password policy Test is not displaying");
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resFail);
			return false;
		}

	}

	public static boolean checkGetTheTestPasswordPolicy()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 12, 5))) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		String actpolicyName = policyName.getAttribute("value");
		String exppolicyName = excelReader.getCellData(xlSheetName, 12, 6);

		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actpolicyName);

		String actpasswordLength = passwordLength.getAttribute("value");
		String exppasswordLength = excelReader.getCellData(xlSheetName, 13, 6);

		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actpasswordLength);

		Thread.sleep(2000);

		Select complexity = new Select(passwordComplexity);

		List<WebElement> options = complexity.getOptions();

		WebElement comboBox = passwordComplexity;
		Select selectedValue = new Select(comboBox);
		String actpasswordComplexity = selectedValue.getFirstSelectedOption().getText();
		String exppasswordComplexity = excelReader.getCellData(xlSheetName, 14, 6);

		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actpasswordComplexity);

		System.out.println("actpolicyName : " + actpolicyName);
		System.out.println("actpasswordLength : " + actpasswordLength);
		System.out.println("actpasswordComplexity :" + actpasswordComplexity);

		System.out.println("exppolicyName : " + exppolicyName);
		System.out.println("exppasswordLength : " + exppasswordLength);
		System.out.println("exppasswordComplexity :" + exppasswordComplexity);

		if (actpolicyName.equalsIgnoreCase(exppolicyName) && actpasswordLength.equalsIgnoreCase(exppasswordLength)
				&& actpasswordComplexity.equalsIgnoreCase(exppasswordComplexity)) {
			excelReader.setCellData(xlfile, xlSheetName, 12, 8, resPass);

			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 12, 8, resFail);

			return false;
		}
	}

	public static boolean checkUpdateTestPasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 15, 5))) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLength));
		passwordLength.click();
		passwordLength.sendKeys(Keys.END);
		passwordLength.sendKeys(Keys.SHIFT, Keys.HOME);
		passwordLength.sendKeys(excelReader.getCellData(xlSheetName, 16, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = excelReader.getCellData(xlSheetName, 15, 6);

		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			excelReader.setCellData(xlfile, xlSheetName, 15, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 15, 8, resFail);
			return false;
		}
	}

	public static boolean checkGetTheUpdatedPasswordPolicy()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 17, 5))) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordComplexity));

		// WebElement comboBox = passwordComplexity;
		Select selectedValue = new Select(passwordComplexity);
		String actpasswordComplexity = selectedValue.getFirstSelectedOption().getText();

		String exppasswordComplexity = excelReader.getCellData(xlSheetName, 17, 6);
		System.out.println("actpasswordComplexity  :  " + actpasswordComplexity);

		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actpasswordComplexity);

		if (actpasswordComplexity.equalsIgnoreCase(exppasswordComplexity)) {
			excelReader.setCellData(xlfile, xlSheetName, 17, 8, resPass);

			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 17, 8, resFail);

			return false;
		}
	}

	public static boolean checkAlertOnDeletePasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyAddBtn));
		passwordPolicyAddBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(policyName));
		policyName.sendKeys(excelReader.getCellData(xlSheetName, 18, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLength));
		passwordLength.sendKeys(excelReader.getCellData(xlSheetName, 19, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String actMessageSave = checkValidationMessage("Password policy created successfully");

		Thread.sleep(3000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 20, 5))) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyDeletebtn));
		passwordPolicyDeletebtn.click();

		Thread.sleep(2000);
		getWaitForAlert();

		String getExpectedCompMsg = excelReader.getCellData(xlSheetName, 18, 6);

		String actCompanySaveMsg = getDriver().switchTo().alert().getText();

		System.out.println("actCompanySaveMsg : " + actCompanySaveMsg);

		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actCompanySaveMsg);

		getDriver().switchTo().alert().accept();
		Thread.sleep(2000);

		String expMessage = excelReader.getCellData(xlSheetName, 19, 6);

		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actMessage);

		if (actCompanySaveMsg.equalsIgnoreCase(getExpectedCompMsg) && actMessage.equalsIgnoreCase(expMessage)) {
			excelReader.setCellData(xlfile, xlSheetName, 18, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 18, 8, resFail);
			return false;
		}

	}

	public boolean CheckSavingAllProfileSTwithAlloptionsEnabled()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();

		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		createProfileProfileNameCombo.click();
		createProfileProfileNameCombo.sendKeys(excelReader.getCellData(xlSheetName, 21, 5));
		Thread.sleep(4000);
		createProfileProfileNameCombo.sendKeys(Keys.TAB);

		Thread.sleep(10000);

		int profileMenuPanelListCount = profileMenuPanelList.size();

		for (int i = 0; i < profileMenuPanelListCount; i++) {
			profileMenuPanelList.get(i).click();

			Thread.sleep(1000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileSelectAllIcon));
			createProfileSelectAllIcon.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		System.out.println(
				"***********************************************CheckSavingAllProfileSTwithAlloptionsEnabled************************************************");

		String expMessage = excelReader.getCellData(xlSheetName, 21, 6);

		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actMessage);

		System.out.println("Saving Message of Profile : " + actMessage + " Value Expected : " + expMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileCloseIcon));
		createProfileCloseIcon.click();

		if (actMessage.equalsIgnoreCase(expMessage)) {
			excelReader.setCellData(xlfile, xlSheetName, 21, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 21, 8, resFail);
			return false;
		}
	}

	public boolean checkCreateProfileLoadFromAllProfileAndValidatePurchasesVouchers()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileLoadFromIcon));
		createProfileLoadFromIcon.click();

		String expMessage = excelReader.getCellData(xlSheetName, 22, 6);

		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actMessage);

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		Thread.sleep(2000);
		createProfileProfileNameCombo.click();
		Thread.sleep(2000);
		createProfileProfileNameCombo.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileLoadFromIcon));
		createProfileLoadFromIcon.click();

		Thread.sleep(4000);

		int profilesCount = createProfileLoadFromPopupProfilesList.size();

		System.out.println("Number of Profiles : " + profilesCount);

		for (int i = 0; i < profilesCount; i++) {
			String actloadProfile = createProfileLoadFromPopupProfilesList.get(i).getText();

			System.out.println("loadProfile  : " + actloadProfile);

			if (actloadProfile.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 23, 5))) {
				createProfileLoadFromPopupProfilesList.get(i).click();

				break;
			} else {
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

		for (int i = 0; i < restrictionTabChkBoxListCount; i++) {
			selected = restrictionTabChkBoxList.get(i).isSelected();

			if (selected == false) {
				String notSelect = restrictionTabChkBoxNameList.get(i).getText();
				System.err.println("Un Selected CheckBox : " + notSelect);
			}
		}

		String actRestrictionTabCheckboxSelected = Boolean.toString(selected);

		String expRestrictionTabCheckboxSelected = excelReader.getCellData(xlSheetName, 23, 6);

		System.out.println("RestrictionTabCheckboxSelected   : " + actRestrictionTabCheckboxSelected
				+ "  value expected  " + expRestrictionTabCheckboxSelected);

		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actRestrictionTabCheckboxSelected);

		if (actMessage.equalsIgnoreCase(expMessage)
				&& actRestrictionTabCheckboxSelected.equalsIgnoreCase(expRestrictionTabCheckboxSelected)) {
			excelReader.setCellData(xlfile, xlSheetName, 22, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 22, 8, resFail);
			return false;
		}
	}

	public boolean checkCreateProfileSavingValidatingAndDeletingLoadedProfile()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 24, 6);

		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actMessage);

		Thread.sleep(6000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();

		Thread.sleep(10000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		createProfileProfileNameCombo.click();
		createProfileProfileNameCombo.sendKeys(Keys.SPACE);

		int actSize = createProfileProfileNameList.size();

		for (int i = 0; i < actSize; i++) {
			String actProfileName = createProfileProfileNameList.get(i).getText();
			if (actProfileName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 24, 5))) {
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

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchasesExpandBtn));
		financialsTransactionsPurchasesExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchasesVoucherOption));
		financialsTransactionsPurchasesVoucherOption.click();

		int restrictionTabChkBoxListCount = restrictionTabChkBoxList.size();

		boolean selected = false;

		for (int i = 0; i < restrictionTabChkBoxListCount; i++) {
			selected = restrictionTabChkBoxList.get(i).isSelected();

			if (selected == false) {
				String notSelect = restrictionTabChkBoxNameList.get(i).getText();
				System.err.println("Un Selected CheckBox : " + notSelect);
			}
		}

		String actRestrictionTabCheckboxSelected = Boolean.toString(selected);

		String expRestrictionTabCheckboxSelected = excelReader.getCellData(xlSheetName, 25, 6);

		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actRestrictionTabCheckboxSelected);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileDeleteIcon));
		createProfileDeleteIcon.click();

		getWaitForAlert();

		String actAlert = getAlert().getText();

		String expAlert = excelReader.getCellData(xlSheetName, 26, 6);

		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actAlert);

		getAlert().accept();

		String expMessage1 = excelReader.getCellData(xlSheetName, 27, 6);

		String actMessage1 = checkValidationMessage(expMessage1);

		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actMessage1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileDeleteIcon));
		createProfileProfileNameCombo.click();
		createProfileProfileNameCombo.sendKeys(Keys.SPACE);

		boolean Deleted = true;

		int sizeAfterDelete = createProfileProfileNameList.size();

		for (int i = 0; i < sizeAfterDelete; i++) {
			String profileName = createProfileProfileNameList.get(i).getText();

			if (profileName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 25, 5))) {
				System.out.println(" Profile Not Deleted ");

				Deleted = false;
			}
		}

		createProfileProfileNameCombo.sendKeys(Keys.TAB);

		String actProfileDeleted = Boolean.toString(Deleted);

		String expProfileDeleted = excelReader.getCellData(xlSheetName, 28, 6);

		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actProfileDeleted);

		System.out.println(actProfileDeleted);
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileCloseIcon));
		createProfileCloseIcon.click();

		System.out.println("RestrictionTabCheckboxSelected     : *" + actRestrictionTabCheckboxSelected
				+ "*  value expected  " + expRestrictionTabCheckboxSelected);
		System.out.println("Alert Text                         : *" + actAlert + "*  value expected  " + expAlert);
		System.out
				.println("Validation Message                 : *" + actMessage1 + "*  value expected  " + expMessage1);
		System.out.println("Deleted Profile                    : *" + actProfileDeleted + "*  value expected  "
				+ expProfileDeleted);

		if (actMessage.equalsIgnoreCase(expMessage)
				&& actRestrictionTabCheckboxSelected.equalsIgnoreCase(expRestrictionTabCheckboxSelected)
				&& actAlert.equalsIgnoreCase(expAlert) && actMessage1.equalsIgnoreCase(expMessage1)
				&& actProfileDeleted.equalsIgnoreCase(expProfileDeleted)) {
			excelReader.setCellData(xlfile, xlSheetName, 24, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 24, 8, resFail);
			return false;
		}
	}

	public boolean checkSaveingRoleWithAllfieldsInAllTabs()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityCreateRoleMenu));
		homeSecurityCreateRoleMenu.click();

		Thread.sleep(6000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();

		createRoleRoleNameCombo.sendKeys("Testing");
		createRoleRoleNameCombo.sendKeys(Keys.TAB);
		Thread.sleep(6000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRolePasswordPolicyDropdown));
		Select s = new Select(createRolePasswordPolicyDropdown);
		s.selectByVisibleText("Simple Policy");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(avaliableAllProfile));
		avaliableAllProfile.click();

		Thread.sleep(2000);

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(createRoleMOveFiledsFromLeftSideToRightSide));
		createRoleMOveFiledsFromLeftSideToRightSide.click();

		Thread.sleep(2000);

		int actAssignedProfilesListCount = assignedProfileList.size();

		Set<String> AssignedProfilesList = new HashSet<String>();

		for (int i = 0; i < actAssignedProfilesListCount; i++) {
			String profile = assignedProfileList.get(i).getText();
			Thread.sleep(2000);
			AssignedProfilesList.add(profile);
		}

		String actAssignedProfilesList = AssignedProfilesList.toString();

		String expAssignedProfilesList = "[All Profile ST]";
		
		System.out.println(actAssignedProfilesList);
		System.out.println(expAssignedProfilesList);

		// Additions Tab
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(additionTab));
		additionTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleAddTabFinancialsMenuExpand));
		createRoleAddTabFinancialsMenuExpand.click();

		Thread.sleep(2000);

		int actFinancialMenusCount = AdditionTabFinancialsMenuList.size();

		Set<String> FinancialMenusList = new HashSet<String>();

		for (int i = 0; i < actFinancialMenusCount; i++) {
			String profile = AdditionTabFinancialsMenuList.get(i).getText();

			FinancialMenusList.add(profile);
		}

		String actFinancialMenusList = FinancialMenusList.toString();

		String expFinancialMenusList = "[Transactions, Final Accounts, Credit Management, Reports, Receivable and Payable Analysis, Budget, Online Payments]";

		System.out.println(actFinancialMenusList);
		System.out.println(expFinancialMenusList);

		for (int i = 0; i < actFinancialMenusCount; i++) {
			String text = AdditionTabFinancialsMenuList.get(i).getText();

			if (text.equalsIgnoreCase("Transactions")) {
				AdditionTabFinancialsMenuList.get(i).click();
			}
		}

		Thread.sleep(2000);

		int actTransactionsMenusCount = AdditionTabFinancialsTransactionMenusList.size();

		Set<String> TransactionsMenusList = new HashSet<String>();

		for (int i = 0; i < actTransactionsMenusCount; i++) {
			String profile = AdditionTabFinancialsTransactionMenusList.get(i).getText();

			TransactionsMenusList.add(profile);
		}

		String actTransactionsMenusList = TransactionsMenusList.toString();

		String expTransactionsMenusList = "[Cash and Bank, Sales, Journals, Purchases, Auto Postings]";

		System.out.println(actTransactionsMenusList);
		System.out.println(expTransactionsMenusList);

		for (int i = 0; i < actTransactionsMenusCount; i++) {
			String text = AdditionTabFinancialsTransactionMenusList.get(i).getText();

			if (text.equalsIgnoreCase("Purchases")) {
				AdditionTabFinancialsTransactionMenusList.get(i).click();
			}
		}

		Thread.sleep(2000);

		int actPurchasesMenusCount = AdditionTabFinancialsTransactionsPurchasesMenusList.size();

		Set<String> PurchasesMenusList = new HashSet<String>();

		for (int i = 0; i < actPurchasesMenusCount; i++) {
			String profile = AdditionTabFinancialsTransactionsPurchasesMenusList.get(i).getText();

			PurchasesMenusList.add(profile);
		}

		String actPurchasesMenusList = PurchasesMenusList.toString();

		String expPurchasesMenusList = "[Purchases Returns, Purchases Vouchers]";

		System.out.println(actPurchasesMenusList);
		System.out.println(expPurchasesMenusList);

		for (int i = 0; i < actPurchasesMenusCount; i++) {
			String text = AdditionTabFinancialsTransactionsPurchasesMenusList.get(i).getText();

			if (text.equalsIgnoreCase("Purchases Vouchers")) {
				AdditionTabFinancialsTransactionsPurchasesMenusList.get(i).click();
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleAddTabSelectAllChkBox));
		createRoleAddTabSelectAllChkBox.click();

		Thread.sleep(2000);

		// Exclusions Tab
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exclusionsTab));
		exclusionsTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExpandExclTabInventoryMenu));
		ExpandExclTabInventoryMenu.click();

		Thread.sleep(5000);

		int actInventoryMenusCount = ExpandExcluInventoryList.size();

		Set<String> InventoryMenusList = new HashSet<String>();

		for (int i = 0; i < actInventoryMenusCount; i++) {
			String menu = ExpandExcluInventoryList.get(i).getText();

			InventoryMenusList.add(menu);
		}

		String actInventoryMenusList = InventoryMenusList.toString();

		String expInventoryMenusList = "[Transactions, Order Management, Reports]";

		System.out.println(actInventoryMenusList);
		System.out.println(expInventoryMenusList);
		
		Thread.sleep(5000);
		
	
		
		for (int i = 0; i < actInventoryMenusCount; i++) 
			
		
		{
			String pp = ExpandExcluInventoryList.get(i).getText();

			
			if (pp.equalsIgnoreCase("Transactions")) {
				
				
				ExpandExcluInventoryList.get(i).click();
				break;
			}
		}
		
		ClickUsingJs(Transaction_expand);

		Thread.sleep(2000);

		int actInventoryTransactionsMenusCount = ExpandExcluInventoryTransactionsList.size();

		Set<String> InventoryTransactionsMenusList = new HashSet<String>();

		for (int i = 0; i < actInventoryTransactionsMenusCount; i++) {
			String menu = ExpandExcluInventoryTransactionsList.get(i).getText();

			InventoryTransactionsMenusList.add(menu);
		}

		String actInventoryTransactionsMenusList = InventoryTransactionsMenusList.toString();

		String expInventoryTransactionsMenusList = "[Sales, Stocks, Hold/Unhold Stock, Stock Reconciliation, Purchases]";

		System.out.println(actInventoryTransactionsMenusList);
		System.out.println(expInventoryTransactionsMenusList);

		for (int i = 0; i < actInventoryTransactionsMenusCount; i++) {
			String menu = ExpandExcluInventoryTransactionsList.get(i).getText();

			if (menu.equalsIgnoreCase("Stock Reconciliation")) {
				ExpandExcluInventoryTransactionsList.get(i).click();
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleExclTabUnSelectAllChkBox));
		createRoleExclTabUnSelectAllChkBox.click();

		// Restriction For Entry Tab

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRestrictionForEntryTab));
		createRoleRestrictionForEntryTab.click();

		Thread.sleep(2000);

		int actRestrictionForEntryMastersListCount = createRoleRestrictionForEntryTabMastersList.size();

		for (int i = 0; i < actRestrictionForEntryMastersListCount; i++) {
			String masters = createRoleRestrictionForEntryTabMastersList.get(i).getText();

			if (masters.equalsIgnoreCase("Account")) {
				createRoleRestrictionForEntryTabMastersList.get(i).click();
			}
		}

		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRestrictionForEntryTabExclChkBox));
		//createRoleRestrictionForEntryTabExclChkBox.click();
		
		if(createRoleRestrictionForEntryTabCheckExclChkBox.isSelected()==false)
		{
			createRoleRestrictionForEntryTabExclChkBox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableSelectRow1Col1));
		restrictionForEntryTabTableSelectRow1Col1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableEnterMasterTxt));
		String v = "v";
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
		Thread.sleep(1000);
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(v);

		int actSize = restrictionForEntryTabTableMasterComboList.size();

		for (int i = 0; i < actSize; i++) {
			String actMasters = restrictionForEntryTabTableMasterComboList.get(i).getText();

			if (actMasters.equalsIgnoreCase("Vendor B")) {
				restrictionForEntryTabTableMasterComboList.get(i).click();
				restrictionForEntryTabTableEnterMasterTxt.sendKeys(Keys.TAB);
				break;
			}
		}

		Thread.sleep(2000);
		
		if(restrictionForEntryTabRow1EntryCheckChkBox.isSelected()==false)
		{
			restrictionForEntryTabRow1EntryChkBox.click();
		}
		
		if(restrictionForEntryTabRow1ReportCheckChkBox.isSelected()==false)
		{
			restrictionForEntryTabRow1ReportChkBox.click();
		}
		
		if(restrictionForEntryTabRow1ViewCheckChkBox.isSelected()==false)
		{
			restrictionForEntryTabRow1ViewChkBox.click();
		}

		//restrictionForEntryTabRow1EntryChkBox.click();
		//restrictionForEntryTabRow1ReportChkBox.click();
		//restrictionForEntryTabRow1ViewChkBox.click();

		for (int i = 0; i < actRestrictionForEntryMastersListCount; i++) {
			String masters = createRoleRestrictionForEntryTabMastersList.get(i).getText();

			if (masters.equalsIgnoreCase("Account")) {
				createRoleRestrictionForEntryTabMastersList.get(i).click();
			}
		}

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableSelectRow2Col1));
		restrictionForEntryTabTableSelectRow2Col1.click();

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableEnterMasterTxt));
		String c = "Customer C";
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(Keys.HOME,Keys.SHIFT,Keys.END);
		Thread.sleep(1000);
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(c);

		int actSiz = restrictionForEntryTabTableMasterComboList.size();

		for (int i = 0; i < actSiz; i++) {
			String actMasters = restrictionForEntryTabTableMasterComboList.get(i).getText();

			if (actMasters.equalsIgnoreCase("Customer C")) {
				restrictionForEntryTabTableMasterComboList.get(i).click();
				restrictionForEntryTabTableEnterMasterTxt.sendKeys(Keys.TAB);
				break;
			}
		}

		Thread.sleep(2000);
		
		if(restrictionForEntryTabRow2EntryCheckChkBox.isSelected()==false)
		{
			restrictionForEntryTabRow2EntryChkBox.click();
		}
		
		if(restrictionForEntryTabRow2ReportCheckChkBox.isSelected()==false)
		{
			restrictionForEntryTabRow2ReportChkBox.click();
		}
		
		if(restrictionForEntryTabRow2ViewCheckChkBox.isSelected()==false)
		{
			restrictionForEntryTabRow2ViewChkBox.click();
		}

		//restrictionForEntryTabRow2EntryChkBox.click();
		//restrictionForEntryTabRow2ReportChkBox.click();
		//restrictionForEntryTabRow2ViewChkBox.click();

		// Restriction For Trees
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionforTreesTab));
		restrictionforTreesTab.click();

		Thread.sleep(2000);

		int actRestrictionForTreesMastersListCount = createRoleRestrictionForTreesTabMastersList.size();

		for (int i = 0; i < actRestrictionForTreesMastersListCount; i++) {
			String masters = createRoleRestrictionForTreesTabMastersList.get(i).getText();

			if (masters.equalsIgnoreCase("Account")) {
				createRoleRestrictionForTreesTabMastersList.get(i).click();
			}
		}
		
		if(restrictionForTreesTabCheckDefaultChkBox.isSelected()==false)
		{
			restrictionForTreesTabDefaultChkBox.click();
		}
		
		//restrictionForTreesTabDefaultChkBox.click();

		// Transaction Rights Tab
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionRightsTab));
		transactionRightsTab.click();

		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(allowBillWiseOnAccountChkBox));
		//allowBillWiseOnAccountChkBox.click();
		
		if(CheckallowBillWiseOnAccountChkBox.isSelected()==false)
		{
			allowBillWiseOnAccountChkBox.click();
		}

		Select budgetLimitWarning = new Select(budgetLimitWarningDropdown);
		budgetLimitWarning.selectByVisibleText("Allow");

		Select creditLimitWarning = new Select(creditLimitWarningDropdown);
		creditLimitWarning.selectByVisibleText("Alert");

		Select negativeCashCheck = new Select(negativeCashCheckDropdown);
		// negativeCashCheck.selectByVisibleText("Allow");
		negativeCashCheck.selectByVisibleText("Alert");

		Select negativeStockCheck = new Select(negativeStockCheckDropdown);
		negativeStockCheck.selectByVisibleText("Allow");
		
		if(cantaddfutureTransChkbox.isSelected()==false)
		{
			cantaddfutureTransChkbox.click();
		}
		//cantaddfutureTransChkbox.click();

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

		Select CantPrintAfterValue = new Select(cantPrintAfterValueDropdown);
		CantPrintAfterValue.selectByVisibleText("Hours");

		CantRePrintAfterValueTxt.click();
		CantRePrintAfterValueTxt.clear();
		CantRePrintAfterValueTxt.sendKeys("1");

		Select CantRePrintAfterValue = new Select(CantRePrintAfterDropdown);
		CantRePrintAfterValue.selectByVisibleText("Days");

		//doNotAllowMasterCustomization.click();
		if(CheckdoNotAllowMasterCustomization.isSelected()==false)
		{
			doNotAllowMasterCustomization.click();
		}

		// AI Rights
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AIRightsTab));
		AIRightsTab.click();

		Thread.sleep(2000);

		//AIRightsEnableVUIChkBox.click();

		//AIRightsEnableAIChkBox.click();
		
	/*	if(CheckAIRightsEnableVUIChkBox.isSelected()==false)
		{
			AIRightsEnableVUIChkBox.click();
		}
		
		if(CheckAIRightsEnableAIChkBox.isSelected()==false)
		{
			AIRightsEnableAIChkBox.click();
		}

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
		Thread.sleep(1000);*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
		createRoleSaveicon.click();

		String expMessage = "Role saved Successfully";

		String actMessage = checkValidationMessage(expMessage);

		// getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleCloseIcon));
		// createRoleCloseIcon.click();

		System.out.println("actAssignedProfilesList : " + actAssignedProfilesList);
		System.out.println("expAssignedProfilesList : " + expAssignedProfilesList);

		System.out.println("actFinancialMenusList : " + actFinancialMenusList);
		System.out.println("expFinancialMenusList : " + expFinancialMenusList);

		System.out.println("actTransactionsMenusList : " + actTransactionsMenusList);
		System.out.println("expTransactionsMenusList : " + expTransactionsMenusList);

		System.out.println("actPurchasesMenusList : " + actPurchasesMenusList);
		System.out.println("expPurchasesMenusList : " + expPurchasesMenusList);

		System.out.println("actInventoryMenusList : " + actInventoryMenusList);
		System.out.println("expInventoryMenusList : " + expInventoryMenusList);

		System.out.println("actInventoryTransactionsMenusList : " + actInventoryTransactionsMenusList);
		System.out.println("expInventoryTransactionsMenusList : " + expInventoryTransactionsMenusList);

		System.out.println("Error Message : " + actMessage + "  value expected  " + expMessage);

		if (actAssignedProfilesList.equalsIgnoreCase(expAssignedProfilesList)
				&& actFinancialMenusList.equalsIgnoreCase(expFinancialMenusList)
				&& actTransactionsMenusList.equalsIgnoreCase(expTransactionsMenusList)
				&& actPurchasesMenusList.equalsIgnoreCase(expPurchasesMenusList)
				&& actInventoryMenusList.equalsIgnoreCase(expInventoryMenusList)
				&& actInventoryTransactionsMenusList.equalsIgnoreCase(expInventoryTransactionsMenusList)
				&& actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {
			return false;
		}
	}

	public boolean checkEditAndUpdateSavedRole() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	
	{
		
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		createRoleRoleNameCombo.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);;
		createRoleRoleNameCombo.sendKeys("Testing");
		Thread.sleep(4000);
		createRoleRoleNameCombo.sendKeys(Keys.TAB);
		Thread.sleep(8000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(assignedProfilesTab));
		assignedProfilesTab.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(avaliableAllProfileSt));
		avaliableAllProfileSt.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleMOveFiledsFromLeftSideToRightSide));
		createRoleMOveFiledsFromLeftSideToRightSide.click();

		// Additions Tab
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

		int actPurchasesMenusCount = additionTabFinancialsTransactionsPurchasesMenusList.size();
		int expPurchasesMenusCount = 2;

		System.out.println("additionTabFinancialsTransactionsPurchasesMenusList : " + actPurchasesMenusCount);
		System.out.println("additionTabFinancialsTransactionsPurchasesMenusList : " + expPurchasesMenusCount);

		boolean actCRAddAlwaysSuspendOnSavingChkBox = CRAddAlwaysSuspendOnSavingChkBox.isDisplayed();
		boolean actCRAddHideMenuChkBox = CRAddHideMenuChkBox.isDisplayed();

		boolean expCRAddAlwaysSuspendOnSavingChkBox = true;
		boolean expCRAddHideMenuChkBox = true;

		System.out.println("CRAddAlwaysSuspendOnSavingChkBox           : " + actCRAddAlwaysSuspendOnSavingChkBox
				+ "  value expected  " + expCRAddAlwaysSuspendOnSavingChkBox);
		System.out.println("CRAddHideMenuChkBox                        : " + actCRAddHideMenuChkBox
				+ "  value expected  " + expCRAddHideMenuChkBox);

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

		Select CantPrintAfterValue = new Select(cantPrintAfterValueDropdown);
		CantPrintAfterValue.selectByVisibleText("Days");

		CantRePrintAfterValueTxt.click();
		CantRePrintAfterValueTxt.clear();
		CantRePrintAfterValueTxt.sendKeys("5");

		Select CantRePrintAfterValue = new Select(CantRePrintAfterDropdown);
		CantRePrintAfterValue.selectByVisibleText("Hours");

		//doNotAllowMasterCustomization.click();
		
		if(CheckdoNotAllowMasterCustomization.isSelected()==false)
		{
			doNotAllowMasterCustomization.click();
		}

		// AI Rights
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AIRightsTab));
		AIRightsTab.click();

		Thread.sleep(2000);

		//AIRightsEnableVUIChkBox.click();

		//AIRightsEnableAIChkBox.click();
		
		if(CheckAIRightsEnableVUIChkBox.isSelected()==false)
		{
			AIRightsEnableVUIChkBox.click();
		}
		
		if(CheckAIRightsEnableAIChkBox.isSelected()==false)
		{
			AIRightsEnableAIChkBox.click();
		}
		
		if(CheckdoNotApplyTheTagResTrictionToTheResultsChkBox.isSelected()==false)
		{
			doNotApplyTheTagResTrictionToTheResultsChkBox.click();
		}

		//doNotApplyTheTagResTrictionToTheResultsChkBox.click();

		NoOfVoiceCommandsPermittedTxt.click();
		NoOfVoiceCommandsPermittedTxt.clear();
		NoOfVoiceCommandsPermittedTxt.sendKeys("1");

		NoOfTextCommandsPermittedTxt.click();
		NoOfTextCommandsPermittedTxt.clear();
		NoOfTextCommandsPermittedTxt.sendKeys("1");

		AIRightsSalesOrderChkBox.click();
		AIRightsPurchasesOrdersChkBox.click();*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
		createRoleSaveicon.click();

		String expMessage = "Role updated Successfully";
		String actMessage = checkValidationMessage(expMessage);

		Thread.sleep(2000);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			return true;
		} else {
			return false;
		}
	}

	public boolean checkCompareAndDeleteUpdatedRole()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		createRoleRoleNameCombo.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		createRoleRoleNameCombo.sendKeys("RoleToDelete");
		Thread.sleep(2000);
		createRoleRoleNameCombo.sendKeys(Keys.TAB);
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRolePasswordPolicyDropdown));
		createRolePasswordPolicyDropdown.click();
		createRolePasswordPolicyDropdown.sendKeys("Simple");
		createRolePasswordPolicyDropdown.sendKeys(Keys.TAB);

		Thread.sleep(10000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(avaliableAllProfileSt));
		avaliableAllProfileSt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleMOveFiledsFromLeftSideToRightSide));
		createRoleMOveFiledsFromLeftSideToRightSide.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
		createRoleSaveicon.click();

		checkValidationMessage("Role saved Successfully");

		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		createRoleRoleNameCombo.sendKeys(Keys.END);
		createRoleRoleNameCombo.sendKeys(Keys.SHIFT, Keys.HOME);
		createRoleRoleNameCombo.sendKeys("RoleToDelete");
		Thread.sleep(1000);
		createRoleRoleNameCombo.sendKeys(Keys.TAB);

		Thread.sleep(20000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CreateRoleDeleteButtonImage));
		getAction().moveToElement(CreateRoleDeleteButtonImage).click().build().perform();
		Thread.sleep(2000);
		/*
		 * getFluentWebDriverWait().until(ExpectedConditions.
		 * elementToBeClickable(CreateRoleDeleteButtonImage));
		 * CreateRoleDeleteButtonImage.click();
		 */
		getWaitForAlert();

		String actAlertText = getAlert().getText();
		String expAlertText = "are you sure that you want to delete Role";

		getAlert().accept();
		Thread.sleep(2000);

		String expMessage = "Role Successfully deleted..";
		String actMessage = checkValidationMessage(expMessage);

		System.out.println("AlertText  : " + actAlertText + "  value expected  " + expAlertText);
		System.out.println("Error Message  : " + actMessage + "  value expected  " + expMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleCloseIcon));
		createRoleCloseIcon.click();

		if (actAlertText.equalsIgnoreCase(expAlertText) && actMessage.equalsIgnoreCase(expMessage)) {
			return true;
		} else {
			return false;
		}
	}

	public static boolean checkCreateUserWithAllMandatoryFileds()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(3000);
	/*	getDriver().navigate().refresh();
		Thread.sleep(2000);*/
		
	//	boolean flag=checkLogoutAndLoginWithSU();
	//	Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		//ClickUsingJs(createUserIcon);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
		createUserIcon.click();
		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
		createUserLoginNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 29, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		createUserERPRoleDropDown.sendKeys(excelReader.getCellData(xlSheetName, 30, 5));

		int count = createUserERPRolesDropdownOptions.size() - 1;
		System.out.println("Count " + count);

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
		s.selectByVisibleText("In which country were you born?");//excelReader.getCellData(xlSheetName, 35, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));
		createUserSecurityAnswerTxt.sendKeys(excelReader.getCellData(xlSheetName, 36, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
		int counnt = createUserLanguageDropdownList.size() - 1;
		String actLanguageDropdownCount = Integer.toString(counnt);
		String expLanguageDropdownCount = excelReader.getCellData(xlSheetName, 29, 6);

		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actLanguageDropdownCount);

		Set LanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < counnt; i++) {
			String data = createUserLanguageDropdownList.get(i).getText();
			LanguageDropdownList.add(data);
		}

		String actLanguageDropdownList = LanguageDropdownList.toString();
		String expLanguageDropdownList = excelReader.getCellData(xlSheetName, 30, 6);

		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actLanguageDropdownList);

		System.out.println("LanguageDropdownCount  : " + actLanguageDropdownCount + "  value expected  "
				+ expLanguageDropdownCount);
		System.out.println(
				"LanguageDropdownList   : " + actLanguageDropdownList + "  value expected  " + expLanguageDropdownList);

		Thread.sleep(2000);	
		Select language = new Select(getLanguage);
		language.selectByIndex(1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
		int cont = createUserAltLanguageDropdownList.size() - 1;
		String actAltLanguageDropdownCount = Integer.toString(count);
		String expAltLanguageDropdownCount = excelReader.getCellData(xlSheetName, 31, 6);

		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actAltLanguageDropdownCount);

		Thread.sleep(2000);	
		
		Set AltLanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < cont; i++) {
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

		System.err.println("Before Save");
	
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		System.err.println("After Save");
		String expMessage = excelReader.getCellData(xlSheetName, 33, 6);
		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			excelReader.setCellData(xlfile, xlSheetName, 29, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 29, 8, resFail);
			return false;
		}
	}

	public static boolean checkLogoutAndLoginWithCreatedUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 40, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 41, 5));
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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

		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();

		Thread.sleep(2000);
		/*lp.reLogin(excelReader.getCellData(xlSheetName, 40, 5), excelReader.getCellData(xlSheetName, 41, 5),
				"Automation Company ‎");

		Thread.sleep(5000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDemoDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 40, 6);

		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actUserInfo);

		Thread.sleep(2000);
		;

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 41, 6);

		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actLoginCompanyName);
		//companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
		System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
				+ expLoginCompanyName + ".");

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {
			System.out.println("***Test Pass: Login Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 40, 8, resPass);
			return true;
		} else {

			System.out.println("***Test Fail: Login Not Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 40, 8, resFail);
			return false;
		}
	}

	public static boolean checkUpdatingTheCreatedUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 44, 5))) {
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

		/*
		 * for (int i = 0; i < 5; i++) { if
		 * (validationConfirmationMessage.getText().isEmpty()==false) {
		 * checkValidationMessage(""); } }
		 * 
		 * Thread.sleep(2000);
		 */

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		getAction().moveToElement(SaveBtn).click().build().perform();

		String expMessage = excelReader.getCellData(xlSheetName, 44, 6);
		String actMessage = checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 44, 8, resPass);
			return true;
		} else {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 44, 8, resFail);
			return false;
		}
	}

	public boolean checkDeletingUpdatedUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(8000);

		int actUsersCount = createUserSavedUsersList.size();

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();
			System.out.println(actUser);

			if (actUser.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 48, 5))) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(2000);
				break;
			}
		}

		Thread.sleep(8000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileDeleteIcon));
		createProfileDeleteIcon.click();
		Thread.sleep(2000);
				//getWaitForAlert();
			if(getIsAlertPresent())
			{
					getAlert().accept();
			}

		String expMessage = excelReader.getCellData(xlSheetName, 48, 6);
		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 48, 8, resPass);

			return true;
		} else {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 48, 8, resFail);
			return false;
		}
	}

	public boolean checInputMandatoryFieldsInCreateUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		
		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(10000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
		createUserIcon.click();
		
		Thread.sleep(6000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
		createUserLoginNameTxt.click();
		createUserLoginNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 49, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		createUserERPRoleDropDown.click();
		createUserERPRoleDropDown.sendKeys(excelReader.getCellData(xlSheetName, 50, 5));

		int count = createUserERPRolesDropdownOptions.size() - 1;
		System.out.println("Count " + count);

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
		s.selectByVisibleText("In which country were you born?");//excelReader.getCellData(xlSheetName, 55, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));
		createUserSecurityAnswerTxt.sendKeys(excelReader.getCellData(xlSheetName, 56, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
		int counnt = createUserLanguageDropdownList.size();
		String actLanguageDropdownCount = Integer.toString(counnt);
		String expLanguageDropdownCount = excelReader.getCellData(xlSheetName, 49, 6);

		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actLanguageDropdownCount);

		Set LanguageDropdownList = new HashSet<String>();

		for (int i = 1; i < counnt; i++) {
			String data = createUserLanguageDropdownList.get(i).getText();
			LanguageDropdownList.add(data);
		}

		String actLanguageDropdownList = LanguageDropdownList.toString();
		String expLanguageDropdownList = excelReader.getCellData(xlSheetName, 50, 6);

		excelReader.setCellData(xlfile, xlSheetName, 50, 7, actLanguageDropdownList);

		System.out.println("LanguageDropdownCount  : " + actLanguageDropdownCount + "  value expected  "
				+ expLanguageDropdownCount);
		System.out.println(
				"LanguageDropdownList   : " + actLanguageDropdownList + "  value expected  " + expLanguageDropdownList);

		Select language = new Select(getLanguage);
		language.selectByVisibleText(excelReader.getCellData(xlSheetName, 57, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
		int cont = createUserAltLanguageDropdownList.size();
		String actAltLanguageDropdownCount = Integer.toString(count);
		String expAltLanguageDropdownCount = excelReader.getCellData(xlSheetName, 51, 6);

		excelReader.setCellData(xlfile, xlSheetName, 51, 7, actAltLanguageDropdownCount);

		Set AltLanguageDropdownList = new HashSet<String>();

		for (int i = 1; i < cont; i++) {
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

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = excelReader.getCellData(xlSheetName, 53, 6);
		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			excelReader.setCellData(xlfile, xlSheetName, 49, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 49, 8, resFail);
			return false;
		}
	}

	public static boolean checkCreateUserWithRestrictionsForAccount()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
		createUserIcon.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
		createUserLoginNameTxt.click();
		createUserLoginNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 62, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		createUserERPRoleDropDown.click();
		createUserERPRoleDropDown.sendKeys(excelReader.getCellData(xlSheetName, 63, 5));

		int count = createUserERPRolesDropdownOptions.size() - 1;
		System.out.println("Count " + count);
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 64, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.sendKeys(excelReader.getCellData(xlSheetName, 65, 5));
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		createUserNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 66, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
		createUserLoginAbbrivationTxt.sendKeys(excelReader.getCellData(xlSheetName, 67, 5));
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityQuestionDropDown));
		createUserSecurityQuestionDropDown.click();
		Thread.sleep(2000);
		Select s = new Select(createUserSecurityQuestionDropDown);
		s.selectByVisibleText("In which country were you born?");//excelReader.getCellData(xlSheetName, 68, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));
		createUserSecurityAnswerTxt.click();
		createUserSecurityAnswerTxt.sendKeys(excelReader.getCellData(xlSheetName, 69, 5));
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
		int counnt = createUserLanguageDropdownList.size() ;
		String actLanguageDropdownCount = Integer.toString(counnt);
		String expLanguageDropdownCount = excelReader.getCellData(xlSheetName, 62, 6);

		excelReader.setCellData(xlfile, xlSheetName, 62, 7, actLanguageDropdownCount);

		Set LanguageDropdownList = new HashSet<String>();

		for (int i = 1; i < counnt; i++) {
			String data = createUserLanguageDropdownList.get(i).getText();
			LanguageDropdownList.add(data);
		}
		Thread.sleep(2000);
		String actLanguageDropdownList = LanguageDropdownList.toString();
		String expLanguageDropdownList = excelReader.getCellData(xlSheetName, 63, 6);

		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actLanguageDropdownList);

		System.out.println("LanguageDropdownCount  : " + actLanguageDropdownCount + "  value expected  "
				+ expLanguageDropdownCount);
		System.out.println(
				"LanguageDropdownList   : " + actLanguageDropdownList + "  value expected  " + expLanguageDropdownList);
		Thread.sleep(2000);
		Select language = new Select(getLanguage);
		language.selectByVisibleText(excelReader.getCellData(xlSheetName, 70, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
		int cont = createUserAltLanguageDropdownList.size() ;
		String actAltLanguageDropdownCount = Integer.toString(count);
		String expAltLanguageDropdownCount = excelReader.getCellData(xlSheetName, 64, 6);

		excelReader.setCellData(xlfile, xlSheetName, 64, 7, actAltLanguageDropdownCount);

		Set AltLanguageDropdownList = new HashSet<String>();

		for (int i = 1; i < cont; i++) {
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

		int actRestrictionForEntryMastersListCount = createRoleRestrictionForEntryTabMastersList.size();

		for (int i = 0; i < actRestrictionForEntryMastersListCount; i++) {
			String masters = createRoleRestrictionForEntryTabMastersList.get(i).getText();

			if (masters.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 75, 5))) {
				createRoleRestrictionForEntryTabMastersList.get(i).click();
				System.err.println("**Account is Selected**");
			}
		}

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(createRoleRestrictionForEntryTabExclChkBox));
		createRoleRestrictionForEntryTabExclChkBox.click();

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableSelectRow1Col1));
		restrictionForEntryTabTableSelectRow1Col1.click();

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableEnterMasterTxt));
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(excelReader.getCellData(xlSheetName, 76, 5));
		Thread.sleep(4000);
		restrictionForEntryTabTableEnterMasterTxt.sendKeys(Keys.TAB);
		Thread.sleep(6000);
		
		restrictionForEntryTabRow1EntryChkBox.click();
		restrictionForEntryTabRow1ReportChkBox.click();
		restrictionForEntryTabRow1ViewChkBox.click();

		getFluentWebDriverWait()
		
				.until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableSelectRow2Col1));
		restrictionForEntryTabTableSelectRow2Col1.click();

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableEnterMasterTxt));

		restrictionForEntryTabTableEnterMasterTxt.sendKeys("Customer C"/*excelReader.getCellData(xlSheetName, 77, 5)*/);
		Thread.sleep(4000);

		restrictionForEntryTabTableEnterMasterTxt.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		restrictionForEntryTabRow2EntryChkBox.click();
		restrictionForEntryTabRow2ReportChkBox.click();
		restrictionForEntryTabRow2ViewChkBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);
		
		String expMessage = excelReader.getCellData(xlSheetName, 66, 6);
		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 66, 7, actMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			excelReader.setCellData(xlfile, xlSheetName, 62, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 62, 8, resFail);
			return false;
		}
	}

	public boolean checkLogoutAndLoginWithNewlySavedUser()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp = new LoginPage(getDriver());

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


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 79, 5));
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);*/
		
		/*lp.reLogin(excelReader.getCellData(xlSheetName, 78, 5), excelReader.getCellData(xlSheetName, 79, 5),
				"Automation Company");

		Thread.sleep(5000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameTestDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 78, 6);

		excelReader.setCellData(xlfile, xlSheetName, 78, 7, actUserInfo);

		Thread.sleep(2000);
		;

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 79, 6);
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actLoginCompanyName);
		//companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
		System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
				+ expLoginCompanyName + ".");

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {
			excelReader.setCellData(xlfile, xlSheetName, 78, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 78, 8, resFail);
			return false;
		}
	}

	public static boolean checkLogoutAndLoginWithSU()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(3000);

		getDriver().navigate().refresh();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 42, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 43, 5));

		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		

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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();

		/*
		 * Thread.sleep(2000); lp.reLogin(excelReader.getCellData(xlSheetName,
		 * 42, 5),excelReader.getCellData(xlSheetName, 43,
		 * 5),"Automation Company");
		 */
		Thread.sleep(10000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 42, 6);

		excelReader.setCellData(xlfile, xlSheetName, 42, 7, actUserInfo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));

		//companyLogo.click();

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 43, 6);
		//companyLogo.click();

		System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
		System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
				+ expLoginCompanyName + ".");

		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actLoginCompanyName);

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {
			System.out.println("***Test Pass: Login Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resPass);
			return true;
		} else {

			System.out.println("***Test Fail: Login Not Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 42, 8, resFail);
			return false;
		}
	}

	public boolean checkEditingAndUpdatingSavedUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 80, 5))) {
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

		for (int i = 0; i < 5; i++) {
			if (validationConfirmationMessage.getText().isEmpty() == false) {
				checkValidationMessage("");
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = excelReader.getCellData(xlSheetName, 80, 6);
		String actMessage = checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 80, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 80, 8, resPass);
			return true;
		} else {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			excelReader.setCellData(xlfile, xlSheetName, 80, 8, resFail);
			return false;
		}
	}

	public boolean checkLogoutAndLoginWithUpdatedUserWithPreviousPassword()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp = new LoginPage(getDriver());
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 84, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 85, 5));

		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		
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


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();
		
		
		/*lp.reLogin(excelReader.getCellData(xlSheetName, 84, 5), excelReader.getCellData(xlSheetName, 85, 5),
				"Automation Company");
*/
		String actMessage = invalidPasswordMsgs.getText();
		String expMessage = "Invalid password";//excelReader.getCellData(xlSheetName, 84, 6);

		//excelReader.setCellData(xlfile, xlSheetName, 84, 7, actMessage);

		System.out.println("mandatoryMsgs : " + actMessage + "  value expected  " + expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			excelReader.setCellData(xlfile, xlSheetName, 84, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 84, 8, resFail);
			return false;
		}
	}

	public boolean checkLoginWithUpdatedPassword()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		clearValueFromTextBox(password);
		password.sendKeys(excelReader.getCellData(xlSheetName, 86, 5));

		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		
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


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();

		Thread.sleep(2000);
		/*lp.reLogin(excelReader.getCellData(xlSheetName, 84, 5), excelReader.getCellData(xlSheetName, 86, 5),
				"Automation Company");

		Thread.sleep(5000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameTestDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 86, 6);

		excelReader.setCellData(xlfile, xlSheetName, 86, 7, actUserInfo);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));

		//companyLogo.click();

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 87, 6);

		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actLoginCompanyName);

		//companyLogo.click();

		System.out.println("UserInfo               : " + actUserInfo + "  value expected  " + expUserInfo);
		System.out.println(
				"LoginCompanyName       : " + actLoginCompanyName + "  value expected  " + expLoginCompanyName);

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {
			System.out.println("***Test Pass: Login Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 86, 8, resPass);

			return true;
		} else {

			System.out.println("***Test Fail: Login Not Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 86, 8, resFail);

			return false;
		}
	}

	public static boolean checkPurchasesVoucherForRestrictedAccountsInUserThroughRole()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getDriver().navigate().refresh();
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesExpandBtn));
		purchasesExpandBtn.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersBtn));
		purchaseVouchersBtn.click();
		Thread.sleep(20000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		////checkValidationMessage("Screen Opened");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountComboExpand));
		vendorAccountComboExpand.click();

		int vendorCount = vendorAccountList.size();
		boolean vendor = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < vendorCount; i++) {
			data = vendorAccountList.get(i).getText();
			accountslist.add(data);
			if (data.equalsIgnoreCase("Vendor B"/*excelReader.getCellData(xlSheetName, 88, 5)*/)) {
				vendor = false;
				break;
			}
		}

		System.out.println("Vendor Accounts: " + accountslist);

		String actual = Boolean.toString(vendor);
		String expected = excelReader.getCellData(xlSheetName, 88, 6);

		excelReader.setCellData(xlfile, xlSheetName, 88, 7, actual);
		
		System.out.println("Act Vendor B  Restriction : "+actual);
		System.out.println("Exp Vendor B  Restriction : "+expected);

		if (actual.equalsIgnoreCase(expected)) {
			excelReader.setCellData(xlfile, xlSheetName, 88, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 88, 8, resFail);
			return false;
		}
	}

	public static boolean checkSalesInvoicesVATForRestrictedAccountsInUserThroughRole()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Thread.sleep(2000);

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
		
		
		
		
		if(getIsAlertPresent())
		{
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
		}
		
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		//checkValidationMessage("Screen Opened");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountComboExpand));
		customerAccountComboExpand.click();

		int customerCount = customerAccountList.size();
		boolean customer = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < customerCount; i++) {
			data = customerAccountList.get(i).getText();
			accountslist.add(data);
			if (data.equalsIgnoreCase("Customer C"/*excelReader.getCellData(xlSheetName, 89, 5)*/)) {
				customer = false;
				break;
			}
		}

		System.out.println("Customer  Accounts: " + accountslist);

		String actual = Boolean.toString(customer);
		String expected = excelReader.getCellData(xlSheetName, 89, 6);

		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actual);

		System.out.println("Act Customer C  Restriction : "+actual);
		System.out.println("Exp Customer C  Restriction : "+expected);
		
		if (actual.equalsIgnoreCase(expected)) {
			excelReader.setCellData(xlfile, xlSheetName, 89, 8, resPass);
			return true;

		} else {
			excelReader.setCellData(xlfile, xlSheetName, 89, 8, resFail);
			return false;
		}

	}

	public static boolean checkLedgerForRestrictedAccountsInUserThroughRole()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsReportsMenu));
		financialsReportsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ledger));
		ledger.click();

		Thread.sleep(3000);
		
		if(getIsAlertPresent())
		{
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));
		accountSearchTextArea.sendKeys(excelReader.getCellData(xlSheetName, 90, 5));
		accountSearchTextArea.sendKeys(Keys.ENTER);

		int accountsCount = accountsListInLedger.size();
		boolean vendor = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < accountsCount; i++) {
			data = accountsListInLedger.get(i).getText();
			accountslist.add(data);
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 91, 5))) {
				vendor = false;
				break;
			}
		}

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));
		accountSearchTextArea.clear();
		Thread.sleep(2000);
		accountSearchTextArea.sendKeys("Customer"/*excelReader.getCellData(xlSheetName, 92, 5)*/);
		accountSearchTextArea.sendKeys(Keys.ENTER);

		boolean customer = true;
		String data1 = null;

		for (int i = 0; i < accountsCount; i++) {
			data1 = accountsListInLedger.get(i).getText();
			accountslist.add(data);
			if (data.equalsIgnoreCase("Customer C"/*excelReader.getCellData(xlSheetName, 93, 5)*/)) {
				customer = false;
				break;
			}
		}

		String actual = Boolean.toString(vendor);
		String expected = excelReader.getCellData(xlSheetName, 90, 6);

		excelReader.setCellData(xlfile, xlSheetName, 90, 7, actual);

		String actual1 = Boolean.toString(customer);
		String expected1 = excelReader.getCellData(xlSheetName, 91, 6);

		excelReader.setCellData(xlfile, xlSheetName, 91, 7, actual1);

		if (actual.equalsIgnoreCase(expected) && actual1.equalsIgnoreCase(expected1)) {

			excelReader.setCellData(xlfile, xlSheetName, 90, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 90, 8, resFail);
			return false;
		}
	}

	public static boolean checkLogoutAndLoginWithCreatedRestrictedUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		Thread.sleep(4000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys(excelReader.getCellData(xlSheetName, 94, 5));
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys(excelReader.getCellData(xlSheetName, 95, 5));
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();

		Thread.sleep(2000);
		/*lp.reLogin(excelReader.getCellData(xlSheetName, 94, 5), excelReader.getCellData(xlSheetName, 95, 5),
				"Automation Company");
		Thread.sleep(5000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameUserRestrictionDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 94, 6);

		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actUserInfo);

		Thread.sleep(2000);
		;

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 95, 6);

		excelReader.setCellData(xlfile, xlSheetName, 95, 7, actLoginCompanyName);
		//companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
		System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
				+ expLoginCompanyName + ".");

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {
			System.out.println("***Test Pass: Login Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 94, 8, resPass);
			return true;
		} else {

			System.out.println("***Test Fail: Login Not Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 94, 8, resFail);
			return false;
		}
	}

	public static boolean checkPurchasesVocherForRestrictedAccountThroughUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesExpandBtn));
		purchasesExpandBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersBtn));
		purchaseVouchersBtn.click();
		
		new WebDriverWait(getDriver(), 300).until(ExpectedConditions.visibilityOf(newBtn));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		//checkValidationMessage("Screen Opened");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountComboExpand));
		vendorAccountComboExpand.click();

		int vendorCount = vendorAccountList.size();
		boolean vendor = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < vendorCount; i++) {
			data = vendorAccountList.get(i).getText();
			accountslist.add(data);
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 96, 5))) {
				vendor = false;
				break;
			}
		}

		System.out.println("Vendor Accounts: " + accountslist);

		String actual = Boolean.toString(vendor);
		String expected = excelReader.getCellData(xlSheetName, 96, 6);

		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actual);

		System.out.println("Act Vendor B  Restriction : "+actual);
		System.out.println("Exp Vendor B  Restriction : "+expected);
		
		if (actual.equalsIgnoreCase(expected)) {
			excelReader.setCellData(xlfile, xlSheetName, 96, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 96, 8, resFail);
			return false;
		}
	}

	public static boolean checkSalesInvoiceVATForRestrictedAccountsThroughUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesMenu));
		salesMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoicesVatBtn));
		salesInvoicesVatBtn.click();
		
		if(getIsAlertPresent())
		{
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
		}

		
		ClickUsingJs(newBtn);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();

		//checkValidationMessage("Screen Opened");

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountComboExpand));
		customerAccountComboExpand.click();

		int customerCount = customerAccountList.size();
		boolean customer = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < customerCount; i++) {
			data = customerAccountList.get(i).getText();
			accountslist.add(data);
			if (data.equalsIgnoreCase("Customer C"/*excelReader.getCellData(xlSheetName, 97, 5)*/)) {
				customer = false;
				break;
			}
		}

		System.out.println("Customer  Accounts: " + accountslist);

		String actual = Boolean.toString(customer);
		String expected = excelReader.getCellData(xlSheetName, 97, 6);

		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actual);
		
		System.out.println("Act Customer C  Restriction : "+actual);
		System.out.println("Exp Customer C  Restriction : "+expected);

		if (actual.equalsIgnoreCase(expected)) {
			excelReader.setCellData(xlfile, xlSheetName, 97, 8, resPass);
			return true;

		} else {
			excelReader.setCellData(xlfile, xlSheetName, 97, 8, resFail);
			return false;
		}
	}

	public static boolean checkLedgerForRestrictedAccountThroughUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsReportsMenu));
		financialsReportsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ledger));
		ledger.click();

		Thread.sleep(3000);
		
		if(getIsAlertPresent())
		{
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountSearchTextArea));
		accountSearchTextArea.sendKeys(excelReader.getCellData(xlSheetName, 98, 5));
		accountSearchTextArea.sendKeys(Keys.ENTER);

		int accountsCount = accountsListInLedger.size();
		boolean vendor = true;
		String data = null;

		ArrayList<String> accountslist = new ArrayList<String>();

		for (int i = 0; i < accountsCount; i++) {
			data = accountsListInLedger.get(i).getText();
			accountslist.add(data);
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 99, 5))) {
				vendor = false;
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

		for (int i = 0; i < accountsCount; i++) {
			data1 = accountsListInLedger.get(i).getText();
			accountslist.add(data);
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 101, 5))) {
				customer = false;
				break;
			}
		}

		String actual = Boolean.toString(vendor);
		String expected = excelReader.getCellData(xlSheetName, 98, 6);

		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actual);

		String actual1 = Boolean.toString(customer);
		String expected1 = excelReader.getCellData(xlSheetName, 99, 6);

		excelReader.setCellData(xlfile, xlSheetName, 99, 7, actual1);

		if (actual.equalsIgnoreCase(expected) && actual1.equalsIgnoreCase(expected1)) {

			excelReader.setCellData(xlfile, xlSheetName, 98, 8, resPass);
			return true;
		} else {
			excelReader.setCellData(xlfile, xlSheetName, 98, 8, resFail);
			return false;
		}
	}

	public boolean checkLogoutAndLoginWithSUToUpdatedUserDeleteUser()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		
		Thread.sleep(4000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		LoginPage lp = new LoginPage(getDriver());

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


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();

		Thread.sleep(2000);
		/*lp.reLogin(excelReader.getCellData(xlSheetName, 102, 5), excelReader.getCellData(xlSheetName, 103, 5),
				"Automation Company");
		Thread.sleep(5000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = excelReader.getCellData(xlSheetName, 102, 6);

		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actUserInfo);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));

		companyLogo.click();*/

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = excelReader.getCellData(xlSheetName, 103, 6);

		excelReader.setCellData(xlfile, xlSheetName, 103, 7, actLoginCompanyName);

		//companyLogo.click();

		System.out.println("UserInfo               : " + actUserInfo + "  value expected  " + expUserInfo);
		System.out.println(
				"LoginCompanyName       : " + actLoginCompanyName + "  value expected  " + expLoginCompanyName);

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {
			System.out.println("***Test Pass: Login Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 102, 8, resPass);

			return true;
		} else {
			System.out.println("***Test Fail: Login Not Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 102, 8, resFail);

			return false;
		}
	}

	///////////////////////////////////////////////
	public boolean checkChangePasswordatUserLevelwithNewPassword()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
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

		String expMessage = "Password Saved Successfully";
		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(actMessage)) {
			return true;
		} else {

			return false;

		}

	}

	public boolean checkLoginangLogoutSUwithChangePassword() throws InterruptedException {

		LoginPage lp = new LoginPage(getDriver());

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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));

		signIn.click();

		Thread.sleep(2000);
		lp.reLogin("su", "1234", "Automation Company");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = "";
		if (actUserInfo.equalsIgnoreCase(expUserInfo)) {
			System.out.println("Test Pass with Change Password");
			return true;

		} else

		{
			return false;
		}

	}

	////////
	public boolean checkChangePasswordatUserLevelwithOldPassword()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
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

		String expMessage = "Password Saved Successfully";
		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(actMessage)) {
			return true;
		} else {

			return false;

		}

	}

	public boolean checkLoginangLogoutSUwithChangePasswordUpdate() throws InterruptedException {

		LoginPage lp = new LoginPage(getDriver());

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
		lp.reLogin("su", "su", "Automation Company");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = "SU";
		if (actUserInfo.equalsIgnoreCase(expUserInfo)) {
			System.out.println("Test Pass with Change Password");
			return true;

		} else

		{
			return false;
		}

	}

	public boolean checkComplexityofPasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordPolicy_SelectComplexity));
		Select s = new Select(PasswordPolicy_SelectComplexity);
		s.selectByVisibleText("Alpha-numeric");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkChangePasswordwhileChangePasswordPolicyComplexityAsAplhaNumericwithRestriction()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase("NewUser")) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(4000);

				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys("focus123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys("focus123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Alphabets and numerics but atleast 1 lower case and 1 upper case alphabet must be given in the password.";
		String actMessage = checkValidationMessage(expMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys("FOCUS123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys("FOCUS123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);
		String expMessage1 = "Alphabets and numerics but atleast 1 lower case and 1 upper case alphabet must be given in the password.";
		String actMessage1 = checkValidationMessage(expMessage1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys("Focus123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys("Focus123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);
		String expMessage2 = "User updated successfully.";
		String actMessage2 = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1)
				&& actMessage2.equalsIgnoreCase(expMessage2)) {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();

			return true;
		} else {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();

			return false;
		}
	}

	public boolean checkCreateaNewUserforPasswordasAlphaNumeric()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
		Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
		createUserIcon.click();
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
		createUserLoginNameTxt.sendKeys("NewUser");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		createUserERPRoleDropDown.sendKeys("Test");

		int count = createUserERPRolesDropdownOptions.size() - 1;
		System.out.println("Count " + count);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordPolicyDropdown));
		createUserPasswordPolicyDropdown.sendKeys("Test");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.sendKeys("abcdef");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.sendKeys("abcdef");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		createUserNameTxt.sendKeys("NewUser");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
		createUserLoginAbbrivationTxt.sendKeys("NewUser");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityQuestionDropDown));
		Select s = new Select(createUserSecurityQuestionDropDown);
		s.selectByVisibleText("In which country were you born?");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));
		createUserSecurityAnswerTxt.sendKeys("India");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
		int counnt = createUserLanguageDropdownList.size() - 1;
		String actLanguageDropdownCount = Integer.toString(counnt);
		String expLanguageDropdownCount = "2";

		Set LanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < counnt; i++) {
			String data = createUserLanguageDropdownList.get(i).getText();
			LanguageDropdownList.add(data);
		}

		String actLanguageDropdownList = LanguageDropdownList.toString();
		String expLanguageDropdownList = "[English, Please select a language type] ";

		System.out.println("LanguageDropdownCount  : " + actLanguageDropdownCount + "  value expected  "
				+ expLanguageDropdownCount);
		System.out.println(
				"LanguageDropdownList   : " + actLanguageDropdownList + "  value expected  " + expLanguageDropdownList);

		Select language = new Select(getLanguage);
		language.selectByIndex(1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
		int cont = createUserAltLanguageDropdownList.size() - 1;
		String actAltLanguageDropdownCount = Integer.toString(count);
		String expAltLanguageDropdownCount = "";

		Set AltLanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < cont; i++) {
			String data = createUserAltLanguageDropdownList.get(i).getText();
			AltLanguageDropdownList.add(data);
		}

		String actAltLanguageDropdownList = AltLanguageDropdownList.toString();
		String expAltLanguageDropdownList = "[English, Please select a language type] ";

		Select altlanguage = new Select(getAltLanguage);
		altlanguage.selectByIndex(2);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userCreationTimeZoneDrpdwn));
		Select timeZone = new Select(userCreationTimeZoneDrpdwn);
		timeZone.selectByVisibleText("(GMT+05:30) Chennai, Kolkatta, Mumbai, New Delhi");

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
		createUserEmailTxt.sendKeys("emailvalidationtwo@gmail.com");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		System.out.println("Password given as Only Alphabets");

		String expMessage = "Alphabets and numerics must be given in the password.";
		String actMessage = checkValidationMessage(expMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserInfoTab));
		createUserInfoTab.click();
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.click();
		createUserPasswordTxt.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		createUserPasswordTxt.sendKeys("123456");
		Thread.sleep(1500);
		createUserPasswordTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.sendKeys("123456");

		System.out.println("Password given as Only Numerics");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage1 = "Alphabets and numerics must be given in the password.";
		String actMessage1 = checkValidationMessage(expMessage1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.click();
		createUserPasswordTxt.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		createUserPasswordTxt.sendKeys("focus123");
		Thread.sleep(1500);
		createUserPasswordTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.sendKeys("focus123");

		System.out.println("Password given as Only Alpha-Numerics");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage2 = "new user created successfully";
		String actMessage2 = checkValidationMessage(expMessage2);

		if (actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1)
				&& actMessage2.equalsIgnoreCase(expMessage2)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkComplexityofPasswordPolicyasAlphaNumericwithRestriction()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		getDriver().navigate().refresh();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordPolicy_SelectComplexity));
		Select s = new Select(PasswordPolicy_SelectComplexity);
		s.selectByVisibleText("Alpha-numeric*");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkComplexityofPasswordPolicyasAlphaNumericandSpecialCharcter()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		getDriver().navigate().refresh();
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordPolicy_SelectComplexity));
		Select s = new Select(PasswordPolicy_SelectComplexity);
		s.selectByVisibleText("Alpha-numeric & special characters");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkChangePasswordwhileChangePasswordPolicyComplexityAsAplhaNumericandSpecialCharacter()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase("NewUser")) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(4000);

				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys("focus123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys("focus123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		Thread.sleep(2000);
		String expMessage = "Password must contain alphabets, numerics and special characters.";
		String actMessage = checkValidationMessage(expMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys("FOCUS123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys("FOCUS123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);
		String expMessage1 = "Password must contain alphabets, numerics and special characters.";
		String actMessage1 = checkValidationMessage(expMessage1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys("focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys("focus@123");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage2 = "User updated successfully.";
		String actMessage2 = checkValidationMessage(expMessage2);

		if (actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1)
				&& actMessage2.equalsIgnoreCase(expMessage2)) {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();

			return true;
		} else {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();

			return false;
		}
	}

	public boolean checkComplexityofPasswordPolicyasAlphaNumericwandSpecialCharacterithRestriction()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PasswordPolicy_SelectComplexity));
		Select s = new Select(PasswordPolicy_SelectComplexity);
		s.selectByVisibleText("Alpha-numeric & special characters*");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkChangePasswordwhileChangePasswordPolicyComplexityAsAplhaNumericandSpecialCharacterWithRestriction()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase("NewUser")) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(4000);

				break;
			}
		}

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys("focus123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys("focus123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Alphabets, numerics and special characters but atleast 1 lower case and 1 upper case alphabet must be given in the password.";
		String actMessage = checkValidationMessage(expMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys("FOCUS123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys("FOCUS123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);
		String expMessage1 = "Alphabets, numerics and special characters but atleast 1 lower case and 1 upper case alphabet must be given in the password.";
		String actMessage1 = checkValidationMessage(expMessage1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys("focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys("focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage2 = "Alphabets, numerics and special characters but atleast 1 lower case and 1 upper case alphabet must be given in the password.";
		String actMessage2 = checkValidationMessage(expMessage2);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.clear();
		createUserPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.clear();
		createUserConfirmPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);
		String expMessage3 = "User updated successfully.";
		String actMessage3 = checkValidationMessage(expMessage3);

		if (actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1)
				&& actMessage2.equalsIgnoreCase(expMessage2) && actMessage3.equalsIgnoreCase(expMessage3)) {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();

			return true;
		} else {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();

			return false;
		}
	}

	public boolean checkDoNotAllowPreviousPasswordinPasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(doNotAllowPrevious));
		doNotAllowPrevious.sendKeys("3");
		Thread.sleep(2000);

		System.out.println(
				"Giving Do Not Allow Policy Previous as 3i,e.,after two attempts of change password it will again allow the same password while creating");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkSavingNewUserafterDoNotAllowPreviousPasswordinPasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

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
		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
		createUserLoginNameTxt.sendKeys("PwdUser");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		createUserERPRoleDropDown.sendKeys("Test");

		int count = createUserERPRolesDropdownOptions.size() - 1;
		System.out.println("Count " + count);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordPolicyDropdown));
		createUserPasswordPolicyDropdown.sendKeys("Test");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		createUserNameTxt.sendKeys("PwdUser");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
		createUserLoginAbbrivationTxt.sendKeys("PwdUser");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityQuestionDropDown));
		Select s = new Select(createUserSecurityQuestionDropDown);
		s.selectByVisibleText("In which country were you born?");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));
		createUserSecurityAnswerTxt.sendKeys("India");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
		int counnt = createUserLanguageDropdownList.size() - 1;
		String actLanguageDropdownCount = Integer.toString(counnt);
		String expLanguageDropdownCount = "2";

		Set LanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < counnt; i++) {
			String data = createUserLanguageDropdownList.get(i).getText();
			LanguageDropdownList.add(data);
		}

		String actLanguageDropdownList = LanguageDropdownList.toString();
		String expLanguageDropdownList = "[English, Please select a language type] ";

		System.out.println("LanguageDropdownCount  : " + actLanguageDropdownCount + "  value expected  "
				+ expLanguageDropdownCount);
		System.out.println(
				"LanguageDropdownList   : " + actLanguageDropdownList + "  value expected  " + expLanguageDropdownList);

		Select language = new Select(getLanguage);
		language.selectByIndex(1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
		int cont = createUserAltLanguageDropdownList.size() - 1;
		String actAltLanguageDropdownCount = Integer.toString(count);
		String expAltLanguageDropdownCount = "";

		Set AltLanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < cont; i++) {
			String data = createUserAltLanguageDropdownList.get(i).getText();
			AltLanguageDropdownList.add(data);
		}

		String actAltLanguageDropdownList = AltLanguageDropdownList.toString();
		String expAltLanguageDropdownList = "[English, Please select a language type] ";

		Select altlanguage = new Select(getAltLanguage);
		altlanguage.selectByIndex(2);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userCreationTimeZoneDrpdwn));
		Select timeZone = new Select(userCreationTimeZoneDrpdwn);
		timeZone.selectByVisibleText("(GMT+05:30) Chennai, Kolkatta, Mumbai, New Delhi");

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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = "new user created successfully";
		String actMessage = checkValidationMessage(expMessage);

		System.out.println(actMessage);
		System.out.println(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkLogoutAndLoginWithNewUser()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		Thread.sleep(2000);

		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("Focus@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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


		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNewUserName.getText();
		String expUserInfo = "NewUser";

		Thread.sleep(2000);
		;

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = "Automation Company ";
		//companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
		System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
				+ expLoginCompanyName + ".");

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {

			return true;
		} else {

			return false;
		}
	}

	public boolean checkLogoutAndLoginWithNewUserforIncorrectPwd()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		Thread.sleep(2000);

		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("FocuS@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));

		String actErrMsg = mandatoryMsgs.getText();
		String expErrMsg = "This is your last chance to try the Password";
		System.out.println("Error Message=	" + actErrMsg);

		if (actErrMsg.equalsIgnoreCase(expErrMsg)) {

			return true;
		} else {

			return false;
		}
	}

	public boolean checkLogoutAndLoginWithPwdUser()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		Thread.sleep(2000);

		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("PwdUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("Focus@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userPwdUserName.getText();
		String expUserInfo = "PwdUser";

		Thread.sleep(2000);
		;

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = "Automation Company ";
		//companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
		System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
				+ expLoginCompanyName + ".");

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {

			return true;
		} else {

			return false;
		}
	}

	public boolean checkChangePasswordforNewUsertoCheckDoNotAllowPreviousPassword()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		getDriver().navigate().refresh();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOption));
		changePasswordOption.click();

		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOldPasswordTxt));
		changePasswordOldPasswordTxt.click();
		changePasswordOldPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordNewPasswordTxt));
		changePasswordNewPasswordTxt.sendKeys("focuS@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordConfirmTxt));
		changePasswordConfirmTxt.sendKeys("focuS@123");

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordSaveBtn));
		changePasswordSaveBtn.click();
		Thread.sleep(2000);
		String expMsg = "Password Saved Successfully";
		String actMsg = checkValidationMessage(expMsg);

		Thread.sleep(2000);

		System.err.println("First Attempt to Change the Password as Previous Password");
		Thread.sleep(3000);
		
		getDriver().navigate().refresh();
		Thread.sleep(6000);

		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOption));
		changePasswordOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOldPasswordTxt));
		changePasswordOldPasswordTxt.click();
		changePasswordOldPasswordTxt.sendKeys("focuS@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordNewPasswordTxt));
		changePasswordNewPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordConfirmTxt));
		changePasswordConfirmTxt.sendKeys("Focus@123");

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordSaveBtn));
		changePasswordSaveBtn.click();
		Thread.sleep(2000);
		String expMsg1 = "Do not allow previous password";
		String actMsg1 = checkValidationMessage(expMsg1);

		Thread.sleep(2000);
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordNewPasswordTxt));
		changePasswordNewPasswordTxt.click();
		changePasswordNewPasswordTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		changePasswordNewPasswordTxt.sendKeys("FocuS@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordConfirmTxt));
		changePasswordConfirmTxt.sendKeys(Keys.SHIFT, Keys.HOME);
		changePasswordConfirmTxt.sendKeys("FocuS@123");

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordSaveBtn));
		changePasswordSaveBtn.click();
		Thread.sleep(2000);
		String expMsg2 = "Password Saved Successfully";
		String actMsg2 = checkValidationMessage(expMsg2);

		Thread.sleep(2000);
		
		getDriver().navigate().refresh();
		Thread.sleep(4000);

		System.err.println("Second Attempt to Change the Password as Previous Password");
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOption));
		changePasswordOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOldPasswordTxt));
		changePasswordOldPasswordTxt.click();
		changePasswordOldPasswordTxt.sendKeys("FocuS@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordNewPasswordTxt));
		changePasswordNewPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordConfirmTxt));
		changePasswordConfirmTxt.sendKeys("Focus@123");

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordSaveBtn));
		changePasswordSaveBtn.click();
		Thread.sleep(2000);
		String expMsg3 = "Password Saved Successfully";
		String actMsg3 = checkValidationMessage(expMsg3);

		Thread.sleep(2000);
		/*
		 * getFluentWebDriverWait().until(ExpectedConditions.
		 * elementToBeClickable(changePasswordNewPasswordTxt));
		 * changePasswordNewPasswordTxt.click();
		 * changePasswordNewPasswordTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		 * changePasswordNewPasswordTxt.sendKeys("FocUS@123");
		 * 
		 * getFluentWebDriverWait().until(ExpectedConditions.
		 * elementToBeClickable(changePasswordConfirmTxt));
		 * changePasswordConfirmTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		 * changePasswordConfirmTxt.sendKeys("FocUS@123");
		 * 
		 * Thread.sleep(2000);
		 * getFluentWebDriverWait().until(ExpectedConditions.
		 * elementToBeClickable(changePasswordSaveBtn));
		 * changePasswordSaveBtn.click(); Thread.sleep(2000);
		 * 
		 * String expMsg4="Password Saved Successfully"; String
		 * actMsg4=checkValidationMessage(expMsg4);
		 * 
		 * Thread.sleep(2000);
		 * 
		 * System.err.
		 * println("Third Attempt to Change the Password as Previous Password");
		 * System.err.
		 * println("In the Third attempt it will allow you to change the password with previous password bcoz do not allow previous password as 3"
		 * );
		 * 
		 * 
		 * getFluentWebDriverWait().until(ExpectedConditions.
		 * elementToBeClickable(userNameDisplayLogo));
		 * userNameDisplayLogo.click();
		 * 
		 * Thread.sleep(2000);
		 * 
		 * getFluentWebDriverWait().until(ExpectedConditions.
		 * elementToBeClickable(changePasswordOption));
		 * changePasswordOption.click();
		 * 
		 * Thread.sleep(2000);
		 * 
		 * getFluentWebDriverWait().until(ExpectedConditions.
		 * elementToBeClickable(changePasswordOldPasswordTxt));
		 * changePasswordOldPasswordTxt.click();
		 * changePasswordOldPasswordTxt.sendKeys("FocUS@123");
		 * 
		 * getFluentWebDriverWait().until(ExpectedConditions.
		 * elementToBeClickable(changePasswordNewPasswordTxt));
		 * changePasswordNewPasswordTxt.sendKeys("Focus@123");
		 * 
		 * getFluentWebDriverWait().until(ExpectedConditions.
		 * elementToBeClickable(changePasswordConfirmTxt));
		 * changePasswordConfirmTxt.sendKeys("Focus@123");
		 * 
		 * Thread.sleep(2000);
		 * getFluentWebDriverWait().until(ExpectedConditions.
		 * elementToBeClickable(changePasswordSaveBtn));
		 * changePasswordSaveBtn.click(); Thread.sleep(2000); String
		 * expMsg5="Password Saved Successfully"; String
		 * actMsg5=checkValidationMessage(expMsg5);
		 * 
		 * Thread.sleep(2000);
		 */

		if (actMsg.equalsIgnoreCase(expMsg) && actMsg1.equalsIgnoreCase(expMsg1) && actMsg2.equalsIgnoreCase(expMsg2)
				&& actMsg3.equalsIgnoreCase(
						expMsg3) /*
									 * && actMsg4.equalsIgnoreCase(expMsg4) &&
									 * actMsg5.equalsIgnoreCase(expMsg5)
									 */) {
			return true;
		} else {
			return false;
		}

	}

	public boolean checkLogoutAndLoginWithSUAgain()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("SU");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("su");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = "SU";

		Thread.sleep(2000);
		;

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = "Automation Company ";
		//companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : " + actUserInfo + "  value expected  " + expUserInfo );
		System.out.println("LoginCompanyName       : " + actLoginCompanyName + " value expected  "
				+ expLoginCompanyName);

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {

			return true;
		} else {

			return false;
		}
	}

	public boolean checkInvalidAtetmptsandLockOutPeriodinPasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordInvalidAttempts));
		passwordInvalidAttempts.click();
		passwordInvalidAttempts.sendKeys(Keys.SHIFT, Keys.HOME);
		passwordInvalidAttempts.sendKeys("2");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLockOut));
		passwordLockOut.click();
		passwordLockOut.sendKeys(Keys.SHIFT, Keys.HOME);
		passwordLockOut.sendKeys("1");
		passwordLockOut.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLockOutUnits));
		passwordLockOutUnits.sendKeys("Minute");

		Thread.sleep(2000);
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	
	public boolean checkInvalidAtetmptsandLockOutPeriodinPasswordPolicyAgain() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);
		
		getAction().moveToElement(otpBasedLogin).build().perform();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(otpBasedLoginSelected));
		if (otpBasedLogin.isSelected() == true) {
			otpBasedLoginSelected.click();
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordInvalidAttempts));
		passwordInvalidAttempts.click();
		passwordInvalidAttempts.sendKeys(Keys.SHIFT, Keys.HOME);
		passwordInvalidAttempts.sendKeys("2");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLockOut));
		passwordLockOut.click();
		passwordLockOut.sendKeys(Keys.SHIFT, Keys.HOME);
		passwordLockOut.sendKeys("1");
		passwordLockOut.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLockOutUnits));
		passwordLockOutUnits.sendKeys("Minute");

		Thread.sleep(2000);
		
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	
	
	
	public boolean checkLogoutAndLoginWithNewUserforInvalidAttempts()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		Thread.sleep(2000);

		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("focuS@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));

		String actErrMsg = mandatoryMsgs.getText();
		String expErrMsg = "This is your last chance to try the Password";
		System.out.println("Error Message=	" + actErrMsg);

		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("focuS@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);
		
		String compname2 = "Automation Company : 08/10/2020";

		Select oSelect2 = new Select(companyDropDownList);

		List<WebElement> elementCount2 = oSelect.getOptions();

		int cqSize2 = elementCount2.size();

		System.out.println("CompanyDropdownList Count :" + cqSize2);

		int i2;

		for (i2 = 0; i2 < elementCount2.size(); i2++) {

			elementCount2.get(i2).getText();

			String optionName = elementCount2.get(i2).getText();
			if (optionName.toUpperCase().startsWith(compname2.toUpperCase())) {
				System.out.println("q" + elementCount2.get(i2).getText());
				elementCount2.get(i2).click();
			}

		}

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));
		String actErrMsg1 = mandatoryMsgs.getText();
		String expErrMsg1 = "Invalid username (or) Password";

		System.out.println("Error Message=	" + actErrMsg1);

		getDriver().navigate().refresh();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("focuS@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);
		
		String compname1 = "Automation Company : 08/10/2020";

		Select oSelect1 = new Select(companyDropDownList);

		List<WebElement> elementCount1 = oSelect1.getOptions();

		int cqSize1 = elementCount1.size();

		System.out.println("CompanyDropdownList Count :" + cqSize);

		int i1;

		for (i1 = 0; i1 < elementCount1.size(); i1++) {

			elementCount1.get(i1).getText();

			String optionName = elementCount1.get(i1).getText();
			if (optionName.toUpperCase().startsWith(compname1.toUpperCase())) {
				System.out.println("q" + elementCount1.get(i1).getText());
				elementCount1.get(i1).click();
			}

		}

		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));
		String actErrMsg2 = mandatoryMsgs.getText();
		String expErrMsg2 = "Account has been locked";

		System.out.println("Error Message=	" + actErrMsg1);

		getDriver().navigate().refresh();
		Thread.sleep(45000);
		getDriver().navigate().refresh();
		Thread.sleep(5000);

		getDriver().navigate().refresh();
		Thread.sleep(5000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("Focus@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);
		String compname3 = "Automation Company : 08/10/2020";

		Select oSelect3 = new Select(companyDropDownList);

		List<WebElement> elementCount3 = oSelect.getOptions();

		int cqSize3 = elementCount3.size();

		System.out.println("CompanyDropdownList Count :" + cqSize3);

		//int i;

		for (i = 0; i < elementCount3.size(); i++) {

			elementCount3.get(i).getText();

			String optionName = elementCount3.get(i).getText();
			if (optionName.toUpperCase().startsWith(compname3.toUpperCase())) {
				System.out.println("q" + elementCount3.get(i).getText());
				elementCount3.get(i).click();
			}

		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNewUserName.getText();
		String expUserInfo = "NewUser";

		Thread.sleep(2000);
		;

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = "Automation Company ";
		//companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
		System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
				+ expLoginCompanyName + ".");

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)
				&& actErrMsg.equalsIgnoreCase(expErrMsg) && actErrMsg1.equalsIgnoreCase(expErrMsg1) && actErrMsg2.equalsIgnoreCase(expErrMsg2)) {

			return true;
		} else {

			return false;
		}
	}

	public boolean checkSendEMailonLoginSuccessinPasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(6000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendMailOnSuccessSelected));
		if (sendMailOnSuccess.isSelected() == false) {
			sendMailOnSuccessSelected.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginSuccessTxt));
			loginSuccessTxt.click();
			loginSuccessTxt.sendKeys("emailvalidationone@gmail.com");

		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	//@FindBy(xpath = "//*[@id=':2d']")
	@FindBy(xpath = "//*[@id=':2c']")
	private static WebElement emailBody;
	
	@FindBy(xpath = "//*[@id=':28']")
	private static WebElement emailBody2;

	@FindBy(xpath = "//tbody/tr/td[5]/div[1]/div/div/span/span")
	private static List<WebElement> emailSubjectList;
	
	@FindBy(xpath = "//*[@id=':28']/tbody/tr/td[5]/div/div/span[1]")
	private static List<WebElement> emailBodyList;

	@FindBy(xpath = "//tbody/tr[1]//td[5]//div[1]/div/div")
	private static WebElement emailRow1Subject;

	@FindBy(xpath = "//*[@id=':h5']")
	public static WebElement emailBodyText;

	@FindBy(xpath = "(//*[@id='gb']//iframe)[2]")
	private static WebElement SignOutFrame;

	//@FindBy(xpath = "//*[@id=':22']/div[1]/span")
	@FindBy(xpath="//*[@id=':1y']/div[1]/span")
	private static WebElement selectAllMailsChckbx;

	@FindBy(xpath = "//*[@id=':4']/div/div[1]/div[1]/div/div/div[2]/div[3]")
	private static WebElement selectAllMailsDltBtn;

	@FindBy(xpath="//div[@class='a3s aiL ']//div[2]")
	public static WebElement mailBodyText;
	
	
	public boolean checkEmailAfterSuccessfulLoginofNewUser() throws InterruptedException, AWTException {

		Thread.sleep(3000);

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
				"https://accounts.google.com/v3/signin/identifier?dsh=S-1191169840%3A1681369527989706&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ifkv=AQMjQ7TFRA0CZTpnwRmp0o0Ne4BHSlr0Ib3M8yvga7SW7c88_N4u5MYcHr87PB9-C1ewPLMsHGzfGg&rip=1&sacu=1&service=mail&flowName=GlifWebSignIn&flowEntry=ServiceLogin");

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

		Thread.sleep(8000);
		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
	

		String actMsg = "";
		String expMsg = "User NewUser has logged in Focus successfully";
		

		
		
		boolean result=false;
		 
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
 	 	
 		if (mailFromTxt.equalsIgnoreCase("Focus Login successful")) 
 		{
 			
 			//actMsg=emailBodyList.get(i).getText();
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
		actMsg=mailBodyText.getText();
		

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
			mailDeleteButton.click();

			Thread.sleep(2000);

			Thread.sleep(2000);
			
			if (actMsg.contains(expMsg)) 
			{
				result=true;
			}
			
	}
	 	 	}	
			ClickUsingJs(gmailUserBtn);
			

			Thread.sleep(3000);
			getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(1000);
			getAction().moveToElement(gmailSignOutBtn).click().build().perform();
			Thread.sleep(2000);
			
		
			 
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

			Thread.sleep(1000);

			getDriver().switchTo().window(newTabs.get(1)).close();
			Thread.sleep(2000);

			getDriver().switchTo().window(newTabs.get(0));
			Thread.sleep(2000);

			
			System.err.println("Message		:		" + actMsg + "Expected 		" + expMsg);
			
			if (result&&count!=0) 
			{
				
				return true;
			} 
			
			else
			{
				return false;
			}

		
		

	}

	public boolean checkSendEMailonLoginFailureinPasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendMailOnSuccessSelected));
		if (sendMailOnSuccess.isSelected() == true) {
			sendMailOnSuccessSelected.click();
		}
		Thread.sleep(2000);
		
		getAction().moveToElement(sendMailOnFailureSelected).build().perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendMailOnFailureSelected));
		if (sendMailOnFailure.isSelected() == false) {
			sendMailOnFailureSelected.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginFailureTxt));
			loginFailureTxt.click();
			loginFailureTxt.sendKeys("emailvalidationone@gmail.com");

		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkEmailAfterFailureLoginofNewUser() throws InterruptedException, AWTException
	{
		
		
		Thread.sleep(6000);

		

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

		Thread.sleep(8000);

		int count1 = row1mailFromList.size();
		System.out.println("No.of mails in the list" + count1);

		String actMsg = "";
		String expMsg = "Focus login has been failed for the user NewUser";
	

		
		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
		boolean result=false;
		 
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
 	 	
 		if (mailFromTxt.equalsIgnoreCase("Focus login failed")) 
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
		actMsg=mailBodyText.getText();
		System.err.println("Message		:		" + actMsg + "Expected 		" + expMsg);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
			mailDeleteButton.click();

			Thread.sleep(2000);

			Thread.sleep(2000);
			
			if (actMsg.contains(expMsg)) 
			{
				result=true;
			}
			
	}
	 	 	}	

		
			
			ClickUsingJs(gmailUserBtn);

			Thread.sleep(3000);
			getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(1000);
			getAction().moveToElement(gmailSignOutBtn).click().build().perform();

			

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

			Thread.sleep(1000);

			getDriver().switchTo().window(newTabs.get(1)).close();
			Thread.sleep(2000);

			getDriver().switchTo().window(newTabs.get(0));
			Thread.sleep(2000);

			

			if (result&&count!=0) 
			{
				
				return true;
			} 
			
			else
			{
				return false;
			}
		

	}

	public boolean checkChangePasswordafterFirstLogininPasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(8000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);
		
		getAction().moveToElement(sendMailOnSuccessSelected).build().perform();
		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendMailOnSuccessSelected));
		if (sendMailOnSuccess.isSelected() == true) {
			sendMailOnSuccessSelected.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendMailOnFailureSelected));
		if (sendMailOnFailure.isSelected() == true) {
			sendMailOnFailureSelected.click();
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordAfterFirstLoginSelected));
		if (changePasswordAfterFirstLogin.isSelected() == false) {
			changePasswordAfterFirstLoginSelected.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkCreateaNewUserforPasswordChangeafterFirstLogin()
			throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {

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
		createUserLoginNameTxt.sendKeys("DemoUser");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		createUserERPRoleDropDown.sendKeys("Test");

		int count = createUserERPRolesDropdownOptions.size() - 1;
		System.out.println("Count " + count);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordPolicyDropdown));
		createUserPasswordPolicyDropdown.sendKeys("Test");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		createUserPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		createUserConfirmPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		createUserNameTxt.sendKeys("DemoUser");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
		createUserLoginAbbrivationTxt.sendKeys("DemoUser");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityQuestionDropDown));
		Select s = new Select(createUserSecurityQuestionDropDown);
		s.selectByVisibleText("In which country were you born?");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));
		createUserSecurityAnswerTxt.sendKeys("India");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
		int counnt = createUserLanguageDropdownList.size() - 1;
		String actLanguageDropdownCount = Integer.toString(counnt);
		String expLanguageDropdownCount = "2";

		Set LanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < counnt; i++) {
			String data = createUserLanguageDropdownList.get(i).getText();
			LanguageDropdownList.add(data);
		}

		String actLanguageDropdownList = LanguageDropdownList.toString();
		String expLanguageDropdownList = "[English, Please select a language type] ";

		System.out.println("LanguageDropdownCount  : " + actLanguageDropdownCount + "  value expected  "
				+ expLanguageDropdownCount);
		System.out.println(
				"LanguageDropdownList   : " + actLanguageDropdownList + "  value expected  " + expLanguageDropdownList);

		Select language = new Select(getLanguage);
		language.selectByIndex(1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
		int cont = createUserAltLanguageDropdownList.size() - 1;
		String actAltLanguageDropdownCount = Integer.toString(count);
		String expAltLanguageDropdownCount = "";

		Set AltLanguageDropdownList = new HashSet<String>();

		for (int i = 0; i < cont; i++) {
			String data = createUserAltLanguageDropdownList.get(i).getText();
			AltLanguageDropdownList.add(data);
		}

		String actAltLanguageDropdownList = AltLanguageDropdownList.toString();
		String expAltLanguageDropdownList = "[English, Please select a language type] ";

		Select altlanguage = new Select(getAltLanguage);
		altlanguage.selectByIndex(2);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userCreationTimeZoneDrpdwn));
		Select timeZone = new Select(userCreationTimeZoneDrpdwn);
		timeZone.selectByVisibleText("(GMT+05:30) Chennai, Kolkatta, Mumbai, New Delhi");

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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();

		String expMessage = "new user created successfullyUser role is Restricted for change password screen,Cant change password after login";
		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkLogoutAndLoginWithDemoUser()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		Thread.sleep(2000);

		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("DemoUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("Focus@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(changePasswordCloseBtn));
		boolean actchangePasswordTitle = changePasswordTitle.isDisplayed();

		Thread.sleep(2000);

		if (actchangePasswordTitle) {

			return true;
		} else {

			return false;
		}
	}

	public boolean checkChangePasswordTitleaftertLogin()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOldPasswordTxt));
		changePasswordOldPasswordTxt.click();
		changePasswordOldPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordNewPasswordTxt));
		changePasswordNewPasswordTxt.sendKeys("focuS@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordConfirmTxt));
		changePasswordConfirmTxt.sendKeys("focuS@123");

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordSaveBtn));
		changePasswordSaveBtn.click();
		Thread.sleep(2000);

		getWaitForAlert();
		Thread.sleep(2000);
		String actMsg = getAlert().getText();
		getAlert().accept();
		Thread.sleep(2000);

		String expMsg = "Password Saved Successfully";

		Thread.sleep(2000);

		if (actMsg.equalsIgnoreCase(expMsg)) {
			return true;
		} else {
			return false;
		}

	}

	public boolean checkLoginWithSU()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		getDriver().navigate().refresh();
		Thread.sleep(2000);
		LoginPage lp = new LoginPage(getDriver());

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("su");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("su");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = "su";

		Thread.sleep(2000);
		;

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = "Automation Company ";
		//companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
		System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
				+ expLoginCompanyName + ".");

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {

			return true;
		} else {

			return false;
		}
	}

	public boolean checkCannotChangePasswordinPasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendMailOnSuccessSelected));
		if (sendMailOnSuccess.isSelected() == true) {
			sendMailOnSuccessSelected.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendMailOnFailureSelected));
		if (sendMailOnFailure.isSelected() == true) {
			sendMailOnFailureSelected.click();
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordAfterFirstLoginSelected));
		if (changePasswordAfterFirstLogin.isSelected() == true) {
			changePasswordAfterFirstLoginSelected.click();
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cannotChangePasswordSelected));
		if (cannotChangePassword.isSelected() == false) {
			cannotChangePasswordSelected.click();
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkCannotChangePasswordforNewUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		getDriver().navigate().refresh();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOption));
		changePasswordOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordOldPasswordTxt));
		changePasswordOldPasswordTxt.click();
		changePasswordOldPasswordTxt.sendKeys("Focus@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordNewPasswordTxt));
		changePasswordNewPasswordTxt.sendKeys("focuS@123");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordConfirmTxt));
		changePasswordConfirmTxt.sendKeys("focuS@123");

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordSaveBtn));
		changePasswordSaveBtn.click();
		Thread.sleep(2000);
		String expMsg = "You are not allowed to change the password";
		String actMsg = checkValidationMessage(expMsg);

		Thread.sleep(2000);

		if (actMsg.equalsIgnoreCase(expMsg)) {
			return true;
		} else {
			return false;
		}
	}

	public boolean checkDeletingPwdUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();
			System.out.println(actUser);

			if (actUser.equalsIgnoreCase("PwdUser")) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(2000);
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileDeleteIcon));
		createProfileDeleteIcon.click();

		getWaitForAlert();

		getAlert().accept();

		String expMessage = "User is deleted successfully";
		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();

			return true;
		} else {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();

			return false;
		}
	}

	public boolean checkOTPBasedLogininPasswordPolicy()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityPasswordPolicyMenu));
		homeSecurityPasswordPolicyMenu.click();

		Thread.sleep(2000);

		int passwordListCount = passwordList.size();

		for (int i = 0; i < passwordListCount; i++) {
			String data = passwordList.get(i).getText();

			if (data.equalsIgnoreCase("Test")) {
				passwordList.get(i).click();
				break;
			}
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordExpiryInMins));
		passwordExpiryInMins.click();
		passwordExpiryInMins.sendKeys(Keys.SHIFT, Keys.HOME);
		passwordExpiryInMins.sendKeys("10");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendMailOnSuccessSelected));
		if (sendMailOnSuccess.isSelected() == true) {
			sendMailOnSuccessSelected.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sendMailOnFailureSelected));
		if (sendMailOnFailure.isSelected() == true) {
			sendMailOnFailureSelected.click();
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(changePasswordAfterFirstLoginSelected));
		if (changePasswordAfterFirstLogin.isSelected() == true) {
			changePasswordAfterFirstLoginSelected.click();
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cannotChangePasswordSelected));
		if (cannotChangePassword.isSelected() == true) {
			cannotChangePasswordSelected.click();
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(otpBasedLoginSelected));
		if (otpBasedLogin.isSelected() == false) {
			otpBasedLoginSelected.click();
		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "Password policy updated successfully";

		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {

			return true;
		} else {

			return false;
		}

	}

	@FindBy(xpath = "//*[@id='OTP_Div_Int']/div[1]/div/label")
	public static WebElement otpLabel;

	@FindBy(xpath = "//*[@id='txtOTP']")
	public static WebElement otpTxt;

	@FindBy(xpath = "//input[@value='Submit']")
	public static WebElement otpSubmitBtn;

	@FindBy(xpath = "//*[@id='btnCancel']")
	public static WebElement otpCancelBtn;

	public boolean checkLogoutAndLoginWithNewUserforOTPLogin() throws InterruptedException {

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("Focus@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(otpLabel));

		if (otpLabel.isDisplayed()) {
			return true;
		} else {
			return false;
		}

	}

	@FindBy(xpath = "//*[@id=':h6']//br[1]")
	public static WebElement otpEmailTxt;
	
	@FindBy(xpath = "//a[contains(text(),'NewUser')]")
	public static WebElement userNamedisplay;
	

	public boolean checkEmailafterOTPSent() throws InterruptedException, AWTException, EncryptedDocumentException, InvalidFormatException, IOException {

		File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\StockLedgerEmailReport.pdf");

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
		ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
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

		Thread.sleep(5000);

		String actUserInfo = "";
		String expUserInfo = "";
		
		int count1 = row1mailFromList.size();
		System.out.println("No.of mails in the list" + count1);

		String actMsg = "";
		String expMsg = "One Time Password (OTP) to access your Login account is:";
	

		
		//ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());
		boolean result=false;
		 
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
 	 	
 		if (mailFromTxt.equalsIgnoreCase("One time password")) 
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
		actMsg=mailBodyText.getText();
		System.err.println("Message		:		" + actMsg );
			
			Thread.sleep(2000);

			Thread.sleep(2000);
			
			
	}
	 	 		


		//String actMsg = emailBody.getText();
		
		//System.err.println("Message		:		" + actMsg + "Expected 		" + expMsg);

		// return actMsg;

		Pattern p = Pattern.compile("[^0-9]");
		String numberOnly = p.matcher(actMsg).replaceAll("");
		System.out.println(numberOnly.substring(0, 6));

		Thread.sleep(2000);

		getDriver().switchTo().window(newTabs.get(0));
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(otpTxt));
		otpTxt.sendKeys(numberOnly.substring(0, 6));
		Thread.sleep(6000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(otpSubmitBtn));
		otpSubmitBtn.click();
		Thread.sleep(15000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNamedisplay));
		 actUserInfo = userNamedisplay.getText();
		 expUserInfo = "NewUser";

		System.out.println("Login With Valid User and Check User Name Display Value Actual : " + actUserInfo
				+ ". Value Expected : " + expUserInfo);

		// Login Company Name is Verify

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Company_Name));
		String actGetLoginCompanyNameInformation = Company_Name.getText();
		String actGetLoginCompanyName = actGetLoginCompanyNameInformation/*
																			 * .substring
																			 * (
																			 * 0,
																			 * 19)
																			 */;
		String expGetLoginCompanyName = "Automation Company ";
		//companyLogo.click();

		System.out.println("Company Name Display Value Actual                : " + actGetLoginCompanyName
				+ ". Value Expected : " + expGetLoginCompanyName);

		Thread.sleep(2000);

		getDriver().switchTo().window(newTabs.get(1));
		Thread.sleep(2000);
	

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mailDeleteButton));
			mailDeleteButton.click();

			Thread.sleep(2000);

			Thread.sleep(2000);
			
			if (actMsg.contains(expMsg)) 
			{
				result=true;
			}

	}
	
	 	ClickUsingJs(gmailUserBtn);

			Thread.sleep(3000);
			getDriver().switchTo().frame(SignOutFrame);
			Thread.sleep(1000);
			getAction().moveToElement(gmailSignOutBtn).click().build().perform();

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

			Thread.sleep(3000);

			getDriver().switchTo().window(newTabs.get(1)).close();
			Thread.sleep(4000);

			getDriver().switchTo().window(newTabs.get(0));
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			userNameDisplayLogo.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();

			Thread.sleep(2000);
			
			boolean login=CheckLogin();
			
			

			if (actMsg.contains(expMsg) && actUserInfo.equalsIgnoreCase(expUserInfo)) {

				return true;
			} else {

				return false;
			}

		

	}

	public static boolean CheckLogin()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getDriver().navigate().refresh();
		Thread.sleep(3000);
		
		getDriver().navigate().refresh();
		Thread.sleep(3000);

		LoginPage lp = new LoginPage(getDriver());

		lp.checkLoginPageTitleByURLInputInBrowser();

		String unamelt = "su";

		String pawslt = "su";

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

		String actUserInfo1 = userNameDisplay.getText();

		System.out.println("User Info  : " + actUserInfo1);

		System.out.println("User Info Capture Text  :  " + userNameDisplay.getText());

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt1 = Company_Name.getText();
		String getLoginCompanyName1 = getCompanyTxt1.substring(0, 31);
		System.out.println("company name  :  " + getLoginCompanyName1);
		//companyLogo.click();

		String expUserInfo1 = "SU";
		String expLoginCompanyName1 = "Automation Company : 08/10/2020";

		System.out.println("UserInfo1             : " + actUserInfo1 + " Value Expected : " + expUserInfo1);
		System.out.println(
				"LoginCompanyName1     : " + getLoginCompanyName1 + " Value Expected : " + expLoginCompanyName1);

		if (actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1)) {
			return true;
		} else {
			return false;
		}

	}

	public static boolean CheckLoginNewUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getDriver().navigate().refresh();

		LoginPage lp = new LoginPage(getDriver());

		lp.checkLoginPageTitleByURLInputInBrowser();

		String unamelt = "newuser";

		String pawslt = "Focus@123";

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

		String actUserInfo1 = userNewUserName.getText();

		System.out.println("User Info  : " + actUserInfo1);

		System.out.println("User Info Capture Text  :  " + userNewUserName.getText());

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt1 = Company_Name.getText();
		String getLoginCompanyName1 = getCompanyTxt1.substring(0, 31);
		System.out.println("company name  :  " + getLoginCompanyName1);
		//companyLogo.click();

		String expUserInfo1 = "SU";
		String expLoginCompanyName1 = "Automation Company : 08/10/2020";

		System.out.println("UserInfo1             : " + actUserInfo1 + " Value Expected : " + expUserInfo1);
		System.out.println(
				"LoginCompanyName1     : " + getLoginCompanyName1 + " Value Expected : " + expLoginCompanyName1);

		if (actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1)) {
			return true;
		} else {
			return false;
		}

	}

	public boolean checkLogoutSmokeSecurityPage()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		getDriver().navigate().refresh();
		Thread.sleep(2000);

		try {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			userNameDisplayLogo.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();

			Thread.sleep(2000);

			boolean actUserLoginPage = username.isDisplayed() && username.isEnabled() && password.isDisplayed()
					&& password.isEnabled();

			boolean expUserLoginPage = true;

			if (actUserLoginPage == expUserLoginPage) {
				System.out.println("***Test Pass: Login Successfull***");

				return true;
			} else {

				System.out.println("***Test Fail: Login Not Successfull***");

				return false;
			}
		} catch (Exception e) {
			String exception = e.getMessage();

			return false;
		}
	}

	
	
	public boolean checkLogoutSmokeSecurityPage11()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {
		getDriver().navigate().refresh();
		Thread.sleep(2000);

		try {
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			userNameDisplayLogo.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();

			Thread.sleep(2000);

			boolean actUserLoginPage = username.isDisplayed() && username.isEnabled() && password.isDisplayed()
					&& password.isEnabled();

			boolean expUserLoginPage = true;
			

			if (actUserLoginPage == expUserLoginPage) {
				System.out.println("***Test Pass: Login Successfull***");

				return true;
			
				
			} else {

				System.out.println("***Test Fail: Login Not Successfull***");

				return false;
			}
		} catch (Exception e) {
			String exception = e.getMessage();

			return false;
		}
	}

	
	
	public static boolean checkDetailsofSUUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase("su")) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(3000);

				break;
			}
		}
		Thread.sleep(2000);

		JavascriptExecutor jsExecutor = (JavascriptExecutor) getDriver();
		String actLoginUserName = (String) jsExecutor.executeScript("return arguments[0].value",createUserLoginNameTxt);

		String expLoginUserName = "SU";

		Thread.sleep(2000);

		if (actLoginUserName.equalsIgnoreCase(expLoginUserName)) {
			return true;
		} else {
			return false;
		}
	}

	public boolean checkLogoutAndLoginWithNewUserforInvalidAttemptsAgain()
			throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException {

		Thread.sleep(2000);

		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("focuS@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));

		String actErrMsg = mandatoryMsgs.getText();
		String expErrMsg = "This is your last chance to try the Password";
		System.out.println("Error Message=	" + actErrMsg);

		getDriver().navigate().refresh();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("focuS@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);
		
		String compname1 = "Automation Company : 08/10/2020";

		Select oSelect1 = new Select(companyDropDownList);

		List<WebElement> elementCount1 = oSelect1.getOptions();

		int cqSize1 = elementCount1.size();

		System.out.println("CompanyDropdownList Count :" + cqSize1);

		
		for (i = 0; i < elementCount1.size(); i++) {

			elementCount1.get(i).getText();

			String optionName = elementCount1.get(i).getText();
			if (optionName.toUpperCase().startsWith(compname1.toUpperCase())) {
				System.out.println("q" + elementCount1.get(i).getText());
				elementCount1.get(i).click();
			}

		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);
		
		getDriver().navigate().refresh();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("focuS@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);
		
		
		String compname2 = "Automation Company : 08/10/2020";

		Select oSelect2 = new Select(companyDropDownList);

		List<WebElement> elementCount2 = oSelect2.getOptions();

		int cqSize2 = elementCount2.size();

		System.out.println("CompanyDropdownList Count :" + cqSize2);

		for (i = 0; i < elementCount2.size(); i++) {

			elementCount2.get(i).getText();

			String optionName = elementCount2.get(i).getText();
			if (optionName.toUpperCase().startsWith(compname2.toUpperCase())) {
				System.out.println("q" + elementCount2.get(i).getText());
				elementCount2.get(i).click();
			}

		}

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);
		
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));
		String actErrMsg1 = mandatoryMsgs.getText();
		String expErrMsg1 = "Account has been locked";

		System.out.println("Error Message=	" + actErrMsg1);

		if (actErrMsg1.equalsIgnoreCase(expErrMsg1)) {
			getDriver().navigate().refresh();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
			username.sendKeys("su");
			getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
			password.sendKeys("su");
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);
			String compname3 = "Automation Company : 08/10/2020";

			Select oSelect3 = new Select(companyDropDownList);

			List<WebElement> elementCount3 = oSelect3.getOptions();

			int cqSize3 = elementCount3.size();

			System.out.println("CompanyDropdownList Count :" + cqSize3);

			
			for (i = 0; i < elementCount3.size(); i++) {

				elementCount3.get(i).getText();

				String optionName = elementCount3.get(i).getText();
				if (optionName.toUpperCase().startsWith(compname3.toUpperCase())) {
					System.out.println("q" + elementCount3.get(i).getText());
					elementCount3.get(i).click();
				}

			}

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
			signIn.click();
			Thread.sleep(3000);

		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		String actUserInfo = userNameDisplay.getText();
		String expUserInfo = "su";

		Thread.sleep(2000);
		;

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt = Company_Name.getText();
		String actLoginCompanyName = getCompanyTxt.substring(0, 19);
		String expLoginCompanyName = "Automation Company ";
		//companyLogo.click();
		Thread.sleep(2000);

		System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
		System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
				+ expLoginCompanyName + ".");

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {

			return true;
		} else {

			return false;
		}

	}

	public boolean checkDoNotLockAccountinNewUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase("NewUser")) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(4000);

				break;
			}
		}

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
		createUserAdditionalInfoTab.click();
		Thread.sleep(2000);
		if(createUserDoNotLockAccountChkBoxSelected.isSelected()==false)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserDoNotLockAccountChkBox));
		createUserDoNotLockAccountChkBox.click();
		}
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "User updated successfully.";
		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			return true;
		} else {
			return false;
		}

	}
	
	
	public boolean checkEnableAccountDisabledinNewUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase("NewUser")) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(4000);

				break;
			}
		}

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
		createUserAdditionalInfoTab.click();
		Thread.sleep(2000);
		if(createUserAccountDisabledChkBoxSelected.isSelected()==false)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAccountDisabledChkBox));
		createUserAccountDisabledChkBox.click();
		}
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "User updated successfully.";
		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			return true;
		} else {
			return false;
		}

	}



public boolean checkLogoutandLoginwithNewUserAfterAccountDisabledChkBoxEnabled() throws InterruptedException {

	Thread.sleep(2000);

	LoginPage lp = new LoginPage(getDriver());

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
	userNameDisplayLogo.click();

	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
	logoutOption.click();

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
	username.sendKeys("NewUser");
	getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
	password.sendKeys("Focus@123");
	getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
	signIn.click();
	Thread.sleep(6000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));

	String actErrMsg = mandatoryMsgs.getText();
	String expErrMsg = "Account disabled";
	System.out.println("Error Message=	" + actErrMsg);

	if (actErrMsg.equalsIgnoreCase(expErrMsg)) {

		return true;
	} else {
		return false;
	}

}
	
	
public boolean checkLoginWithSUAfterAccountDisabledChkBoxEnabled() throws InterruptedException {
	Thread.sleep(2000);

	getDriver().navigate().refresh();
	Thread.sleep(4000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
	username.sendKeys("su");
	getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
	password.sendKeys("su");
	getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
	signIn.click();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
	userNameDisplayLogo.click();

	String actUserInfo = userNameDisplay.getText();
	String expUserInfo = "su";

	Thread.sleep(2000);
	;

	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	companyLogo.click();*/

	String getCompanyTxt = Company_Name.getText();
	String actLoginCompanyName = getCompanyTxt.substring(0, 19);
	String expLoginCompanyName = "Automation Company ";
	//companyLogo.click();
	Thread.sleep(2000);

	System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
	System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
			+ expLoginCompanyName + ".");

	if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {

		return true;
	} else {

		return false;
	}

}

public boolean checkDisableAccountDisabledinNewUser()
		throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

	Thread.sleep(3000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
	securityMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
	createUserMenu.click();

	Thread.sleep(4000);

	int actUsersCount = createUserSavedUsersList.size();

	System.out.println(actUsersCount);

	for (int i = 0; i < actUsersCount; i++) {
		String actUser = createUserSavedUsersList.get(i).getText();

		if (actUser.equalsIgnoreCase("NewUser")) {
			createUserSavedUsersList.get(i).click();

			Thread.sleep(4000);

			break;
		}
	}

	Thread.sleep(3000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
	createUserAdditionalInfoTab.click();
	Thread.sleep(2000);
	if (createUserAccountDisabledChkBoxSelected.isSelected() == true) {

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAccountDisabledChkBox));
		createUserAccountDisabledChkBox.click();
	}
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
	SaveBtn.click();
	Thread.sleep(2000);

	String expMessage = "User updated successfully.";
	String actMessage = checkValidationMessage(expMessage);

	if (actMessage.equalsIgnoreCase(expMessage)) {
		return true;
	} else {
		return false;
	}

}




public boolean checkLogoutAndLoginWithNewUserAfterDisableAccountDisableChkBox() throws InterruptedException {
	Thread.sleep(2000);

	LoginPage lp = new LoginPage(getDriver());

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
	userNameDisplayLogo.click();

	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
	logoutOption.click();

	Thread.sleep(2000);
	
	getDriver().navigate().refresh();
	Thread.sleep(4000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
	username.sendKeys("NewUser");
	getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
	password.sendKeys("Focus@123");
	getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
	signIn.click();
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
	userNameDisplayLogo.click();

	String actUserInfo = userNewUserName.getText();
	String expUserInfo = "NewUser";

	Thread.sleep(2000);
	;

	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	companyLogo.click();*/

	String getCompanyTxt = Company_Name.getText();
	String actLoginCompanyName = getCompanyTxt.substring(0, 19);
	String expLoginCompanyName = "Automation Company ";
	//companyLogo.click();
	Thread.sleep(2000);

	System.out.println("UserInfo               : ." + actUserInfo + ".  value expected  ." + expUserInfo + ".");
	System.out.println("LoginCompanyName       : ." + actLoginCompanyName + ".  value expected  ."
			+ expLoginCompanyName + ".");

	if (actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)) {

		return true;
	} else {

		return false;
	}

}


	public void browser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException, AWTException {

		Thread.sleep(2000);
		checkLogoutAndLoginWithNewUser();
		Thread.sleep(3000);
		openBrowser("firefox");
		Thread.sleep(2000);
		LoginPage.checkLoginPageTitleByURLInputInBrowser();
		Thread.sleep(2000);
		CheckLoginNewUser();
		Thread.sleep(2000);
		checkLogoutSmokeSecurityPage11();
		Thread.sleep(3000);
		
		System.err.println("Window :	"+getDriver().getWindowHandles().size());
		
		
		getDriver().switchTo().defaultContent();
		Thread.sleep(2000);
		/*Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_F4);
		robot.keyRelease(KeyEvent.VK_ALT);
		robot.keyRelease(KeyEvent.VK_F4);

		Thread.sleep(3000);*/

		System.out.println("sdjhsjdhdhjs");
		
		
	
		Thread.sleep(3000);
		Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\DashboardClick.exe");
		
		Thread.sleep(10000);

		getDriver().navigate().refresh();
		System.out.println("After click");
		Thread.sleep(3000);
		
		

	}
	WebDriver driver;
	WebDriver driver1;
	
	@Test(priority=1)
	public void crossbrowser()
	{
		
		driver = new ChromeDriver();
		driver.get("http://localhost/FocusX#");
		
		
	}
	
	@Test(priority=2)
	public void crossbrowser1()
	{
		
		
		driver1 = new FirefoxDriver();
		driver1.get("http://localhost/FocusX#");
		driver1.close();
		
		
	}
	
	@Test(priority=3)
	public void crossbrowser3()
	{
		
		
		driver.findElement(By.id("txtUsername ")).sendKeys("sanjeev");
		
		System.out.println("hello");
		
		
		
	}

	
	public boolean checkLogoutAndLoginWithNewUserAfterDoNotLockCheckBoxwithWrongPassword() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{

		Thread.sleep(2000);

		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("focuS@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));

		String actErrMsg = mandatoryMsgs.getText();
		String expErrMsg = "Password can not be blank";
		System.out.println("Error Message=	" + actErrMsg);
		if(actErrMsg.equalsIgnoreCase(expErrMsg))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	public boolean checkEnableAllowMultiLoginCheckBoxinNewUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase("NewUser")) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(4000);

				break;
			}
		}

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
		createUserAdditionalInfoTab.click();
		Thread.sleep(2000);
		if(createUserAccountDisabledChkBoxSelected.isSelected()==true)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAccountDisabledChkBox));
		createUserAccountDisabledChkBox.click();
		}
		Thread.sleep(2000);
		
		if(createUserDoNotLockAccountChkBoxSelected.isSelected()==true)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserDoNotLockAccountChkBox));
		createUserDoNotLockAccountChkBox.click();
		}
		Thread.sleep(2000);
		if(createUserAllowMultiLoginChkBoxSelected.isSelected()==false)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAllowMultiLoginChkBox));
		createUserAllowMultiLoginChkBox.click();
		}
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "User updated successfully.";
		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			return true;
		} else {
			return false;
		}

	}
	
	public boolean  checkNewUsertoAllowMultiLogin() throws InterruptedException, AWTException
	{
		Thread.sleep(4000);
		
		WebDriverManager.firefoxdriver().setup();
		Thread.sleep(4000);
		driver1 = new FirefoxDriver();
		driver1.get("http://localhost/FocusX#");
		Thread.sleep(4000);
		
		login1();
		Thread.sleep(6000);
		Logout1();
		Thread.sleep(2000);
		
		getWaitForAlert();
		getAlert().accept();
		Thread.sleep(4000);
		
		Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_TAB);
		robot.keyRelease(KeyEvent.VK_ALT);
		robot.keyRelease(KeyEvent.VK_TAB);
		
		Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		
		String actUserInfo1 = userNewUserName.getText();

		System.out.println("User Info  : " + actUserInfo1);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_TAB);
		robot.keyRelease(KeyEvent.VK_ALT);
		robot.keyRelease(KeyEvent.VK_TAB);
		
		
		Thread.sleep(4000);
		driver1.close();
		Thread.sleep(4000);
		
		
		
		Capabilities cap = ((RemoteWebDriver) getDriver()).getCapabilities();
	    String browserName = cap.getBrowserName().toLowerCase();
	    System.out.println(browserName);
	    if(browserName.equalsIgnoreCase("chrome") && actUserInfo1.equalsIgnoreCase("NewUser"))
	    {
	    	return true;
	    }
	    else
	    {
	    	return false;
	    }
		
	}
	
	public void login1() throws InterruptedException
	{
		Thread.sleep(2000);
		driver1.navigate().refresh();
		Thread.sleep(2000);
		driver1.findElement(By.id("txtUsername")).sendKeys("newuser");
		Thread.sleep(2000);
		driver1.findElement(By.id("txtPassword")).click();
		driver1.findElement(By.id("txtPassword")).sendKeys("Focus@123");
		Thread.sleep(2000);
	
		String compname = "Automation Company : 08/10/2020";
		driver1.findElement(By.id("ddlCompany")).click();
		Thread.sleep(2000);
		Select oSelect = new Select(driver1.findElement(By.id("ddlCompany")));

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
		
		driver1.findElement(By.id("btnSignin")).click();
		Thread.sleep(4000);
	}
	
	public void Logout1() throws InterruptedException
	{
		Thread.sleep(2000);
		driver1.findElement(By.xpath("//*[@class='adminprofile']/a")).click();
		Thread.sleep(2000);
		driver1.findElement(By.xpath("//span[@class='icon icon-logout']")).click();
		Thread.sleep(2000);
	}
	
	
	
	public boolean checkDisableAllowMultiLoginCheckBoxinNewUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase("NewUser")) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(4000);

				break;
			}
		}

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
		createUserAdditionalInfoTab.click();
		Thread.sleep(2000);
		if(createUserAccountDisabledChkBoxSelected.isSelected()==true)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAccountDisabledChkBox));
		createUserAccountDisabledChkBox.click();
		}
		Thread.sleep(2000);
		
		if(createUserDoNotLockAccountChkBoxSelected.isSelected()==true)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserDoNotLockAccountChkBox));
		createUserDoNotLockAccountChkBox.click();
		}
		Thread.sleep(2000);
		if(createUserAllowMultiLoginChkBoxSelected.isSelected()==true)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAllowMultiLoginChkBox));
		createUserAllowMultiLoginChkBox.click();
		}
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "User updated successfully.";
		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			return true;
		} else {
			return false;
		}

	}
	
	
	public boolean  checkNewUserDonotAllowMultiLogin() throws InterruptedException, AWTException
	{
		Thread.sleep(4000);
		WebDriverManager.firefoxdriver().setup();
		driver1 = new FirefoxDriver();
		driver1.get("http://localhost/FocusX#");
		Thread.sleep(4000);
		
		login1();
		Thread.sleep(6000);
		Logout1();
		Thread.sleep(4000);
		Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_TAB);
		robot.keyRelease(KeyEvent.VK_ALT);
		robot.keyRelease(KeyEvent.VK_TAB);
		
		Thread.sleep(2000);
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		getDriver().navigate().refresh();
		Thread.sleep(4000);
		
		boolean actUserInfo1 = username.isDisplayed();

		System.out.println("User Info  : " + actUserInfo1);
		robot.keyPress(KeyEvent.VK_ALT);
		robot.keyPress(KeyEvent.VK_TAB);
		robot.keyRelease(KeyEvent.VK_ALT);
		robot.keyRelease(KeyEvent.VK_TAB);
		
		driver1.close();
		Thread.sleep(4000);
		
		Capabilities cap = ((RemoteWebDriver) getDriver()).getCapabilities();
	    String browserName = cap.getBrowserName().toLowerCase();
	    System.out.println(browserName);
	    if(browserName.equalsIgnoreCase("chrome") && actUserInfo1)
	    {
	    	return true;
	    }
	    else
	    {
	    	return false;
	    }
		
	}
	
	
	public boolean checkSendEmailNotificationinNewUser()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {

		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();

		Thread.sleep(4000);

		int actUsersCount = createUserSavedUsersList.size();

		System.out.println(actUsersCount);

		for (int i = 0; i < actUsersCount; i++) {
			String actUser = createUserSavedUsersList.get(i).getText();

			if (actUser.equalsIgnoreCase("NewUser")) {
				createUserSavedUsersList.get(i).click();

				Thread.sleep(4000);

				break;
			}
		}

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
		createUserAdditionalInfoTab.click();
		Thread.sleep(2000);
		if(createUserAccountDisabledChkBoxSelected.isSelected()==true)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAccountDisabledChkBox));
		createUserAccountDisabledChkBox.click();
		}
		Thread.sleep(2000);
		
		if(createUserDoNotLockAccountChkBoxSelected.isSelected()==true)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserDoNotLockAccountChkBox));
		createUserDoNotLockAccountChkBox.click();
		}
		Thread.sleep(2000);
		if(createUserAllowMultiLoginChkBoxSelected.isSelected()==true)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAllowMultiLoginChkBox));
		createUserAllowMultiLoginChkBox.click();
		}
		Thread.sleep(2000);

		if(createUserSendEmailNotificationChkBoxSelected.isSelected()==false)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSendEmailNotificationChkBox));
		createUserSendEmailNotificationChkBox.click();
		}
		Thread.sleep(2000);
		
		if(createUserEmailuserLoginSuccessChkBoxSelected.isSelected()==false)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailuserLoginSuccessChkBox));
		createUserEmailuserLoginSuccessChkBox.click();
		}
		Thread.sleep(2000);
		
		
		if(createUserEmailonLoginFailureChkBoxSelected.isSelected()==false)
		{

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailonLoginFailureChkBox));
		createUserEmailonLoginFailureChkBox.click();
		}
		Thread.sleep(2000);
		

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
		SaveBtn.click();
		Thread.sleep(2000);

		String expMessage = "User updated successfully.";
		String actMessage = checkValidationMessage(expMessage);

		if (actMessage.equalsIgnoreCase(expMessage)) {
			return true;
		} else {
			return false;
		}

	}
	
	
	public boolean checkEmailAfterSuccessfulLoginofUser() throws InterruptedException, AWTException
	{
		boolean actSuccess=checkEmailAfterSuccessfulLoginofNewUser();
		if(actSuccess)
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	public boolean checkLogoutandLoginNewUserforLoginFailure() throws InterruptedException
	{
		Thread.sleep(2000);

		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		username.sendKeys("NewUser");
		getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
		password.sendKeys("focuS@123");
		getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
		signIn.click();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));

		String actErrMsg = mandatoryMsgs.getText();
		String expErrMsg = "This is your last chance to try the Password";
		System.out.println("Error Message=	" + actErrMsg);
		if(actErrMsg.equalsIgnoreCase(expErrMsg))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	public boolean checkEmailAfterLoginFailureofUser() throws InterruptedException, AWTException
	{
		boolean actFailure=checkEmailAfterFailureLoginofNewUser();
		if(actFailure)
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	
		public boolean checkCreateUserFromLoadFromOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
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
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoadFrom));
			createUserLoadFrom.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoadFromNewUser));
			createUserLoadFromNewUser.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoadFormLoadBtn));
			createUserLoadFormLoadBtn.click();
			Thread.sleep(2000);
			
			String expMsg="User Loaded Successfully! Please Provide the LoginName and EmailID";
			String actMsg=checkValidationMessage(expMsg);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
			createUserLoginNameTxt.sendKeys("Load User");
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
			createUserAdditionalInfoTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailTxt));
			createUserEmailTxt.click();
			createUserEmailTxt.clear();
			createUserEmailTxt.sendKeys("emailvalidationthree1@gmail.com");
			Thread.sleep(2000);
			
			if(createUserSendEmailNotificationChkBoxSelected.isSelected()==true)
			{

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSendEmailNotificationChkBox));
			createUserSendEmailNotificationChkBox.click();
			}
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
			SaveBtn.click();
			Thread.sleep(2000);
			
			Date date = Calendar.getInstance().getTime();
			DateFormat dateFormat = new SimpleDateFormat("M/d/yyyy");
			String expuserCreatedDate = dateFormat.format(date);
			
			String actuserCreatedDate="";
			
		
			
			for(int i=0;i<createUserSearchUsersList.size();i++)
			{
				
				if(createUserSearchUsersList.get(i).getText().equalsIgnoreCase("Load User"))
				{
					actuserCreatedDate=createUsersDateList.get(i).getText();
					break;
					
				}
			}
			
			
			System.out.println(actuserCreatedDate);
			System.out.println(expuserCreatedDate);
			
			
			if(actMsg.equalsIgnoreCase(expMsg) )
			{
				return true;
			}
			else
			{
				return false;
			}
		}
			
		
		
		public boolean checkCreatingNewUserFromLoadFromSearchOption()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
			createUserIcon.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoadFrom));
			createUserLoadFrom.click();
			Thread.sleep(6000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoadFromSearchUser));
			createUserLoadFromSearchUser.click();
			createUserLoadFromSearchUser.sendKeys("UserAllOptionsST");

			Thread.sleep(4000);
			
			ArrayList<String> searchedUserList=new ArrayList<String>();
			
			for(WebElement e:createUserLoadUsersList)
			{
				searchedUserList.add(e.getText());
				
			}
			
			System.out.println("Searched UserList from Load From Option		"	+	searchedUserList.toString());
			
			createloadfromUseralloptionst.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoadFormLoadBtn));
			createUserLoadFormLoadBtn.click();
			Thread.sleep(4000);

			String expMsg="User Loaded Successfully! Please Provide the LoginName and EmailID";
			String actMsg=checkValidationMessage(expMsg);
			
			Thread.sleep(2000);
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
			createUserLoginNameTxt.click();
			createUserLoginNameTxt.sendKeys("Load User From Search");

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
			createUserAdditionalInfoTab.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailTxt));
			createUserEmailTxt.click();
			createUserEmailTxt.clear();
			createUserEmailTxt.sendKeys("emailvalidationfour4@gmail.com");
			Thread.sleep(2000);

			if (createUserSendEmailNotificationChkBox.isSelected() == true) {

				getFluentWebDriverWait()
						.until(ExpectedConditions.elementToBeClickable(createUserSendEmailNotificationChkBox));
				createUserSendEmailNotificationChkBox.click();
			}
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSaveIcon));
			createUserSaveIcon.click();
			Thread.sleep(2000);

			
			if (actMsg.equalsIgnoreCase(expMsg)) {
				return true;
			} else {
				return false;
			}
		}
	
	@FindBy(xpath="//*[@id='btnUnlockUser']")
	public static WebElement unlockUser;
	
	@FindBy(xpath="//*[@id='unlockbtbn']")
	public static WebElement unlockUserUnlockBtn;
	
	
	@FindBy(xpath="//*[@id='LoadLockedUsers']//div[3]//input[2]")
	public static WebElement unlockUserCancelBtn;
	
		
	//input[@id='11']//../span
	
	@FindBy(xpath="//input[@id='11']//../span")
	public static WebElement unlockUserChckBox;
	
	//ul[@id='loadLockedUsersList']
	
	
	@FindBy(xpath="//*[@id='loadLockedUsersList']//li//span")
	public static List<WebElement> unlockUsersList;
	
	
	
	
	@FindBy(xpath="//*[@id='maindiv']/nav/div/ol/li[2]/span")
	public static WebElement createUserLabel;
	
	@FindBy(xpath="//*[@id='unlockbtbn']")
	public static WebElement unlockBtn;
	
	
	
	
		public boolean checkUnlockUserOptioninCreateUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(3000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
			securityMenu.click();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
			createUserMenu.click();
		
			Thread.sleep(8000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unlockUser));
			unlockUser.click();
			Thread.sleep(2000);
			
			String actMsg,expMsg=null;
			
			if(!unlockUsersList.isEmpty())
			{
				for(int i=0;i<unlockUsersList.size();i++)
				{
					unlockUsersList.get(i).click();
					Thread.sleep(1000);
				}
				
				Thread.sleep(2000);
				unlockBtn.click();
				Thread.sleep(6000);
				 expMsg="Users Unlocked Successfully";
				 actMsg=checkValidationMessage(expMsg);
				 Thread.sleep(2000);
				 ClickUsingJs(unlockUserCancelBtn);
			}
			
			
			
			
			else
			{
			
				actMsg="No Users Locked";
				expMsg="User Unlocked";
				 ClickUsingJs(unlockUserCancelBtn);
			}
			
			System.out.println("Actual Message	"	+	actMsg);
			System.out.println("Expect Message	"	+	expMsg);
			
			
			if(actMsg.equalsIgnoreCase(expMsg))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		
		public boolean checkLogoutandLoginforUserstoLocktheUser() throws InterruptedException 
		{
			
			Thread.sleep(2000);

			LoginPage lp = new LoginPage(getDriver());

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
			userNameDisplayLogo.click();

			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
			username.sendKeys("NewUser");
			getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
			password.sendKeys("focuS@123");
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
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

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
			signIn.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));

			String actErrMsg = mandatoryMsgs.getText();
			String expErrMsg = "This is your last chance to try the Password";
			System.out.println("Error Message=	" + actErrMsg);

		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
			username.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			username.sendKeys("NewUser");
			getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
			password.sendKeys("focuS@123");
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);

			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
			signIn.click();
			Thread.sleep(2000);

			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
			username.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			username.sendKeys("NewUser");
			getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
			password.sendKeys("focuS@123");
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
			signIn.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));
			String actErrMsg1 = mandatoryMsgs.getText();
			String expErrMsg1 = "Account has been locked";
			
			///Demo User
			
			

			getDriver().navigate().refresh();
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
			username.sendKeys("DemoUser");
			getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
			password.sendKeys("FocuS@123");
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);

			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
			signIn.click();
			Thread.sleep(4000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));

			String actErrMsg11 = mandatoryMsgs.getText();
			String expErrMsg11 = "This is your last chance to try the Password";
			System.out.println("Error Message=	" + actErrMsg);

			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
			username.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			username.sendKeys("DemoUser");
			getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
			password.sendKeys("FocuS@123");
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);


			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
			signIn.click();
			Thread.sleep(2000);

			getDriver().navigate().refresh();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
			username.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			username.sendKeys("DemoUser");
			getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
			password.sendKeys("FocuS@123");
			getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
			signIn.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mandatoryMsgs));
			String actErrMsg12 = mandatoryMsgs.getText();
			String expErrMsg12 = "Account has been locked";
			
			
			if(actErrMsg1.equalsIgnoreCase(expErrMsg1) && actErrMsg12.equalsIgnoreCase(expErrMsg12))
			{
				return true;
			}
			else
			{
				return false;
			}
		
			
		}
		
		
		@FindBy(xpath="//*[@id='createGroupBtn']")
		public static WebElement createGroupUser;
		
	
		
		public boolean checkCreatingGroupUserinCreateUser() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
			securityMenu.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
			createUserMenu.click();

			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createGroupUser));
			createGroupUser.click();
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
			createUserLoginNameTxt.click();
			createUserLoginNameTxt.sendKeys("Group User");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
			createUserERPRoleDropDown.click();
			
			Select s=new Select(createUserERPRoleDropDown);
			s.selectByVisibleText("Test");
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
			getAltLanguage.click();
			Select s1=new Select(getAltLanguage);
			s1.selectByVisibleText("Arabic");
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
			createUserAdditionalInfoTab.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailTxt));
			createUserEmailTxt.click();
			createUserEmailTxt.clear();
			createUserEmailTxt.sendKeys("emailvalidationfour@gmail.com");
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
			SaveBtn.click();
			
			checkValidationMessage("");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createGroupUser));
			createGroupUser.click();
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
			createUserLoginNameTxt.click();
			createUserLoginNameTxt.sendKeys("Group User1");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
			createUserERPRoleDropDown.click();
			
			Select s2=new Select(createUserERPRoleDropDown);
			s2.selectByVisibleText("Test");
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
			getAltLanguage.click();
			Select s3=new Select(getAltLanguage);
			s3.selectByVisibleText("Arabic");
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
			createUserAdditionalInfoTab.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailTxt));
			createUserEmailTxt.click();
			createUserEmailTxt.clear();
			createUserEmailTxt.sendKeys("emailvalidationthree@gmail.com");
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
			SaveBtn.click();
			
			
			String expMsg="New group created successfully";
			String actMsg=checkValidationMessage(expMsg);
			
			System.out.println("Actual Message		"	+	actMsg);
			System.out.println("Expect Message		"	+	expMsg);
			
			if(expMsg.equalsIgnoreCase(actMsg))
			{
				return true;
			}
			else
			{
				return false;
			}
			
		}


		
		
		@FindBy(xpath="//*[@class='salesnav']//li//a")
		public static List<WebElement> groupUsersList;
		
		
		
		
		public boolean checkCreatingUserinGroupUser() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			
			for(int i=0;i<groupUsersList.size();i++)
			{
				String data=groupUsersList.get(i).getText();
				if(data.equalsIgnoreCase("Group User"))
				{
					groupUsersList.get(i).click();
					break;
				}
			}
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
			createUserIcon.click();
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
			createUserLoginNameTxt.click();
			createUserLoginNameTxt.sendKeys("Grp_User1");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
			createUserERPRoleDropDown.click();
			createUserERPRoleDropDown.sendKeys("RoleST");

			int count = createUserERPRolesDropdownOptions.size() - 1;
			System.out.println("Count " + count);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
			createUserPasswordTxt.sendKeys("12345");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
			createUserConfirmPasswordTxt.sendKeys("12345");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
			createUserNameTxt.sendKeys("Grp_User1");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
			createUserLoginAbbrivationTxt.sendKeys("GU");

			createUserSecurityQuestionDropDown.click();
			Select s = new Select(createUserSecurityQuestionDropDown);
			s.selectByVisibleText("In which country were you born?");//excelReader.getCellData(xlSheetName, 55, 5));

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));
			createUserSecurityAnswerTxt.sendKeys(excelReader.getCellData(xlSheetName, 56, 5));

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
			int counnt = createUserLanguageDropdownList.size();
			
			Set LanguageDropdownList = new HashSet<String>();

			for (int i = 1; i < counnt; i++) {
				String data = createUserLanguageDropdownList.get(i).getText();
				LanguageDropdownList.add(data);
			}

			
			Select language = new Select(getLanguage);
			language.selectByVisibleText(excelReader.getCellData(xlSheetName, 57, 5));

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
			int cont = createUserAltLanguageDropdownList.size();
			String actAltLanguageDropdownCount = Integer.toString(count);
			String expAltLanguageDropdownCount = excelReader.getCellData(xlSheetName, 51, 6);

			excelReader.setCellData(xlfile, xlSheetName, 51, 7, actAltLanguageDropdownCount);

			Set AltLanguageDropdownList = new HashSet<String>();

			for (int i = 1; i < cont; i++) {
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

			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SaveBtn));
			SaveBtn.click();
			Thread.sleep(2000);

			String expMessage = excelReader.getCellData(xlSheetName, 53, 6);
			String actMessage = checkValidationMessage(expMessage);

			excelReader.setCellData(xlfile, xlSheetName, 53, 7, actMessage);

			System.out.println(actMessage);
			System.out.println(expMessage);

			if (actMessage.equalsIgnoreCase(expMessage)) {
				excelReader.setCellData(xlfile, xlSheetName, 49, 8, resPass);
				return true;
			} else {
				excelReader.setCellData(xlfile, xlSheetName, 49, 8, resFail);
				return false;
			}
		
			
			
		}
		
		
		
		@FindBy(xpath = "//select[@id='timezone']")
		public static WebElement createUserSelectTimeZoneDropdown;
		
		@FindBy(xpath = "//*[@id='btnReset']")
		public static WebElement createUserReset;

		public boolean checkResetOptioninCreateUser()
				throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException {
			Thread.sleep(4000);

			int actUsersCount = createUserSavedUsersList.size();

			System.out.println(actUsersCount);

			for (int i = 0; i < actUsersCount; i++) {
				String actUser = createUserSavedUsersList.get(i).getText();

				if (actUser.equalsIgnoreCase("Load User")) {
					createUserSavedUsersList.get(i).click();

					Thread.sleep(3000);

					break;
				}
			}
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserReset));
			createUserReset.click();
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
			createUserERPRoleDropDown.click();
			createUserERPRoleDropDown.sendKeys("Test");

			int count = createUserERPRolesDropdownOptions.size() - 1;
			System.out.println("Count " + count);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
			createUserPasswordTxt.sendKeys("Test@123");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
			createUserConfirmPasswordTxt.sendKeys("Test@123");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
			createUserNameTxt.sendKeys("Load User");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
			createUserLoginAbbrivationTxt.sendKeys("LU");

			createUserSecurityQuestionDropDown.click();
			Select s = new Select(createUserSecurityQuestionDropDown);
			s.selectByVisibleText("In which country were you born?");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));
			createUserSecurityAnswerTxt.sendKeys("India");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
			getLanguage.click();
			int counnt = createUserLanguageDropdownList.size() - 1;
			String actLanguageDropdownCount = Integer.toString(counnt);
			String expLanguageDropdownCount = "2";

			Set LanguageDropdownList = new HashSet<String>();

			for (int i = 1; i < counnt; i++) {
				String data = createUserLanguageDropdownList.get(i).getText();
				LanguageDropdownList.add(data);
			}

			String actLanguageDropdownList = LanguageDropdownList.toString();
			String expLanguageDropdownList = "[English, Arabic]";

			System.out.println("LanguageDropdownCount  : " + actLanguageDropdownCount + "  value expected  "
					+ expLanguageDropdownCount);
			System.out.println(
					"LanguageDropdownList   : " + actLanguageDropdownList + "  value expected  " + expLanguageDropdownList);

			getLanguage.sendKeys(Keys.TAB);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
			getAltLanguage.click();
			int cont = createUserAltLanguageDropdownList.size() - 1;
			String actAltLanguageDropdownCount = Integer.toString(count);
			String expAltLanguageDropdownCount = "2";

			Set AltLanguageDropdownList = new HashSet<String>();

			for (int i = 1; i < cont; i++) {
				String data = createUserAltLanguageDropdownList.get(i).getText();
				AltLanguageDropdownList.add(data);
			}

			String actAltLanguageDropdownList = AltLanguageDropdownList.toString();
			String expAltLanguageDropdownList = "[English, Arabic]";

			getAltLanguage.sendKeys(Keys.TAB);

			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSelectTimeZoneDropdown));
			createUserSelectTimeZoneDropdown.click();
			Select s2 = new Select(createUserSelectTimeZoneDropdown);
			s2.selectByValue("63");
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserAdditionalInfoTab));
			createUserAdditionalInfoTab.click();

			Thread.sleep(2000);

		/*	JavascriptExecutor js = (JavascriptExecutor) driver;
			String script = "window.confirm = function(message) {"
					+ "document.lastConfirmationMessage = message; return true; }";
			js.executeScript(script);*/

			/*
			 * if(getIsAlertPresent()) { getAlert().accept(); }
			 */

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserUserTypeDrpDwn));
			Select user = new Select(createUserUserTypeDrpDwn);
			user.selectByVisibleText("Customer");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserEmailTxt));
			createUserEmailTxt.click();
			createUserEmailTxt.clear();
			createUserEmailTxt.sendKeys("emailvalidationthree@gmail.com");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSaveIcon));
			createUserSaveIcon.click();

			String expMessage = "User updated successfully.";
			String actMessage = checkValidationMessage(expMessage);

			System.out.println(actMessage);
			System.out.println(expMessage);

			if (actMessage.equalsIgnoreCase(expMessage)) {

				return true;
			} else {

				return false;
			}

		}
		
		
		
		
		public boolean checkSearchOptioninCreateUser()
				throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSearchTxt));
			createUserSearchTxt.click();

			createUserSearchTxt.sendKeys("Load User");
			Thread.sleep(4000);

			int actUsersCount = createUserSearchUsersList.size();

			System.out.println(actUsersCount);

			ArrayList<String> actUsersList = new ArrayList<String>();

			for (int i = 0; i < actUsersCount; i++) {
				String actUser = createUserSearchUsersList.get(i).getText();
				actUsersList.add(actUser);
			}

			String actUserList = actUsersList.toString();
			String expUserList = "[Load User, Load User From Search]";

			System.out.println("Actual Users List After Search			" + actUserList);
			System.out.println("Expected  Users List After Search		" + expUserList);

			for (int i = 0; i < actUsersCount; i++) {
				String actUser = createUserSearchUsersList.get(i).getText();

				if (actUser.equalsIgnoreCase("Load User")) {
					createUserSearchUsersList.get(i).click();

					Thread.sleep(3000);

					break;
				}
			}
			Thread.sleep(6000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSaveIcon));
			createUserSaveIcon.click();

			String expMessage = "User updated successfully.";
			String actMessage = checkValidationMessage(expMessage);

			System.out.println(actMessage);
			System.out.println(expMessage);
			Thread.sleep(2000);

			if (actMessage.equalsIgnoreCase(expMessage)) {

				return true;
			} else {

				return false;
			}

		}
		
		
	public static boolean checkDonotallowcopypasteinprofileinuserlevel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		createProfileProfileNameCombo.click();
		createProfileProfileNameCombo.sendKeys("All Profile ST");
		Thread.sleep(1000);
		createProfileProfileNameCombo.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileFinancalsExpandBtn));
		createProfileFinancalsExpandBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsExpandBtn));
		financialsTransactionsExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchasesExpandBtn));
		financialsTransactionsPurchasesExpandBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchasesVoucherOption));
		financialsTransactionsPurchasesVoucherOption.click();
		
		int count=Purchasevoucherexpandlabels.size();
		ArrayList<String> labels=new ArrayList<String>();
		
		for(int i=0;i<count;i++)
		{
			String data=Purchasevoucherexpandlabels.get(i).getText();
			System.out.println(data);
		}
		
		ClickUsingJs(selectall);
		
		if(CheckDoNotAllowCopyPaste.isSelected()==true)
		{
			ClickUsingJs(DoNotAllowCopyPaste);
		}
		
		boolean actcheck=CheckDoNotAllowCopyPaste.isSelected();
		boolean expcheck=false;
		
		System.err.println("Actual Do not allow copy/paste is   : "+ actcheck);
		System.err.println("Expected Do not allow copy/paste is : "+ expcheck);
		
		SaveBtn.click();
		
		String expmsg="Profile Updated Successfully";
		String actmsg=checkValidationMessage(expmsg);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityCreateRoleMenu));
		homeSecurityCreateRoleMenu.click();

		Thread.sleep(1500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		createRoleRoleNameCombo.sendKeys("Testing");
		Thread.sleep(1500);
		createRoleRoleNameCombo.sendKeys(Keys.TAB);
		Thread.sleep(6000);

		int actAssignedProfilesListCount = assignedProfileList.size();

		Set<String> AssignedProfilesList = new HashSet<String>();

		for (int i = 0; i < actAssignedProfilesListCount; i++) {
			String profile = assignedProfileList.get(i).getText();
			Thread.sleep(2000);
			AssignedProfilesList.add(profile);
		}

		String actAssignedProfilesList = AssignedProfilesList.toString();

		String expAssignedProfilesList = "[All Profile ST]";
		
		System.out.println(actAssignedProfilesList);
		System.out.println(expAssignedProfilesList);
		
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(additionTab));
		additionTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleAddTabFinancialsMenuExpand));
		createRoleAddTabFinancialsMenuExpand.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RoleFinancialTransaction));
		RoleFinancialTransaction.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RoleFinancialTransactionPurchase));
		RoleFinancialTransactionPurchase.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RoleFinancialTransactionPurchaseVoucher));
		RoleFinancialTransactionPurchaseVoucher.click();

		Thread.sleep(2000);
		
		if(AddDoNotAllowCopyPasteText.isSelected()==true)
			{
				ClickUsingJs(AddDoNotAllowCopyPaste);
			}
		
			boolean actcheck1=AddDoNotAllowCopyPasteText.isSelected();
			boolean expcheck1=false;
		
			System.err.println("Actual Role Do not allow copy/paste is   : "+ actcheck1);
			System.err.println("Expected Role Do not allow copy/paste is : "+ expcheck1);
		
			Thread.sleep(1500);
		
		SaveBtn.click();
		
		String expmsg1="Role updated Successfully";
		String actmsg1=checkValidationMessage(expmsg1);
	
		if(actmsg.equalsIgnoreCase(expmsg) && actcheck==expcheck 
				&& actmsg1.equalsIgnoreCase(expmsg1))
		{
		return true;
		}
		else
		{
			return false;
		}
	}
	
	
	
	public SmokeSecurityPage(WebDriver driver) {
		PageFactory.initElements(driver, this);
	}

	public static boolean checkLogoutAndLoginTestUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		
		Thread.sleep(3000);
		
		LoginPage lp = new LoginPage(getDriver());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Usernamearrow));
		Usernamearrow.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();


		getDriver().navigate().refresh();
		Thread.sleep(6000);

		lp.checkLoginPageTitleByURLInputInBrowser();
		

		String unamelt = "Test";

		String pawslt = "54321";

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

		Thread.sleep(5000);

		String actUserInfo1 = TestUsername.getText();

		System.out.println("User Info  : " + actUserInfo1);

		System.out.println("User Info Capture Text  :  " + TestUsername.getText());

		String getCompanyTxt1 = Company_Name.getText();
		String getLoginCompanyName1 = getCompanyTxt1.substring(0, 31);
		System.out.println("company name  :  " + getLoginCompanyName1);
		//companyLogo.click();

		String expUserInfo1 = "Test";
		String expLoginCompanyName1 = "Automation Company : 08/10/2020";

		System.out.println("UserInfo1             : " + actUserInfo1 + " Value Expected : " + expUserInfo1);
		System.out.println(
				"LoginCompanyName1     : " + getLoginCompanyName1 + " Value Expected : " + expLoginCompanyName1);

		if (actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1)) {
			return true;
		} else {
			return false;
		}

	}

	public static boolean VerifyDonotallowcopypasteDisplayedInPurchasevoucherInuserlevel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	{
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucher));
		purchaseVoucher.click();

		Thread.sleep(5500);
		
		newBtn.click();
		
		Thread.sleep(2000);
		
		//checkValidationMessage("Screen Opened");
		
		Thread.sleep(2000);
		
		toggleBtn.click();
		
		Thread.sleep(2000);
		
		int count=Toggledropdowns.size();
		ArrayList<String> actToggleDropdwonList=new ArrayList<String>();
		
		Thread.sleep(2000);
		for(int i=0;i<count;i++)
		{
			String DropdownList=Toggledropdowns.get(i).getText();
			
			if(DropdownList.isEmpty()==false)
			{
				actToggleDropdwonList.add(DropdownList);
			}
			
		}
		
		String actdropdownlist=actToggleDropdwonList.toString();
		String expdropdownlist="[Copy Document, Copy to Clipboard, Paste from Clipboard, Add To Stock, Posting details, BOM Input, Print barcode, Close Links, Export to XML, Setting, Calculator, Auto Load]";
		
		Thread.sleep(2000);
		
		System.out.println("act dropdown list : " + actdropdownlist);
		System.out.println("exp dropdown list : " + expdropdownlist);
		
		ClickUsingJs(new_CloseBtn);
		
		Thread.sleep(2000);
		
		logout();
		
		getDriver().navigate().refresh();
		Thread.sleep(6000);
		
		CheckLogin();
		
		Thread.sleep(3000);
		
		if(actdropdownlist.equalsIgnoreCase(expdropdownlist))
		{
		return true;
		}
		else
		{
			return false;
		}
	}

	public static boolean EnableDonotallowcopypasteinprofileInSuUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	{
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		createProfileProfileNameCombo.click();
		createProfileProfileNameCombo.sendKeys("All Profile ST");
		Thread.sleep(1000);
		createProfileProfileNameCombo.sendKeys(Keys.TAB);
		
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
		
		int count=Purchasevoucherexpandlabels.size();
		ArrayList<String> labels=new ArrayList<String>();
		
		for(int i=0;i<count;i++)
		{
			String data=Purchasevoucherexpandlabels.get(i).getText();
			System.out.println(data);
		}
	
		if(CheckDoNotAllowCopyPaste.isSelected()==false)
		{
			ClickUsingJs(DoNotAllowCopyPaste);
		}
		
		boolean actcheck=CheckDoNotAllowCopyPaste.isSelected();
		boolean expcheck=true;
		
		System.err.println("Actual Do not allow copy/paste is   :"+ actcheck);
		System.err.println("Expected Do not allow copy/paste is :"+ expcheck);
		
		SaveBtn.click();
		
		Thread.sleep(2000);
		
		String expmsg="Profile Updated Successfully";
		String actmsg=checkValidationMessage(expmsg);
		
		if(actcheck==expcheck && expmsg.equalsIgnoreCase(actmsg))
		{
		return true;
		}
		
		else
		{
		return false;
		}
	}

	public static boolean VerifyDonotallowcopypasteInPurchasevoucherInuserlevel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	{
		Thread.sleep(2000);
		
		checkLogoutAndLoginTestUser();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		financialsMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
		financialsTransactionMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
		financialsTransactionsPurchaseMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucher));
		purchaseVoucher.click();

		Thread.sleep(2000);
		
		newBtn.click();
		
		Thread.sleep(2000);
		
		//checkValidationMessage("Screen Opened");
		
		Thread.sleep(2000);
		
		toggleBtn.click();
		
		Thread.sleep(2000);
		
		int count=Toggledropdowns.size();
		ArrayList<String> actToggleDropdwonList=new ArrayList<String>();
		
		Thread.sleep(2000);
		for(int i=0;i<count;i++)
		{
			String DropdownList=Toggledropdowns.get(i).getText();
			
			if(DropdownList.isEmpty()==false)
			{
				actToggleDropdwonList.add(DropdownList);
			}
			
		}
		
		String actdropdownlist=actToggleDropdwonList.toString();
		String expdropdownlist="[Add To Stock, Posting details, BOM Input, Print barcode, Close Links, Export to XML, Setting, Calculator, Auto Load]";
		
		Thread.sleep(2000);
		
		System.out.println("act dropdown list : " + actdropdownlist);
		System.out.println("exp dropdown list : " + expdropdownlist);
		
		ClickUsingJs(new_CloseBtn);
		
		 
		if(actdropdownlist.equalsIgnoreCase(expdropdownlist))
		{
		return true;
		}
		else
		{
			return false;
		}
	}

	public static boolean SavingRolewithprofile() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	{
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityCreateRoleMenu));
		homeSecurityCreateRoleMenu.click();

		Thread.sleep(1500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		createRoleRoleNameCombo.sendKeys("Testing");
		createRoleRoleNameCombo.sendKeys(Keys.TAB);
		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(availableProfilepv));
		ClickUsingJs(availableProfilepv);

		Thread.sleep(5000);

		getFluentWebDriverWait()
				.until(ExpectedConditions.elementToBeClickable(createRoleMOveFiledsFromLeftSideToRightSide));
		createRoleMOveFiledsFromLeftSideToRightSide.click();

		Thread.sleep(8000);

		int actAssignedProfilesListCount = assignedProfileList.size();

		Set<String> AssignedProfilesList = new HashSet<String>();

		for (int i = 0; i < actAssignedProfilesListCount; i++) {
			String profile = assignedProfileList.get(i).getText();
			Thread.sleep(2000);
			AssignedProfilesList.add(profile);
		}

		String actAssignedProfilesList = AssignedProfilesList.toString();

		String expAssignedProfilesList = "[All Profile ST,profilepv]";
		
		System.out.println(actAssignedProfilesList);
		System.out.println(expAssignedProfilesList);
		
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
		createRoleSaveicon.click();

		String expMessage = "Role saved Successfully";

		String actMessage = checkValidationMessage(expMessage);
		
		logout();
		
		getDriver().navigate().refresh();
		Thread.sleep(3000);
		
		CheckLogin();

		if(actAssignedProfilesList.equalsIgnoreCase(expAssignedProfilesList) 
				&& actMessage.equalsIgnoreCase(expMessage))
		{
		
		return true;
		}
		else
		{
			return false;
		}
	}

	public static boolean checkDonotallowcopypastewithRoleinUserlevel() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	{
		Thread.sleep(3000);

		logout();
		
		getDriver().navigate().refresh();
		Thread.sleep(3000);
		
		CheckLogin();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityCreateRoleMenu));
		homeSecurityCreateRoleMenu.click();

		Thread.sleep(1500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		createRoleRoleNameCombo.sendKeys("Testing");
		Thread.sleep(1500);
		createRoleRoleNameCombo.sendKeys(Keys.TAB);
		Thread.sleep(6000);

		int actAssignedProfilesListCount = assignedProfileList.size();

		Set<String> AssignedProfilesList = new HashSet<String>();

		for (int i = 0; i < actAssignedProfilesListCount; i++) {
			String profile = assignedProfileList.get(i).getText();
			Thread.sleep(2000);
			AssignedProfilesList.add(profile);
		}

		String actAssignedProfilesList = AssignedProfilesList.toString();

		String expAssignedProfilesList = "[All Profile ST]";
		
		System.out.println(actAssignedProfilesList);
		System.out.println(expAssignedProfilesList);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exclusionsTab));
		exclusionsTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleExcTabFinancialsMenuExpand));
		createRoleExcTabFinancialsMenuExpand.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RoleExcFinancialTransaction));
		RoleExcFinancialTransaction.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RoleExcPurchase));
		RoleExcPurchase.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RoleExcPurchaseVoucher));
		RoleExcPurchaseVoucher.click();

		Thread.sleep(2000);
		
		if(ExcCheckDoNotAllowCopyPaste.isSelected()==true)
		{
			ClickUsingJs(ExcDoNotAllowCopyPaste);
		}
		
		boolean actcheck=ExcCheckDoNotAllowCopyPaste.isSelected();
		boolean expcheck=false;
		
		System.err.println("Actual Do not allow copy/paste is   :"+ actcheck);
		System.err.println("Expected Do not allow copy/paste is :"+ expcheck);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
		createRoleSaveicon.click();

		String expMessage = "Role updated Successfully";

		String actMessage = checkValidationMessage(expMessage);
		
		checkLogoutAndLoginTestUser();
		
		
		if(actAssignedProfilesList.equalsIgnoreCase(expAssignedProfilesList) 
				&& actcheck==expcheck && expMessage.equalsIgnoreCase(actMessage) )
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
}

