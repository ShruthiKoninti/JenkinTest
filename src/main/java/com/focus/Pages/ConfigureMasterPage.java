package com.focus.Pages;

import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class ConfigureMasterPage extends BaseEngine
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "ConfigureMasterPage";
	

	public static String checkValidationMessageString(String ExpMessage) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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





	public boolean checkconfigureMasters() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
		Thread.sleep(2000);
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,300)","");
		Thread.sleep(2000);
		
		

		int count = settingsSubMenusList.size();
		
		for (int i = 1; i < count; i++) 
		{
			String data = settingsSubMenusList.get(i).getText();
			System.err.println(data);
			
			JavascriptExecutor js = (JavascriptExecutor) getDriver(); 
			
			js.executeScript("arguments[0].scrollIntoView();", settingsSubMenusList.get(i));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
			{
				settingsSubMenusList.get(i).click();
				break;
			}
		}		

		Thread.sleep(2000);
		
		int configMastersRibbonControlCount = configMastersRibbonControl.size();

		ArrayList<String> configMastersRibbonControlArray = new ArrayList<String>();

		for(int i=0;i<configMastersRibbonControlCount;i++)
		{
			String data = configMastersRibbonControl.get(i).getAttribute("title");
			configMastersRibbonControlArray.add(data);
		}

		String actconfigMastersRibbonControl = configMastersRibbonControlArray.toString();

		String expconfigMastersRibbonControl = excelReader.getCellData(xlSheetName, 8, 6);

		System.out.println("configMastersRibbonControl Actual   : " + actconfigMastersRibbonControl);
		System.out.println("configMastersRibbonControl Expected : " + expconfigMastersRibbonControl);
		
		 excelReader.setCellData(xlfile, xlSheetName, 8, 7, actconfigMastersRibbonControl);

		if(actconfigMastersRibbonControl.equalsIgnoreCase(expconfigMastersRibbonControl))
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

@FindBy(xpath="//*[@id='DefinationTab']")
private static WebElement definitionTab;
	


	public static boolean checkSavingMasterInConfigureMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationTab));
		customizationTab.click();

		String expMessageOnClickingCustomizeTab = excelReader.getCellData(xlSheetName, 9, 6);

		String actMessageOnClickingCustomizeTab = checkValidationMessageString(expMessageOnClickingCustomizeTab);
		
		 excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessageOnClickingCustomizeTab);


		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionTab));
		definitionTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cmSaveBtn));
		cmSaveBtn.click();

		String expMessageOnSaveEmpty = excelReader.getCellData(xlSheetName, 10, 6);

		String actMessageOnSaveEmpty = checkValidationMessageString(expMessageOnSaveEmpty);
		
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actMessageOnSaveEmpty);
		
		//click(definationTab);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameComboBox));
		nameComboBox.click();
		nameComboBox.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));

		Thread.sleep(2000);
		nameComboBox.sendKeys(Keys.TAB);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cmSaveBtn));
		cmSaveBtn.click();

		String expMessageOnSaveNameInput = excelReader.getCellData(xlSheetName, 11, 6);

		String actMessageOnSaveNameInput = checkValidationMessageString(expMessageOnSaveNameInput);
		
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actMessageOnSaveNameInput);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenuExpandBtn));
		homeMenuExpandBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterMenuBtn));
		masterMenuBtn.click();

		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameComboBox));
			nameComboBox.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(captionTxt));
			captionTxt.click();*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cmSaveBtn));
		cmSaveBtn.click();

		String expMessageOnSave = excelReader.getCellData(xlSheetName, 12, 6);

		String actMessageOnSave = checkValidationMessageString(expMessageOnSave);
		
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actMessageOnSave);

		if(actMessageOnClickingCustomizeTab .equalsIgnoreCase(expMessageOnClickingCustomizeTab ) && actMessageOnSaveEmpty.equalsIgnoreCase(expMessageOnSaveEmpty) 
				&& actMessageOnSaveNameInput.equalsIgnoreCase(expMessageOnSaveNameInput) && actMessageOnSave.equalsIgnoreCase(expMessageOnSave))
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





	


	public boolean checkSavedMasterAndUpdatingTheMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameComboBox));
		nameComboBox.click();
		nameComboBox.sendKeys(Keys.END);
		nameComboBox.sendKeys(Keys.SHIFT,Keys.HOME);
		nameComboBox.sendKeys(Keys.BACK_SPACE);
		nameComboBox.sendKeys(Keys.SPACE);

		int nameComboBoxListCount = nameComboBoxList.size();

		System.err.println("nameComboBoxListCount : "+nameComboBoxListCount);

		ArrayList<String> nameComboBoxListArray = new ArrayList<String>();

		for(int i=0;i<nameComboBoxListCount;i++)
		{
			String data = nameComboBoxList.get(i).getText();

			System.err.println("Data : " + data);

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 13, 5)))
			{
				nameComboBoxList.get(i).click();
				break;
			}
		}

		nameComboBox.sendKeys(Keys.TAB);


		String actnameComboBox = nameComboBox.getAttribute("value");
		String expnameComboBox = excelReader.getCellData(xlSheetName, 13, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actnameComboBox);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(captionTxt));

		String actcaptionTxt = captionTxt.getAttribute("value");
		String expcaptionTxt = excelReader.getCellData(xlSheetName, 14, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actnameComboBox);


		//boolean actcmMastermenu = cmMastermenu.isDisplayed();  Boolean.toString(masterNewBtn.isDisplayed())
		
		String actcmMastermenu = Boolean.toString(cmMastermenu.isDisplayed());
		//boolean expcmMastermenu = true;
		String expcmMastermenu = excelReader.getCellData(xlSheetName, 15, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actcmMastermenu);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(captionTxt));
		captionTxt.click();
		captionTxt.clear();
		captionTxt.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		captionTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cmSaveBtn));
		cmSaveBtn.click();

		String expMessageOnSaveUpdate = excelReader.getCellData(xlSheetName, 16, 6);

		String actMessageOnSaveUpdate = checkValidationMessageString(expMessageOnSaveUpdate);
		
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actMessageOnSaveUpdate);

		System.out.println("nameComboBox Actual            : " + actnameComboBox          + " Value Expected : "+expnameComboBox);
		System.out.println("captionTxt Actual              : " + actcaptionTxt            + " Value Expected : "+expcaptionTxt);
		System.out.println("cmMastermenu Actual            : " + actcmMastermenu          + " Value Expected : "+expcmMastermenu);

		if(actnameComboBox.equalsIgnoreCase(expnameComboBox) && actcaptionTxt.equalsIgnoreCase(expcaptionTxt)
				&& actcmMastermenu.equalsIgnoreCase(expcmMastermenu) && actMessageOnSaveUpdate.equalsIgnoreCase(expMessageOnSaveUpdate))
		{	
			System.out.println("Test Pass : Bin Popup Displayed");
			excelReader.setCellData(xlfile, xlSheetName, 13, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Bin Popup NOT Displayed");
			excelReader.setCellData(xlfile, xlSheetName, 13, 8, resFail);
			return false;
		}
	}




	

	public static boolean checkUpdatedMasterAndDeleteThatMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameComboBox));
		nameComboBox.click();
		nameComboBox.sendKeys(Keys.END);
		nameComboBox.sendKeys(Keys.SHIFT,Keys.HOME);
		nameComboBox.sendKeys(Keys.BACK_SPACE);
		nameComboBox.sendKeys(Keys.SPACE);

		int nameComboBoxListCount = nameComboBoxList.size();

		System.err.println("nameComboBoxListCount : "+nameComboBoxListCount);

		ArrayList<String> nameComboBoxListArray = new ArrayList<String>();

		for(int i=0;i<nameComboBoxListCount;i++)
		{
			String data = nameComboBoxList.get(i).getText();

			System.err.println("Data : " + data);

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 17, 5)))
			{
				nameComboBoxList.get(i).click();
				break;
			}
		}

		nameComboBox.sendKeys(Keys.TAB);


		String actnameComboBox = nameComboBox.getAttribute("value");
		String expnameComboBox = excelReader.getCellData(xlSheetName, 17, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actnameComboBox);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(captionTxt));

		String actcaptionTxt = captionTxt.getAttribute("value");
		String expcaptionTxt = excelReader.getCellData(xlSheetName, 18, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actcaptionTxt);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cmDeleteBtn));
		cmDeleteBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cmGetMgsonDelete));

		String actMessageOnClickOnDeleteButton = cmGetMgsonDelete.getText();

		String expMessageOnClickOnDeleteButton = excelReader.getCellData(xlSheetName, 19, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actMessageOnClickOnDeleteButton);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cmclickOnOkInDelete));
		cmclickOnOkInDelete.click();

		String expMessageOnClickingOkInDelete = excelReader.getCellData(xlSheetName, 20, 6);

		String actMessageOnClickingOkInDelete = checkValidationMessageString(expMessageOnClickingOkInDelete);
		
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actMessageOnClickingOkInDelete);

		System.out.println("MessageOnClickOnDeleteButton Actual : " + actMessageOnClickOnDeleteButton + " Value Expected : " + expMessageOnClickOnDeleteButton);
		System.out.println("MessageOnClickingOkInDelete  Actual : " + actMessageOnClickingOkInDelete  + " Value Expected : " + expMessageOnClickingOkInDelete);

		if(actnameComboBox.equalsIgnoreCase(expnameComboBox) && actcaptionTxt.equalsIgnoreCase(expcaptionTxt) 
				&& actMessageOnClickOnDeleteButton.equalsIgnoreCase(expMessageOnClickOnDeleteButton)
				&& actMessageOnClickingOkInDelete.equalsIgnoreCase(expMessageOnClickingOkInDelete))
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







	public static boolean checkSavingMasterForRecordCreation() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{	
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameComboBox));
		nameComboBox.click();
		nameComboBox.sendKeys(excelReader.getCellData(xlSheetName, 21, 5));
		nameComboBox.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenuExpandBtn));
		homeMenuExpandBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterMenuBtn));
		masterMenuBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameComboBox));
		nameComboBox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(captionTxt));
		captionTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cmSaveBtn));
		cmSaveBtn.click();

		String expMessageOnSave = excelReader.getCellData(xlSheetName, 21, 6);

		String actMessageOnSave = checkValidationMessageString(expMessageOnSave);
		
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actMessageOnSave);

		if(actMessageOnSave.equalsIgnoreCase(expMessageOnSave))
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




	@FindBy(xpath="//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]/i")
	private static WebElement newClsBtn;



	public static boolean checkSavedMaserInMastersListAndSaveRecord() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		Thread.sleep(3000);

		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", ConfigMasterInMastersList);
		//Thread.sleep(2000);
		
		//getAction().moveToElement(ConfigMasterInMastersList).build().perform();
		//boolean actConfigMaster = ConfigMasterInMastersList.isDisplayed();
		//boolean expConfigMaster = true;
		
		String actConfigMaster = Boolean.toString(ConfigMasterInMastersList.isDisplayed());
		String expConfigMaster =  excelReader.getCellData(xlSheetName, 22, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actConfigMaster);

		System.out.println("actConfigMaster "+actConfigMaster);
		System.out.println("expConfigMaster "+expConfigMaster);

		if(actConfigMaster.equalsIgnoreCase( expConfigMaster ))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ConfigMasterInMastersList));
			ConfigMasterInMastersList.click();
		}

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		masterNewBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NameTxt));
		NameTxt.sendKeys(excelReader.getCellData(xlSheetName, 22, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 23, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newClsBtn));
		newClsBtn.click();

		String expSaveMessage = excelReader.getCellData(xlSheetName, 23, 6);
		String actSaveMessage = checkValidationMessageString(expSaveMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actSaveMessage);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if (actSaveMessage.equalsIgnoreCase(expSaveMessage)) 
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



	



	public static boolean checkAddingGroupMasterInConfigureMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
		Thread.sleep(2000);
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,300)","");
		Thread.sleep(2000);
		
		JavascriptExecutor js = (JavascriptExecutor) getDriver(); 
		js.executeScript("arguments[0].scrollIntoView(true);", settingsConfigureMasters);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsConfigureMasters));
		settingsConfigureMasters.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addGroupBtnInconfigureMaster));
		addGroupBtnInconfigureMaster.click();

		String expMessageOnWithoutMenu =  excelReader.getCellData(xlSheetName, 24, 6);
		String actMessageOnWithoutMenu = checkValidationMessageString(expMessageOnWithoutMenu);
		
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actMessageOnWithoutMenu);


		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenuExpandBtn));
		homeMenuExpandBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterMenuBtn));
		masterMenuBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addGroupBtnInconfigureMaster));
		addGroupBtnInconfigureMaster.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(groupNameTxt));
		groupNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 25, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(okBtnInAddgroupPopup));
		okBtnInAddgroupPopup.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenuExpandBtn));
		homeMenuExpandBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterMenuExpandBtn));
		masterMenuExpandBtn.click();

		Thread.sleep(1000);
		getAction().moveToElement(GroupMasterMenuBtn).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupMasterMenuBtn));
		String actGroup = GroupMasterMenuBtn.getText();
		String expGroup = excelReader.getCellData(xlSheetName, 25, 6); 
		
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actGroup);

		System.out.println("Added Group : "+actGroup+"  Value Expected  "+expGroup);

		if (actMessageOnWithoutMenu.equalsIgnoreCase(expMessageOnWithoutMenu) && actGroup.equalsIgnoreCase(expGroup)) 
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






	public static boolean checkAddingChildMasterInGroupMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameComboBox));
		nameComboBox.click();
		nameComboBox.sendKeys(excelReader.getCellData(xlSheetName, 26, 5));
		Thread.sleep(2000);
		nameComboBox.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupMasterMenuBtn));
		GroupMasterMenuBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cmSaveBtn));
		cmSaveBtn.click();

		String expMessageOnSave = excelReader.getCellData(xlSheetName, 26, 6);

		String actMessageOnSave = checkValidationMessageString(expMessageOnSave);
		
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actMessageOnSave);

		if(actMessageOnSave.equalsIgnoreCase(expMessageOnSave))
		{	
			
			excelReader.setCellData(xlfile, xlSheetName, 26, 8, resPass);
			return true;
		}
		else
		{
			
			excelReader.setCellData(xlfile, xlSheetName, 26, 8, resFail);
			return false;
		}
	}


	
	
	


	public static boolean checkDeletingSavedGroupmasterWithChildmaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenuExpandBtn));
		homeMenuExpandBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterMenuExpandBtn));
		masterMenuExpandBtn.click();

		Thread.sleep(2000);
		getAction().moveToElement(GroupMasterBtn).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupMasterBtn));
		GroupMasterBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteGroupBtnInconfigureMaster));
		deleteGroupBtnInconfigureMaster.click();
		Thread.sleep(2000);

		String expDeleteMessage = excelReader.getCellData(xlSheetName, 27, 6);
		String actDeleteMessage = checkValidationMessageString(expDeleteMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actDeleteMessage);

		if (actDeleteMessage.equalsIgnoreCase(expDeleteMessage))
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






	public static boolean checkAddingGroupMasterForDeletion() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
		Thread.sleep(2000);
		
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,300)","");
		Thread.sleep(2000);
		
		getAction().moveToElement(settingsConfigureMasters).build().perform();
		
		Thread.sleep(1500);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsConfigureMasters));
		settingsConfigureMasters.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenuExpandBtn));
		homeMenuExpandBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterMenuBtn));
		masterMenuBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addGroupBtnInconfigureMaster));
		addGroupBtnInconfigureMaster.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(groupNameTxt));
		groupNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 28, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(okBtnInAddgroupPopup));
		okBtnInAddgroupPopup.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenuExpandBtn));
		homeMenuExpandBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterMenuExpandBtn));
		masterMenuExpandBtn.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupMasterToDeleteMenuBtn));
		GroupMasterToDeleteMenuBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteGroupBtnInconfigureMaster));
		deleteGroupBtnInconfigureMaster.click();

		String expDeleteMessage = excelReader.getCellData(xlSheetName, 28, 6);
		String actDeleteMessage = checkValidationMessageString(expDeleteMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actDeleteMessage);


		if (actDeleteMessage.equalsIgnoreCase(expDeleteMessage))
		{
			
			 excelReader.setCellData(xlfile, xlSheetName, 28, 8, resPass);
			return true;
		}
		else
		{
			
			 excelReader.setCellData(xlfile, xlSheetName, 28, 8, resFail);
			return false;
		} 
	}








	public ConfigureMasterPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
}
