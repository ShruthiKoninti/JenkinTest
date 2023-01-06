package com.focus.Pages;

import java.io.IOException;
import java.util.ArrayList;
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
import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class BinsPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "BinsPage";

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


	
//	@FindBy(xpath="//*[@id='1115']")
	@FindBy(xpath="//*[@id='navigation_menu']/li[1]/ul/li[4]/ul/li[19]")
	public static WebElement binElement;

	public  boolean checkBinsMasterInMasterMenu() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		System.out.println("***********checkBinsMasterInBinsMenu Method Excutes **********");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		Thread.sleep(3000);
		
		JavascriptExecutor js = (JavascriptExecutor) getDriver(); 
		js.executeScript("arguments[0].scrollIntoView(true);",binElement);
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binElement));
		binElement.click();
		
		/*JavascriptExecutor js = (JavascriptExecutor) getDriver(); 
		js.executeScript("arguments[0].scrollIntoView(true);",binElement);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binElement));
		getAction().moveToElement(binElement).build().perform();
		binElement.click();
		*/
		Thread.sleep(2000);
		
		/*int count = masterSubMenusList.size();
		
		for (int i = 1; i < count; i++) 
		{
			String data = masterSubMenusList.get(i).getText();
			System.err.println(data);
			
			JavascriptExecutor js = (JavascriptExecutor) getDriver(); 
			//js.executeScript("arguments[0].scrollIntoView(true);",masterSubMenusList.get(i));
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 8, 5))) 
			{
					
				masterSubMenusList.get(i).click();
				break;
			}
			js.executeScript("arguments[0].scrollIntoView(true);",masterSubMenusList.get(i));

		}
		
		Thread.sleep(2000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		
		String expected = excelReader.getCellData(xlSheetName, 8, 6);
    	
    	String actual = Boolean.toString(masterNewBtn.isDisplayed());
    	
    	System.out.println("New Button is Displayed : "+actual+"  "+expected);
    	
    	excelReader.setCellData(xlfile, xlSheetName, 8, 7, actual);
    	

    	if(actual.equalsIgnoreCase(expected))
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





	public static boolean checkSavingBinsInBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		System.out.println("********************************checkSavingBinsInBinsMaster Method Excutes **********");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		masterNewBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
		nameTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 10, 5));
		codeTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binColNoTxt));
		binColNoTxt.click();
		binColNoTxt.clear();
		binColNoTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		binColNoTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 9, 6);

		String actMessage=checkValidationMessage(expMessage);

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();

		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		
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




	public static boolean checkEditingInSavedBinsINBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		int count = masterItemUnitsList.size();

		for (int i = 0; i < count; i++) 
		{
			String actbins = masterItemUnitsList.get(i).getText();

			if (actbins.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 12, 5))) 
			{
				if (masterItemUnitsChkboxList.get(i).isSelected()==false) 
				{
					masterItemUnitsChkboxList.get(i).click();
				}
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterEditBtn));
		masterEditBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		nameTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 14, 5));
		codeTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 12, 6);

		String actMessage=checkValidationMessage(expMessage);

		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actMessage);

		if(actMessage.equalsIgnoreCase(expMessage))
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




	public static boolean checkUpdatedBinsMasterDisplayInBins() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		int count = masterItemUnitsList.size();

		ArrayList<String> actBinsList = new ArrayList<String>();

		for (int i = 0; i < count; i++) 
		{
			String bins = masterItemUnitsList.get(i).getText();

			actBinsList.add(bins);
		}

		ArrayList<String> expBinsList = new ArrayList<String>();
		expBinsList.add(excelReader.getCellData(xlSheetName, 15, 6));

		String actBins = actBinsList.toString();
		String expBins = expBinsList.toString();

		System.out.println("************************checkUpdatedBinsMasterDisplayInBins***********************");

		System.out.println("Updated Bin : "+actBins+"  Value Expected  "+expBins);
		
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actBins);

		if(actBins.equalsIgnoreCase(expBins))
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



	public static boolean checkDeleteInBinsInBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		System.out.println("************************checkDeleteInBinsInBinsMaster  Method EXCUTES ***************");

		int count = masterItemUnitsList.size();

		for (int i = 0; i < count; i++) 
		{
			String actbins = masterItemUnitsList.get(i).getText();

			if (actbins.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 16, 5))) 
			{
				if (masterItemUnitsChkboxList.get(i).isSelected()==false) 
				{
					masterItemUnitsChkboxList.get(i).click();
				}
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));

		masterDeleteBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgOnDelete));
		System.out.println(getMsgOnDelete.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
		clickOnOkInDelete.click();

		String expMessage=excelReader.getCellData(xlSheetName, 16, 6);

		String actMessage=checkValidationMessage(expMessage);

		System.out.println("********************************checkDeleteItemInItemMaster***************************");

		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if(actMessage.equalsIgnoreCase(expMessage))
		{
			excelReader.setCellData(xlfile, xlSheetName, 16, 8, resPass);

			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 16, 8, resFail);
			return false;
		}	 
	}


	@FindBy(xpath="//*[@id='drpdownSlider']")
	public static WebElement toggleRibbon;

	@FindBy(xpath="//*[@id='btnAdvImportForMaster']")
	public static WebElement advanceMasterImportExportBtn;;

	@FindBy(xpath="//*[@id='viewheading']")
	public static WebElement advanceMasterLabel;
					 
	@FindBy(xpath="//button[@id='btnFile']")
	public static WebElement FileUploadBtn;
		
	@FindBy(xpath="//*[@class='icon-font7 icon-close']")
	public static WebElement advMasterCloseBtn;


	@FindBy(xpath="//*[@id='btnImportDataForAdvanceMasterPopUp']")
	public static WebElement advMasterImportDataBtn;;


	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[10]")
	private static List<WebElement> binNameList;

		
	public boolean checkBinsListBeforeImportingDataFromExcel() throws InterruptedException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		   	homeMenu.click();
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		   	mastersMenu.click();
		   	
		   	
			JavascriptExecutor js = (JavascriptExecutor) getDriver(); 
			js.executeScript("arguments[0].scrollIntoView(true);",binsMenu);
			Thread.sleep(1000);
			
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binsMenu));
			     
		   	binsMenu.click();
			   
		   	Thread.sleep(2000);
		   	
		
		 
		 int binsCount = binNameList.size();
		

		 ArrayList<String> binlist = new ArrayList<String>();

		 		for (int i = 0; i < binsCount; i++)
		 		{
		 			 String data = binNameList.get(i).getText();
		 			binlist.add(data);
		 			
		 		}
		 		
		 		String actBinList=binlist.toString();
		 		String expBinList="[BSTB, Bin3, BG1, BG2]";
		 		
		 		System.out.println("Actual BinsList:  "+actBinList);
		 		System.out.println("Expected Bins List:  "+expBinList);
		 		
		 		if(actBinList.equalsIgnoreCase(expBinList))
		 		{
		 			return true;
		 		}
		 		else
		 		{
		 			return false;
		 		}
		 		
	}

	public boolean checkAdvanceMasterImportExportBins() throws InterruptedException
	{
		Thread.sleep(2000);
		
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
	   	toggleRibbon.click();
	   	Thread.sleep(2000);
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
	   	advanceMasterImportExportBtn.click();
	   	
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterLabel));
	    System.out.println("Advance Master import/Export Label  "+advanceMasterLabel.getText());
	      	
	  	if(advanceMasterLabel.getText().equalsIgnoreCase("Advance Master Import/Export----> Bins"))
		{	
			return true;
		}	 
		else
		{
			return false;
		}
	   	
	   	
	}
	public boolean checkImportDatainAdvanceMasterImportExportBins() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
	{
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterImportDataBtn));
		 advMasterImportDataBtn.click();
		 Thread.sleep(2000);
		 String expMessage="Please load excel sheet";
		String actMessage=checkValidationMessage(expMessage);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
		FileUploadBtn.click();
		Thread.sleep(2000);
		
		 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BinsImportFromExcel.exe");
		 Thread.sleep(5000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterImportDataBtn));
		 advMasterImportDataBtn.click();
		 Thread.sleep(2000);
		 
		/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterCloseBtn));
		 advMasterCloseBtn.click();
		 Thread.sleep(2000);*/
		 
		 String expMessage1="10 Records Imported Successfully";
		 String actMessage1=checkValidationMessage(expMessage);
		 
		/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterCloseBtn));
		 advMasterCloseBtn.click();
		 Thread.sleep(2000);*/
		 
		 if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
		 {
			 return true;
		 }
		 else
		 {
			 return false;
		 }
		
		}


	public boolean checkBinsListAfterImportingDataFromExcel() throws InterruptedException
	{
		Thread.sleep(2000);
		 	 
		 int binsCount = binNameList.size();
			

		 ArrayList<String> binlist = new ArrayList<String>();

		 		for (int i = 0; i < binsCount; i++)
		 		{
		 			 String data = binNameList.get(i).getText();
		 			binlist.add(data);
		 			
		 		}
		 		
		 		String actBinList=binlist.toString();
		 		String expBinList="[BSTB, Bin3, BG1, BG2, Bin7, Bin8]";
		 		
		 		System.out.println("Actual Bins List After Import:  "+actBinList);
		 		System.out.println("Expected Bins List After Import:  "+expBinList);
		 		
		 		if(actBinList.equalsIgnoreCase(expBinList))
		 		{
		 			return true;
		 		}
		 		else
		 		{
		 			return false;
		 		}
		 		
	}


	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[8]//input")
	private static List<WebElement> masterGridBodyChkbox;

	public boolean checkImportedBinDelete() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 int count = binNameList.size();
			
			for (int i = 0; i < count; i++) 
			{
				String data = binNameList.get(i).getText();
				
				if (data.equalsIgnoreCase("Bin7")) 
				{
					if (masterGridBodyChkbox.get(i).isSelected()==false) 
					{
						masterGridBodyChkbox.get(i).click();
						break;
					}
				}
			}
		    
		    Thread.sleep(2000);	
		 
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));

	  masterDeleteBtn.click();
	  
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgOnDelete));
	  System.out.println(getMsgOnDelete.getText());
	  
	  //Thread.sleep(2000);
	  
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
	  clickOnOkInDelete.click();
	  
	  String expMessage="Record Deleted Successfully.";
		
	  String actMessage=checkValidationMessage(expMessage);
	  
	  System.out.println("********************************checkDeleteAccountBeforeCustomization***************************");
			
		 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			
		 if(actMessage.equalsIgnoreCase(expMessage))
		 {
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			 closeBtn.click();
			 return true;
		 }
		 
		 else
		 {
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			 closeBtn.click();
			 return false;
		 }
	}	











	public BinsPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
