package com.focus.Pages;

import static org.testng.Assert.assertTrue;

import java.awt.AWTException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
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

public class AccountsPage extends BaseEngine
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "AccountsPage";
	
	
	
	
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
		
	
	
	
	
	public static boolean checkSavingAccountINAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		
		try 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		    homeMenu.click();
			Thread.sleep(1500);
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		    mastersMenu.click();
		    Thread.sleep(1500);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
		    accounts.click();
		    Thread.sleep(8000);
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		    masterNewBtn.click();
		    Thread.sleep(6000);
			System.out.println("********************************checkSavingAccountINAccountMaster*****************************");
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
			nameTxt.sendKeys(Keys.TAB);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			codeTxt.click();
			codeTxt.clear();
			codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 9, 5));
			codeTxt.sendKeys(Keys.TAB);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountTypeDropdown));
			accountTypeDropdown.sendKeys(Keys.TAB);
			 
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			 
			Thread.sleep(2000);
			String expMessage=excelReader.getCellData(xlSheetName, 8, 6);
				
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);
			 
			System.out.println("actMessage     : "+actMessage);
			System.out.println("expMessage     : "+expMessage);
			
			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
			masterCloseBtn.click();
			
			Thread.sleep(4000);
				
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
			
		} catch (Exception e) 
		{
			String Exception=e.getMessage();
			System.err.println(Exception);
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, Exception);
			return false;
		}
	}
	
	
	
	
	
	public static boolean checkEditingInSavedAccountINAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		 System.out.println("********************************checkEditingInSavedAccountINAccountMaster********************");
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountFirstCheckBox));
			
		 int count = masterGridBodyName.size();
			
			for (int i = 0; i < count; i++) 
			{
				String data = masterGridBodyName.get(i).getText();
				
				if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 10, 5))) 
				{
					getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
					break;
				}
			}
		    
		    Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 11, 5));
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		 codeTxt.sendKeys(Keys.TAB);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountTypeDropdown));
		 accountTypeDropdown.sendKeys(Keys.TAB);
         
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 Thread.sleep(2000);
		 
		 String expMessage = excelReader.getCellData(xlSheetName, 10, 6);
			
		 String actMessage=checkValidationMessage(expMessage);
		 
		 excelReader.setCellData(xlfile, xlSheetName, 10, 7, actMessage);
		 
		 System.out.println("actMessage     : "+actMessage);
		 System.out.println("expMessage     : "+expMessage);
		 
		 if(actMessage.equalsIgnoreCase(expMessage))
		 {
			 excelReader.setCellData(xlfile, xlSheetName, 10, 8, resPass);
			 return true;
		 }
		 
		 else
		 {
			 excelReader.setCellData(xlfile, xlSheetName, 10, 8, resFail);
			 return false;
		 }
	}
	
	
	
	
	
	
	
	public static boolean checkUpdatedAccountDisplayInAccountGridBeforeCustomization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		String actaccountNewCreationName  = null;
		
		 int count = masterItemGridBodyName.size();
			
			for (int i = 0; i < count; i++) 
			{
				String data = masterItemGridBodyName.get(i).getText();
				
				if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 13, 5))) 
				{
					actaccountNewCreationName = data;
					break;
				}
			}
			    
			 Thread.sleep(4000);	
		 String expaccountNewCreationName  =excelReader.getCellData(xlSheetName, 13, 6);
		 
		 System.out.println("************************checkUpdatedAccountDisplayInAccountMasterGridBeforeCustomization***********************");
		 
		 System.err.println("actaccountNewCreationName  : "+actaccountNewCreationName);
		 System.err.println("expaccountNewCreationName  : "+expaccountNewCreationName);
		 
		 
		 excelReader.setCellData(xlfile, xlSheetName, 13, 7, actaccountNewCreationName);
		 
		 if(actaccountNewCreationName.equalsIgnoreCase(expaccountNewCreationName))
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
	
	
	
	
	
	
	
	
	
	public static boolean checkDeleteAccountInAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 
		 Thread.sleep(6000);	
		 
		 int count = masterGridBodyName.size();
			
			for (int i = 0; i < count; i++) 
			{
				String data = masterGridBodyName.get(i).getText();
				
				System.err.println("*"+data+"*");
				
				if (data.equalsIgnoreCase(/*excelReader.getCellData(xlSheetName, 14, 5)*/"Trial Account Update") )
				{
					if (masterGridBodyChkbox.get(i).isSelected()==false) 
					{
						masterGridBodyChkbox.get(i).click();
						break;
					}
				}
			}
		    
		    Thread.sleep(6000);	
		 
         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));
         masterDeleteBtn.click();
         
         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgOnDelete));
         System.out.println(getMsgOnDelete.getText());
         
         Thread.sleep(2000);
         
         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
         clickOnOkInDelete.click();
         
         String expMessage=excelReader.getCellData(xlSheetName, 14, 6);
 		
         String actMessage=checkValidationMessage(expMessage);
         
         excelReader.setCellData(xlfile, xlSheetName, 14, 7, actMessage);
         
         System.out.println("********************************checkDeleteAccountBeforeCustomization***************************");
			
		 System.out.println("actMessage  : "+actMessage);
		 System.out.println("expMessage  : "+expMessage);
		 
			
		 if(actMessage.equalsIgnoreCase(expMessage))
		 {
			/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			 closeBtn.click();*/
			 excelReader.setCellData(xlfile, xlSheetName, 14, 8, resPass);
			 return true;
		 }
		 
		 else
		 {
			/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			 closeBtn.click();*/
			 excelReader.setCellData(xlfile, xlSheetName, 14, 8, resFail);
			 return false;
		 }
	}
	
	
	///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	
	@FindBy(xpath="//*[@id='drpdownSlider']")
	public static WebElement toggleRibbon;

	@FindBy(xpath="//*[@id='btnAdvImportForMaster']")
	public static WebElement advanceMasterImportExportBtn;;

	@FindBy(xpath="//*[contains(text(),'Advance Master Import/Export')]")
	public static WebElement advanceMasterLabel;
					 
	@FindBy(xpath="//button[@id='btnFile']")
	public static WebElement FileUploadBtn;
		

	@FindBy(xpath="(//*[@class='icon-close'])[1]")
	public static WebElement advMasterCloseBtn;

	@FindBy(xpath="//*[@id='btnImportDataForAdvanceMasterPopUp']")
	public static WebElement advMasterImportDataBtn;;


	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[11]")
	private static List<WebElement> accNameList;

	
	public String actAccList=null;
	public String expAccList=null;
	
		
	public boolean checkAccountsListBeforeImportingDataFromExcel() throws InterruptedException
	{
		/*getDriver().navigate().refresh();
		Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		   	homeMenu.click();
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		   	mastersMenu.click();
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
			     
		   	accounts.click();*/
			   
		   	Thread.sleep(2000);
		   	
		
		 
		 int accCount = accNameList.size();
		

		 ArrayList<String> acclist = new ArrayList<String>();

		 		for (int i = 0; i < accCount; i++)
		 		{
		 			 String data = accNameList.get(i).getText();
		 			acclist.add(data);
		 			
		 		}
		 		
		 		 actAccList=acclist.toString();
		 		 expAccList="[Pricebook Vendor Account Two, Pricebook Vendor Account One, Pricebook Account One, Pricebook Account Two, Pricebook Group Two, Pricebook Group One, Round off Exchange gain / loss, ASSETS, EXPENSES, CONTROL ACCOUNTS, REVENUE, EQUITIES, LIABILITIES, COGS POSTING ACC, BR COGS ACC INV, FIFO COGS ACC INV, WA COGS ACC INV, STD RATE COGS ACC INV, SR COGS POSTING ACC, SHORTAGE COGS POSTING ACC, EXCESS COGS POSTING ACC, VAT OUTPUT, VAT ADVANCE SALE, VAT ADVANCE PURCHASE, PURCHASE VARIANCE, VAT INPUT]";
		 		
		 		System.out.println("Actual Accounts List:  "+actAccList);
		 		System.out.println("Expected Accounts List:  "+expAccList);
		 		
		 		if(actAccList.equalsIgnoreCase(expAccList))
		 		{
		 			return true;
		 		}
		 		else
		 		{
		 			return false;
		 		}
		 		
	}

	public boolean checkAdvanceMasterImportExportAccount() throws InterruptedException
	{
		Thread.sleep(2000);
		
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
	   	toggleRibbon.click();
	   	Thread.sleep(2000);
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
	   	advanceMasterImportExportBtn.click();
	   	
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterLabel));
	    System.out.println("Advance Master import/Export Label  "+advanceMasterLabel.getText());
	      	
	  	if(advanceMasterLabel.getText().equalsIgnoreCase("Advance Master Import/Export"))
		{	
			return true;
		}	 
		else
		{
			return false;
		}
	   	
	   	
	}
	public boolean checkImportDatainAdvanceMasterImportExportAccount() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
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
		
		 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\AccountImportFromExcel.exe");
		 Thread.sleep(25000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterImportDataBtn));
		 advMasterImportDataBtn.click();
		 Thread.sleep(15000);
		 
		/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterCloseBtn));
		 advMasterCloseBtn.click();
		 Thread.sleep(3000);*/
		 
		 String expMessage1="110 Records Imported Successfully";
		 String actMessage1=checkValidationMessage(expMessage);
		 
	
		 if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1))
		 {
			 return true;
		 }
		 else
		 {
			 return false;
		 }
		
		}

	@FindBy(xpath="//*[@id='LandingGridBody']/tr[28]/td")
	public static List<WebElement> accounts28thRowList;

	public boolean checkAccountsListAfterImportingDataFromExcel() throws InterruptedException
	{
		Thread.sleep(4000);
		
		
		 int accCount = accNameList.size();
			

		 ArrayList<String> acclist = new ArrayList<String>();

		 		for (int i = 0; i < accCount; i++)
		 		{
		 			 String data = accNameList.get(i).getText();
		 			acclist.add(data);
		 			
		 		}
		 		
		 		 actAccList=acclist.toString();
		 		 expAccList="[Pricebook Vendor Account Two, Pricebook Vendor Account One, Pricebook Account One, Pricebook Account Two, Pricebook Group Two, Pricebook Group One, SALES, Round off Exchange gain / loss, ASSETS, EXPENSES, CONTROL ACCOUNTS, REVENUE, EQUITIES, LIABILITIES, COGS POSTING ACC, BR COGS ACC INV, FIFO COGS ACC INV, WA COGS ACC INV, STD RATE COGS ACC INV, SR COGS POSTING ACC, SHORTAGE COGS POSTING ACC, EXCESS COGS POSTING ACC, VAT OUTPUT, VAT ADVANCE SALE, VAT ADVANCE PURCHASE, PURCHASE VARIANCE, VAT INPUT, AccountImport]";
		 		
		 		System.out.println("Actual Accounts List:  "+actAccList);
		 		System.out.println("Expected Accounts List:  "+expAccList);
		 		
		 		if(actAccList.equalsIgnoreCase(expAccList))
		 		{
		 			return true;
		 		}
		 		else
		 		{
		 			return false;
		 		}
		 		
		 		
		 		
		 		
	}
public boolean checkImportedAccountDelete() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	 int count = masterGridBodyName.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = masterGridBodyName.get(i).getText();
			Thread.sleep(2000);
			
			if (data.equalsIgnoreCase("AccountImport")) 
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
  Thread.sleep(2000);
  
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



/*@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[12]")
public static List<WebElement> masterGridBodyName;*/


	

public boolean checkImportedGroupAccountDetails() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	 int count = masterGridBodyName.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = masterGridBodyName.get(i).getText();
			
			if (data.equalsIgnoreCase("GroupAccountImport")) 
			{
				if (masterGridBodyChkbox.get(i).isSelected()==false) 
				{
					masterGridBodyChkbox.get(i).click();
					getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
					break;
				}
			}
		}
	    
	    Thread.sleep(2000);	
	    

		 int accCount = accNameList.size();
			

		 ArrayList<String> acclist = new ArrayList<String>();

		 		for (int i = 0; i < accCount; i++)
		 		{
		 			 String data = accNameList.get(i).getText();
		 			acclist.add(data);
		 			
		 		}
		 		
		 		 actAccList=acclist.toString();
		 		 expAccList="[A1]";
		 		
		 		System.out.println("Actual Accounts List:  "+actAccList);
		 		System.out.println("Expected Accounts List:  "+expAccList);
		 		
		 		if(actAccList.equalsIgnoreCase(expAccList))
		 		{
		 			return true;
		 		}
		 		else
		 		{
		 			return false;
		 		}

}
	
	
/////////////////////Adding Groups and subgroups//////////////////////////

public boolean checkSavingAccountGroupinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
		accounts.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountList));
		accountList.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
		masterAddGroupBtn.click();
		Thread.sleep(2000);
		System.out.println("********************************checkSavingAccountGroupAccountMaster*****************************");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys("Group Account");
		nameTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys("GroupAccount");
		codeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountTypeDropdown));
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		String expMessage="Saved Successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("actMessage  : "+actMessage);
		System.out.println("expMessage  : "+actMessage);
		
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{
		return true;
		}
		else
		{
		return false;
		}
		
		
}


		@FindBy(xpath="//*[@id='0']/a")
		public static WebElement accountList;
		
		
		
		@FindBy(xpath="(//*[@id='0']/li)[1]/a[1]")
		public static WebElement accountGrp;
		
		
		@FindBy(xpath="//*[@id='lblnchecked']/input")
		public static List<WebElement> accChkBoxList;
		
		@FindBy(xpath="//*[@id='ol_treeNavigation']/li[2]")
		public static WebElement GroupAccountLabel;

public static boolean checkSavingAccountSubGroupinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
		int accNameListCount = accNameList.size();
		
		for(int i=0;i<accNameListCount;i++)
		{
		String data = accNameList.get(i).getText();
		Thread.sleep(2000);
		if(data.equals("Group Account"))
		{
		Thread.sleep(2000);
		getAction().doubleClick(accChkBoxList.get(i)).build().perform();
		Thread.sleep(2000);
		break;
		}
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
		masterAddGroupBtn.click();
		Thread.sleep(8000);
		System.out.println("********************************checkSavingAccountSubGroupAccountMaster*****************************");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys("SubGroup1 Account");
		nameTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys("SubGroup1Account");
		codeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountTypeDropdown));
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(8000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys("SubGroup2 Account");
		nameTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys("SubGroup2Account");
		codeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountTypeDropdown));
		accountTypeDropdown.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		Thread.sleep(6000);
		
		String expMessage="Saved Successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("actMessage  : "+actMessage);
		System.out.println("expMessage  : "+expMessage);
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{
		return true;
		}
		else
		{
		return false;
		}
		
}

	public boolean checkCreatedGroupDisplayedonLeftPanel()
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountGrp));
		
		boolean ActGrpAccDisplayed=accountGrp.isDisplayed();
		boolean ExpGrpAccDisplayed=true;
		
		System.out.println("ActGrpAccDisplayed  :"+ ActGrpAccDisplayed);
		System.out.println("ExpGrpAccDisplayed  :"+ ExpGrpAccDisplayed);
		
		if(ActGrpAccDisplayed==ExpGrpAccDisplayed)
		{
		return true;
		}
		
		else
		{
		return false;
		}
		
	}


	public boolean checkSavingAccountsinSubGroupsinAccountMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
			
			int accNameListCount = accNameList.size();
			
			for(int i=0;i<accNameListCount;i++)
			{
			String data = accNameList.get(i).getText();
			Thread.sleep(2000);
			if(data.equals("SubGroup1 Account"))
			{
			Thread.sleep(2000);
			getAction().doubleClick(accChkBoxList.get(i)).build().perform();
			Thread.sleep(2000);
			break;
			}
			}
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
			masterNewBtn.click();
			Thread.sleep(4000);
			
			
			System.out.println("********************************checkSavingAccountSubGroupAccountMaster*****************************");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys("Account1");
			nameTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			codeTxt.click();
			codeTxt.clear();
			codeTxt.sendKeys("Account1");
			codeTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountTypeDropdown));
			accountTypeDropdown.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(2000);
			
			String expMessage1="Saved Successfully";
			
			String actMessage1=checkValidationMessage(expMessage1);
			
			System.out.println("actMessage 1 : "+actMessage1);
			System.out.println("expMessage 1 : "+expMessage1);
			
			Thread.sleep(8000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
			masterCloseBtn.click();
			Thread.sleep(6000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupAccountLabel));
			GroupAccountLabel.click();
			Thread.sleep(2000);
			
			int accNameListCount1 = accNameList.size();
			
			for(int i=0;i<accNameListCount1;i++)
			{
			String data = accNameList.get(i).getText();
			Thread.sleep(2000);
			if(data.equals("SubGroup2 Account"))
			{
			Thread.sleep(2000);
			getAction().doubleClick(accChkBoxList.get(i)).build().perform();
			Thread.sleep(2000);
			break;
			}
			}
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
			masterNewBtn.click();
			
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys("Account2");
			nameTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			codeTxt.click();
			codeTxt.clear();
			codeTxt.sendKeys("Account2");
			codeTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountTypeDropdown));
			accountTypeDropdown.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			Thread.sleep(5000);
			String expMessage2="Saved Successfully";
			
			String actMessage2=checkValidationMessage(expMessage2);
			
			System.out.println("actMessage 2 : "+actMessage2);
			System.out.println("expMessage 2 : "+expMessage2);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
			masterCloseBtn.click();
			Thread.sleep(1000);
			
			if(actMessage1.equalsIgnoreCase(expMessage1) && actMessage2.equalsIgnoreCase(expMessage2))
			{
			return true;
			}
			else
			{
			return false;
			}
}

	
	@FindBy(xpath="//*[@id='ol_treeNavigation']/li[1]")
	public static WebElement accountLabel;

	public static boolean checkAccountsSavedinGroupsandSubGroupsinAccountMaster() throws InterruptedException
	{
		Thread.sleep(4000);
		getDriver().navigate().refresh();
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
		
		accounts.click();
		
		Thread.sleep(6000);
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountLabel));
	   	accountLabel.click();
		Thread.sleep(2000);*/
		
		int accCount = accNameList.size();
		
		ArrayList<String> acclist = new ArrayList<String>();
		
		for (int i = 0; i < accCount; i++)
		{
		String data = accNameList.get(i).getText();
		acclist.add(data);
		
		}
		
		String actAccList=acclist.toString();
		String  expAccList="[CurrencyBank, CurrencyVendor, CurrencyCustomer, Pricebook Vendor Account Two, Pricebook Vendor Account One, Pricebook Account One, Pricebook Account Two, Pricebook Group Two, Pricebook Group One, Group Account, Round off Exchange gain / loss, ASSETS, EXPENSES, CONTROL ACCOUNTS, REVENUE, EQUITIES, LIABILITIES, COGS POSTING ACC, BR COGS ACC INV, FIFO COGS ACC INV, WA COGS ACC INV, STD RATE COGS ACC INV, SR COGS POSTING ACC, SHORTAGE COGS POSTING ACC, EXCESS COGS POSTING ACC, VAT OUTPUT, VAT ADVANCE SALE, VAT ADVANCE PURCHASE, PURCHASE VARIANCE, VAT INPUT]";
		
		
		int accNameListCount = accNameList.size();
		
		for(int i=0;i<accNameListCount;i++)
		{
		String data = accNameList.get(i).getText();
		Thread.sleep(2000);
		if(data.equals("Group Account"))
		{
			Thread.sleep(2000);
			getAction().doubleClick(accChkBoxList.get(i)).build().perform();
			Thread.sleep(2000);
			break;
		}
		}
		
		
		int accCount1= accNameList.size();
		ArrayList<String> acclist1 = new ArrayList<String>();
		
		for (int i = 0; i < accCount1; i++)
		{
		String data = accNameList.get(i).getText();
		acclist1.add(data);
		}
			
		String actGrpAccList=acclist1.toString();
		String  expGrpAccList="[SubGroup2 Account, SubGroup1 Account]";
			
		
		int accNameListCount1 = accNameList.size();
		for(int i=0;i<accNameListCount1;i++)
		{
		String data = accNameList.get(i).getText();
		Thread.sleep(2000);
		if(data.equals("SubGroup1 Account"))
		{
		Thread.sleep(2000);
		getAction().doubleClick(accChkBoxList.get(i)).build().perform();
		Thread.sleep(2000);
		break;
		}
		}
		
		
		int accCount2= accNameList.size();
		ArrayList<String> acclist2 = new ArrayList<String>();
		for (int i = 0; i < accCount2; i++)
		{
		String data = accNameList.get(i).getText();
		acclist2.add(data);
		}
		
		String actSubGrpAccList=acclist2.toString();
		String  expSubGrpAccList="[Account1]";
			 		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupAccountLabel));
		GroupAccountLabel.click();
		Thread.sleep(2000);
		
		int accNameListCount3 = accNameList.size();
		for(int i=0;i<accNameListCount3;i++)
		{
		String data = accNameList.get(i).getText();
		Thread.sleep(2000);
		if(data.equals("SubGroup2 Account"))
		{
		Thread.sleep(2000);
		getAction().doubleClick(accChkBoxList.get(i)).build().perform();
		Thread.sleep(2000);
		break;
		}
		}
		
		
		int accCount3= accNameList.size();
		ArrayList<String> acclist3 = new ArrayList<String>();
		for (int i = 0; i < accCount3; i++)
		{
		String data = accNameList.get(i).getText();
		acclist3.add(data);
		}
		
		String actSubGrpAcc2List=acclist3.toString();
		String  expSubGrpAcc2List="[Account2]";
			 		
		System.out.println("********************************checkSavingGroupsandSubGroupAccountMaster*****************************");
		
		
		System.out.println("Actual Accounts List  :  "+actAccList);
		System.out.println("Expected Accounts List:  "+expAccList);
		
		System.out.println("Actual Sub Group Accounts List  :  "+actGrpAccList);
		System.out.println("Expected Sub Group Accounts List:  "+expGrpAccList);
		
		System.out.println("Actual  Accounts in Sub Group1 List :  "+actSubGrpAccList);
		System.out.println("Expected Accounts in Sub Group1 List:  "+expSubGrpAccList);
		
		System.out.println("Actual  Accounts in Sub Group2 List :  "+actSubGrpAcc2List);
		System.out.println("Expected Accounts in Sub Group2 List:  "+expSubGrpAcc2List);
		
		
		if(actAccList.contains(expAccList) && actGrpAccList.equalsIgnoreCase(expGrpAccList)
		&& actSubGrpAccList.equalsIgnoreCase(expSubGrpAccList) && actSubGrpAcc2List.equalsIgnoreCase(expSubGrpAcc2List))
		{
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();*/
		return true;
		}
		else
		{
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();*/
		return false;
		}
		
	}


	public boolean checkDeleteGroupAccountinAccountMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
		accounts.click();
		
		
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountList));
		accountList.click();
		Thread.sleep(2000);
		
		
		int count1 = masterGridBodyName.size();
		
		for (int i = 0; i < count1; i++) 
		{
		String data = masterGridBodyName.get(i).getText();
		
		if (data.equalsIgnoreCase("Group Account")) 
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
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
		clickOnOkInDelete.click();
		
		String expMessage="Record Deleted Successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("********************************checkDeleteAccountBeforeCustomization***************************");
		
		System.out.println("actMessage : "+actMessage);
		System.out.println("expMessage : "+expMessage);
		
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();*/
		return true;
		}
		
		else
		{
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();*/
		return false;
		}
		
		
	}

	



public boolean checkOpenGroupAccountMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
    homeMenu.click();
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
    mastersMenu.click();
     
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
    accounts.click();
    	  
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountsTitle));
    
  
	 
	System.out.println("********************************checkOpeningGroupAccountINAccountMaster*****************************");
	 
	int accNameListCount = accNameList.size();
	ArrayList<String>accNameListArray = new ArrayList<String>();
	for(int i=0;i<accNameListCount;i++)
	{
		String data = accNameList.get(i).getText();
		Thread.sleep(2000);
		if(data.equals("ASSETS"))
		{
			Thread.sleep(2000);
			getAction().doubleClick(accChkBoxList.get(i)).build().perform();
			Thread.sleep(2000);
			break;
		}
	}
	
	int accNameListCount1 = accNameList.size();

	for(int i=0;i<accNameListCount1;i++)
	{
		String data = accNameList.get(i).getText();
		accNameListArray.add(data);
		
	}
	

	
	String actRow1List = accNameListArray.toString();
	String expRow1List = "[Fixed Assets, Current Assets]";
	
	
	System.out.println("actRow1List : "+actRow1List);
	System.out.println("expRow1List : "+expRow1List);	
	
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
	masterCloseBtn.click();*/
		
	if(actRow1List.equalsIgnoreCase(expRow1List))
	{
		 return true;
	}
	else
	{
		 return false;
	}

	
	
	
}
	
	

@FindBy(xpath="//*[@id='cmbUserTypeMaster']")
public static WebElement comboSearchBox;

@FindBy(xpath="//*[@id='ol_treeNavigation']")
public static WebElement navigationAcc;;



public boolean checkSearchinganAccountUnderSubGroup() throws InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(comboSearchBox));
	comboSearchBox.click();
	comboSearchBox.clear();
	comboSearchBox.sendKeys("Vehicles");
	Thread.sleep(2000);
	comboSearchBox.sendKeys(Keys.TAB);
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(navigationAcc));
	System.out.println(navigationAcc.getText());
	
	
	int accNameListCount = accNameList.size();
	boolean data=false ;
	for(int i=0;i<accNameListCount;i++)
	{
		 data = accChkBoxList.get(i).isSelected();
		 
		if(data==true)
		{
			
			break;
					
		}
	}
	
	if(data==true)
	{
		return true;
	}
		else
		{
			return false;
		}
	}

public static boolean checkRestoreAutomationCompany() throws InterruptedException, IOException, AWTException
{
	//String actMessage=BaseEngine.restoreCompany("SanityBackupwithoutPronghorn","Automation Company");
	String actMessage=BaseEngine.restoreCompany("SanityBackupWithPronghorn","Automation Company");
	String expMessage="Restore company code : 010";
	
	System.err.println("Actual Text :"	+actMessage);
	System.err.println("Expected Text :"	+expMessage);
	
	if(actMessage.equalsIgnoreCase(expMessage))
	{
		return true;
	}
	else
	{
		return false;
	}
	
}




	
///////Adding Extra Fields///////////////

@FindBy(xpath="//*[@id='drpdownSlider']/a/i")
public static WebElement masterRibbonToExpandOptions; 


@FindBy(xpath="//*[@id='masterFields_list']/li[1]")
public static WebElement masterGeneralField; 

@FindBy(xpath="//*[@id='tab_HeaderDetails_0']")
public static WebElement masterGeneralHeaderDetails; 

@FindBy(xpath="(//*[@id='btnAdd_Field_MasterCust']/a/i)[1]")
public static WebElement customizeMasterCreateExtraFieldAddButton; 	

@FindBy(xpath="//*[@id='ExtraField_FieldCaption']")
public static WebElement extraFields_FieldDetailsCaption;   


@FindBy(xpath="//*[@id='btnSaveRule']/a/i")
public static WebElement extraFields_OkBtn; 


@FindBy(xpath="//*[@id='example']//tbody//tr[9]//td[4]")
public static WebElement extraFields_CreatedByUserInHeaderDetailsTab;

@FindBy(xpath="(//*[@class='icon-save hiconright2'])[1]")
public static WebElement extraFields_SaveBtn; 

@FindBy(xpath="(//*[@id='btnCloseMasterCutomizationFromMasterScreen']/a/i)[1]")
public static WebElement extraFields_CloseBtn;


public boolean checkCreatingExtraFieldsinAccountMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	
	Thread.sleep(2000);
	
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	mastersMenu.click();
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
	accounts.click();
	
	
	Thread.sleep(4000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));
	masterRibbonToExpandOptions.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCustamizemasterBtn));
	masterCustamizemasterBtn.click();
	
	Thread.sleep(6500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterGeneralField));
	masterGeneralField.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterGeneralHeaderDetails));
	masterGeneralHeaderDetails.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCreateExtraFieldAddButton));
	customizeMasterCreateExtraFieldAddButton.click();
	Thread.sleep(8000);			
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_FieldDetailsCaption));
	extraFields_FieldDetailsCaption.sendKeys("Acc Extra Field ");
	extraFields_FieldDetailsCaption.sendKeys(Keys.TAB);
         
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_OkBtn));
	extraFields_OkBtn.click();
	Thread.sleep(6000);
         
	String expMessage="Field Added Successfully.";
	
	String actMessage=checkValidationMessage(expMessage);
	
	Thread.sleep(5500);	
	
	System.out.println("actMessage  :"+actMessage);
	System.out.println("expMessage  :"+expMessage);
	
	
	boolean actextraFields_CreatedByUserInHeaderDetailsTab =extraFields_CreatedByUserInHeaderDetailsTab.isDisplayed();
	
	
	boolean expextraFields_CreatedByUserInHeaderDetailsTab =true;
	
	System.out.println("***************************checkCreatedFieldDisplayInCustomizeTab*****************************");
	
	
	System.out.println("actextraFields_CreatedByUserInHeaderDetailsTab : "+actextraFields_CreatedByUserInHeaderDetailsTab);
	System.out.println("expextraFields_CreatedByUserInHeaderDetailsTab : "+expextraFields_CreatedByUserInHeaderDetailsTab);
			 
	System.out.println("created field is :" + extraFields_CreatedByUserInHeaderDetailsTab.getText());
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_SaveBtn));
	extraFields_SaveBtn.click();
	Thread.sleep(6000);	
	
	String expMessage1="Master updated successfully";
	
	String actMessage1=checkValidationMessage(expMessage1);
	
	System.out.println("********************************checkExtraFieldCreated***************************");
	
	System.out.println("actMessage1  : "+actMessage1);
	System.out.println("expMessage1  : "+expMessage1);
	
	if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) &&actextraFields_CreatedByUserInHeaderDetailsTab==expextraFields_CreatedByUserInHeaderDetailsTab)
	{
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_CloseBtn));
		extraFields_CloseBtn.click();
		
		return true;
	}
	else
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_CloseBtn));
		extraFields_CloseBtn.click();
		return false;
	}
}
 
 

	
@FindBy(xpath="//*[@id='newMasterDiv0']//div[28]//label")
public static WebElement createdextraFieldinGeneralTab;
	
	
public boolean checkAddedExtraFieldinGeneralTabofAccountMaster() throws InterruptedException
{
	 Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
    homeMenu.click();
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
    mastersMenu.click();
     
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
    accounts.click();
    Thread.sleep(2000);
    
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
	masterNewBtn.click();
	Thread.sleep(6000);
	  
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdextraFieldinGeneralTab));
	String actCreatedExtraField=createdextraFieldinGeneralTab.getText();
	String expCreatedEXtraField="Acc Extra Field";
	  
	System.out.println("actCreatedExtraField     : "+actCreatedExtraField);
	System.out.println("expCreatedEXtraField     : "+expCreatedEXtraField);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
	masterCloseBtn.click();
			
		if(actCreatedExtraField.equalsIgnoreCase(expCreatedEXtraField))
		{
			
			
			return true;
		}
		else
		{
			
			return false;
		}
		 
}
	


@FindBy(xpath="(//*[@id='example']/tbody/tr[9]/td[1]/i[1])[1]")
public static WebElement extraFields_EditOptionInHeaderDetailsTab;

@FindBy(xpath="(//*[@id='example']/tbody/tr[9]/td[1]/i[2])[1]")
public static WebElement extraFields_DeleteOptionInHeaderDetailsTab;


public boolean checkEditandUpdateExtraFieldinCustomizeMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));
	masterRibbonToExpandOptions.click();
			 
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCustamizemasterBtn));
	masterCustamizemasterBtn.click();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterGeneralField));
	masterGeneralField.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterGeneralHeaderDetails));
	masterGeneralHeaderDetails.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_EditOptionInHeaderDetailsTab));
	extraFields_EditOptionInHeaderDetailsTab.click();
			 
	Thread.sleep(2000);
			 
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_FieldDetailsCaption));
	extraFields_FieldDetailsCaption.click();
	extraFields_FieldDetailsCaption.clear();
	extraFields_FieldDetailsCaption.sendKeys("Acc Extra Field Update");
	Thread.sleep(1500);
	extraFields_FieldDetailsCaption.sendKeys(Keys.TAB);	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_OkBtn));
	extraFields_OkBtn.click();
	Thread.sleep(2000);
         
	String expMessage="Field Updated Successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	Thread.sleep(2000);
	
	System.out.println("actMessage  : "+actMessage);
	System.out.println("expMessage  : "+expMessage);
	
	boolean actextraFields_CreatedByUserInHeaderDetailsTab =extraFields_CreatedByUserInHeaderDetailsTab.isDisplayed();
	
	
	boolean expextraFields_CreatedByUserInHeaderDetailsTab =true;
	
	System.out.println("***************************checkCreatedFieldDisplayInCustomizeTab*****************************");
	
	
	System.out.println("actextraFields_CreatedByUserInHeaderDetailsTab : "+actextraFields_CreatedByUserInHeaderDetailsTab);
	System.out.println("expextraFields_CreatedByUserInHeaderDetailsTab : "+expextraFields_CreatedByUserInHeaderDetailsTab);
			 
	System.out.println("created field is :" + extraFields_CreatedByUserInHeaderDetailsTab.getText());
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_SaveBtn));
	extraFields_SaveBtn.click();
	Thread.sleep(2000);
	
	String expMessage1="Master updated successfully";
	
	String actMessage1=checkValidationMessage(expMessage1);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_CloseBtn));
	extraFields_CloseBtn.click();
	
	System.out.println("********************************checkExtraFieldCreated***************************");
	
	System.out.println("actMessage1 : "+actMessage1);
	System.out.println("expMessage1 : "+expMessage1);
	
	
	if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) &&actextraFields_CreatedByUserInHeaderDetailsTab==expextraFields_CreatedByUserInHeaderDetailsTab)
	{
		
		return true;
	}
	else
	{
	
		return false;
	}
}
	
	
public boolean checkUpdatedExtraFieldinGeneralTabofAccountMaster() throws InterruptedException
{
	  Thread.sleep(2000);
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	    homeMenu.click();
		
      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	    mastersMenu.click();
	     
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
	    accounts.click();
	    Thread.sleep(2000);
	
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
	  masterNewBtn.click();
	  Thread.sleep(2000);
	  
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdextraFieldinGeneralTab));
	  String actCreatedExtraField=createdextraFieldinGeneralTab.getText();
	  String expCreatedEXtraField="Acc Extra Field Update";
	  
		System.out.println("Message     : "+actCreatedExtraField+" Value Expected : "+expCreatedEXtraField);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();
			
		if(actCreatedExtraField.equalsIgnoreCase(expCreatedEXtraField))
		{
			return true;
		}
		else
		{
			return false;
		}
		 
}
	

@FindBy(xpath="(//*[@id='lblWarnMessage_FieldModal'])[1]")
public static WebElement getMsgonDelField;


@FindBy(xpath="//*[@id='MCustomize_DeleteField_Modal']//input[1]")
public static WebElement clickOnOkInDelField;

public boolean checkDeleteExtraFieldinCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	

	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));
	masterRibbonToExpandOptions.click();
			 
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCustamizemasterBtn));
	masterCustamizemasterBtn.click();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterGeneralField));
	masterGeneralField.click();
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterGeneralHeaderDetails));
	masterGeneralHeaderDetails.click();
	Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_DeleteOptionInHeaderDetailsTab));
	extraFields_DeleteOptionInHeaderDetailsTab.click();
			 
	Thread.sleep(3000);

         
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgonDelField));
     System.out.println(getMsgonDelField.getText());
     
    
     
     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelField));
     clickOnOkInDelField.click();
     
     Thread.sleep(6000);
     
     String expMessage="Field deleted Successfully";
		
     String actMessage=checkValidationMessage(expMessage);
     
     System.out.println("********************************checkDeleteAccountBeforeCustomization***************************");
		
	 System.out.println("actMessage : "+actMessage);
	 System.out.println("actMessage : "+actMessage);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_SaveBtn));
	extraFields_SaveBtn.click();
	Thread.sleep(6000);
	
	String expMessage1="Master updated successfully";
	
	String actMessage1=checkValidationMessage(expMessage1);
	
	System.out.println("********************************checkExtraFieldCreated***************************");
	
	System.out.println("actMessage1     :  "+actMessage1);
	System.out.println("expMessage1     :  "+expMessage1);
	
	if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) )
	{
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_CloseBtn));
		extraFields_CloseBtn.click();
		
		return true;
	}
	else
	{
	
		return false;
	}

}

@FindBy(xpath="//*[@id='newMasterDiv0']//div//label")
public static List<WebElement> detailsList;


public boolean checkDeletedExtraFieldinGeneralTabofAccountMaster() throws InterruptedException
{
	  Thread.sleep(2000);
	  
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	    homeMenu.click();
		
      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	    mastersMenu.click();
	    Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
	    accounts.click();
	    
	    Thread.sleep(2000);
	  boolean result=true;
	
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
	  masterNewBtn.click();
	  Thread.sleep(2000);
	  
	  for(WebElement e:detailsList)
	  {
		  
		  if(e.getText().equals("Acc Extra Field Update"))
		  {
			  result=false;
		  }
		 
	  }
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();
			
		if(result)
		{
			return true;
		}
		else
		{
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

        Thread.sleep(5000);

        String actUserInfo1=userNameDisplay.getText();

        System.out.println("User Info  : "+actUserInfo1);

        System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

  
        String getCompanyTxt1=Company_Name.getText();
        String getLoginCompanyName1=getCompanyTxt1;
        System.out.println("company name  :  "+ getLoginCompanyName1);
        

        String expUserInfo1           ="SU";
        String expLoginCompanyName1   ="Automation Company";

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


public boolean checkLogoutAccountsPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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


@FindBy(xpath="//button[@id='btnCreateTab_MCust']")
public static WebElement createTab; 

@FindBy(xpath="//a[text()='Tree Tab']")//*[@id="masterCustomizingTabList"]/div[2]/div/ul/li[1]/a
public static WebElement createTabTreeTab;

@FindBy(xpath="//a[contains(text(),'Document Tab')]")
public static WebElement createTab_DocumentTab;  

@FindBy(xpath="//a[contains(text(),'Create Tab')]")
public static WebElement createTab_CreateTab;  

@FindBy(xpath="//a[contains(text(),'Image Tab')]")
public static WebElement createTab_ImageTab;  

@FindBy(xpath="(//*[@value='Save'])[2]")
public static WebElement  treeTab_OkBtn;

@FindBy(xpath="//label[@id='lblNameError_TreeTab']")
public static WebElement customizeMasterCreateTabsMessage; 

@FindBy(xpath="//input[@id='txtNewTreeTabCaption']")
public static WebElement  treeTabCaptionTxt;

@FindBy(xpath="//input[@id='txtNewTreeTabName']")
public static WebElement  treeTabTabNameTxt;

@FindBy(xpath="//input[@id='fOption_tab_masterName']")
public static WebElement  treeTabNewMastertxt;

@FindBy(xpath="//input[@id='txtNewTreeTabFieldCaption']")
public static WebElement  treeTab_FieldCaptiontxt;

public boolean checkCreatingaTreeTabinAccountMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(5000);
	  
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  homeMenu.click();
		
      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	  mastersMenu.click();
	    
	  Thread.sleep(2000);
	    
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
	  accounts.click();
	    
	  Thread.sleep(2000);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));
	masterRibbonToExpandOptions.click();
			 
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCustamizemasterBtn));
	masterCustamizemasterBtn.click();
	
	Thread.sleep(5500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab));
	createTab.click();
	Thread.sleep(1500);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTabTreeTab));
	
	boolean actcreateTabTreeTab  		=createTabTreeTab.isDisplayed();
	boolean actcreateTab_DocumentTab  	=createTab_DocumentTab.isDisplayed();
	boolean actcreateTab_CreateTab  	=createTab_CreateTab.isDisplayed();
	boolean actcreateTab_ImageTab  		=createTab_ImageTab.isDisplayed();
	
	boolean expcreateTabTreeTab  		=true;
	boolean expcreateTab_DocumentTab  	=true;
	boolean expcreateTab_CreateTab  	=true;
	boolean expcreateTab_ImageTab  		=true;
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTabTreeTab));
	createTabTreeTab.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTab_OkBtn));
	treeTab_OkBtn.click();   		 
   		 
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCreateTabsMessage));
	
	String actcustomizeMasterCreateTabsMessage  =customizeMasterCreateTabsMessage.getText();
	String expcustomizeMasterCreateTabsMessage  ="Please enter Tab name.";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTabTabNameTxt));
	treeTabTabNameTxt.click();
	treeTabTabNameTxt.sendKeys("Tree Tab");
	treeTabTabNameTxt.sendKeys(Keys.TAB); 
	Thread.sleep(2000);	 
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTabCaptionTxt));
	treeTabCaptionTxt.sendKeys("Tree Tab");
		 
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTabNewMastertxt));
	treeTabNewMastertxt.sendKeys("Account");
	Thread.sleep(3000);
	treeTabNewMastertxt.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTab_FieldCaptiontxt));
	treeTab_FieldCaptiontxt.sendKeys("Tree Tab");	
	Thread.sleep(2000);	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTab_OkBtn));
	treeTab_OkBtn.click();
	Thread.sleep(2000);	
	
	String expMessage="Tab Added Successfully";
	
	String actMessage=checkValidationMessage(expMessage);
	
	System.out.println("actMessage : "+actMessage);
	System.out.println("expMessage : "+expMessage);
	
	System.out.println("createTabTreeTab        : "+actcreateTabTreeTab      +" Value Expected : "+expcreateTabTreeTab);
	System.out.println("createTab_DocumentTab   : "+actcreateTab_DocumentTab +" Value Expected : "+expcreateTab_DocumentTab);
	System.out.println("createTab_CreateTab     : "+actcreateTab_CreateTab   +" Value Expected : "+expcreateTab_CreateTab);
	
	
	if(actcreateTabTreeTab==expcreateTabTreeTab && actcreateTab_DocumentTab==expcreateTab_DocumentTab && actcreateTab_CreateTab==expcreateTab_CreateTab&&actMessage.equalsIgnoreCase(expMessage))
	{	 
		return true;
	}
	else
	{
		return false;
	}
	
}
	
@FindBy(xpath="(//*[@value='Save'])[6]")
public static WebElement  treeTab_UpdateOkBtn;

@FindBy(xpath="(//*[contains(text(),'Tree Tab')])[1]")
public static WebElement customizeMasterCreateTreeTabNewlyCreate; 

@FindBy(xpath="//input[@id='txtEditTabCaption']")
public static WebElement customizeMasterCreateTreeTabNewlyCreateTabCaption; 	

	public boolean checkEditingNewlyCreatedTreeTabDisplay() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCreateTreeTabNewlyCreate));
	
		boolean actcustomizeMasterCreateTreeTabNewlyCreate  =customizeMasterCreateTreeTabNewlyCreate.isDisplayed();
		boolean expcustomizeMasterCreateTreeTabNewlyCreate  =true;
		
		getAction().doubleClick(customizeMasterCreateTreeTabNewlyCreate).build().perform();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCreateTreeTabNewlyCreateTabCaption));
    	customizeMasterCreateTreeTabNewlyCreateTabCaption.clear();
    	
    	customizeMasterCreateTreeTabNewlyCreateTabCaption.sendKeys("Tree Tab Update");	
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTab_UpdateOkBtn));
    	treeTab_UpdateOkBtn.click();
    	Thread.sleep(2000);
    	
    	String expMessage="Tab Caption Updated Successfully.";
			
    	String actMessage=checkValidationMessage(expMessage);
    	
    	System.out.println("actMessage : "+actMessage);
    	System.out.println("expMessage : "+expMessage);
		
    	if(actMessage.equalsIgnoreCase(expMessage))
    	{
	return true;
    	}
    	else
    	{
    		return false;
    	}
	
}

	@FindBy(xpath="(//*[contains(text(),'Tree Tab Delete')])[1]")
	public static WebElement  CustomizeCreateTreeTab;
	
	@FindBy(xpath="//*[@id='lblWarnMessage_TabModal']")
	public static WebElement  CustomizeDeleteTabMessage;
	
	@FindBy(xpath="//*[@id='MCustomize_DeleteTab_Modal']//input[@class='FButton-Primary'][@value='Yes']")
	public static WebElement  CustomizeDeleteTabMessageYesBtn;
	
	
	@FindBy(xpath="//button[@id='btnDeleteTab_MCust']")
	public static WebElement tabs_DeleteBtn; 

	@FindBy(xpath="(//*[contains(text(),'Tree Tab Update')])[1]")
	public static WebElement tabs_TreetabUpdate; 
	
	
	public boolean checkDeleteTabOptionForCreateTreeTabDisplayInTabs() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 Thread.sleep(2000);
				
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab));
    	createTab.click();
    	    	    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTabTreeTab));
    	createTabTreeTab.click();
    	 Thread.sleep(2000);
    	 	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTabCaptionTxt));
    	treeTabCaptionTxt.sendKeys("Tree Tab Delete");
    	treeTabCaptionTxt.sendKeys(Keys.TAB);
    	
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTabTabNameTxt));
    	treeTabTabNameTxt.sendKeys(Keys.TAB);
    	
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTabNewMastertxt));
    	treeTabNewMastertxt.sendKeys("Account");
    	Thread.sleep(3000);
    	treeTabNewMastertxt.sendKeys(Keys.TAB);
   			
    	   		
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTab_FieldCaptiontxt));
    	treeTab_FieldCaptiontxt.sendKeys("Tree Tab Delete");	
    	
    	
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTab_OkBtn));
    	treeTab_OkBtn.click();
    	
    	Thread.sleep(2000);
    	
    	
    	try
    	{
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
	    	errorMessageCloseBtn.click();
    	}
    	catch(Exception e)
    	{
    		System.out.println("Error Message Close Btn Not Displayed");
    	}
    	*/
   		 
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tabs_TreetabUpdate));
    	tabs_TreetabUpdate.click();
   		
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tabs_DeleteBtn));
    	tabs_DeleteBtn.click();
   		 
    	 Thread.sleep(2000);
    	 
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomizeDeleteTabMessage));
    	
    	String actCustomizeDeleteTabMessage  =CustomizeDeleteTabMessage.getText();
    	String expCustomizeDeleteTabMessage  ="Are you sure that you want to delete the selected tab and Fields associated with this";
    	
    	System.out.println("actCustomizeDeleteTabMessage  : "+actCustomizeDeleteTabMessage);
    	System.out.println("expCustomizeDeleteTabMessage  : "+expCustomizeDeleteTabMessage);
    	
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomizeDeleteTabMessageYesBtn));
    	CustomizeDeleteTabMessageYesBtn.click();
   		 Thread.sleep(4000);
   		 
   		String expMessage = "Tab Deleted Successfully";
		
   		String actMessage=checkValidationMessage(expMessage);
   	
   		 
    	System.out.println("**********************************************************");
    	
    	System.out.println("actMessage  : "+actMessage);
    	System.out.println("expMessage  : "+expMessage);
    	
    	if(actCustomizeDeleteTabMessage.equalsIgnoreCase(expCustomizeDeleteTabMessage) && actMessage.equalsIgnoreCase(expMessage))
    	{
			 
			/*
			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * CustomizeMasterCloseBtn)); CustomizeMasterCloseBtn.click();
			 */
    		return true;
    	}
    	else
    	{
			/*
			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * CustomizeMasterCloseBtn)); CustomizeMasterCloseBtn.click();
			 */
    		return false;
    	}
		
	}

	
			@FindBy(xpath="(//*[@id='btnCloseMasterCutomizationFromMasterScreen'])[1]")
			public static WebElement  CustomizeMasterCloseBtn;
			
			@FindBy(xpath="//input[@id='txtNewDocTabCaption']")
			public static WebElement documentTab_TabCaption;  

			@FindBy(xpath="//input[@id='txtNewDocTabName']")
			public static WebElement documentTab_TabName;  
	
			@FindBy(xpath="(//*[@value='Save'])[3]")
			public static WebElement documentTab_OkBtn;  
			
			
	public static boolean checkCreateDocumentTabOptionsInCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
    {
    	Thread.sleep(5000);	 
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab));
    	createTab.click();
	 	     	         
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab_DocumentTab));
    	createTab_DocumentTab.click();
    	Thread.sleep(2000);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentTab_TabCaption));
		documentTab_TabCaption.sendKeys("Document Tab");	
		Thread.sleep(2000);
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentTab_OkBtn));
   		documentTab_OkBtn.click();
   		 
   		String expMessage="Tab Added Successfully";
		
   		String actMessage=checkValidationMessage(expMessage);
   		Thread.sleep(2000);
   		
   		System.out.println("********************************checkOKWithInputCheckInDocumentTab***************************");
		
		System.out.println("actMessage  : "+actMessage);
		System.out.println("expMessage  : "+expMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage))
   		{
			
			/*  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomizeMasterCloseBtn));
			  CustomizeMasterCloseBtn.click();*/
			 
   			 return true;
   		}
   		else
   		{    
   			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomizeMasterCloseBtn));
   			CustomizeMasterCloseBtn.click();*/
   			
   			 return false;
   		}
    	   	
    }
	
	public boolean checkCreatingaTreeTabAfterLogoutinAccountMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	    homeMenu.click();
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	    mastersMenu.click();
	     
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
	    accounts.click();
	    Thread.sleep(4000);
	    
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));
		masterRibbonToExpandOptions.click();
				 
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCustamizemasterBtn));
		masterCustamizemasterBtn.click();
		
		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCustamizemasterBtn));	
		masterCustamizemasterBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab));
		createTab.click();
		Thread.sleep(1500);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTabTreeTab));
		createTabTreeTab.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTabTabNameTxt));
		treeTabTabNameTxt.click();
		treeTabTabNameTxt.sendKeys("Tree Tab New");
		treeTabTabNameTxt.sendKeys(Keys.TAB); 
		Thread.sleep(2000);	 
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTabCaptionTxt));
		treeTabCaptionTxt.sendKeys("Tree Tab New");
			 
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTabNewMastertxt));
		treeTabNewMastertxt.sendKeys("Account");
		Thread.sleep(3000);
		treeTabNewMastertxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTab_FieldCaptiontxt));
		treeTab_FieldCaptiontxt.sendKeys("Tree Tab");	
		Thread.sleep(2000);	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(treeTab_OkBtn));
		treeTab_OkBtn.click();
		Thread.sleep(2000);	
		
		String expMessage="Tab Added Successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		
		
		
	
		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{	 
			return true;
		}
		else
		{
			return false;
		}
		

	}
	
	  public static boolean checkCreateDocumentTabCreationWithExistingDocumentTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	    {
	    	Thread.sleep(10000);	
	    	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab));
	    	createTab.click();
		 	     	         
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab_DocumentTab));
	    	createTab_DocumentTab.click();
	    	Thread.sleep(2000);
	    	
	        
	        String expMessage="Only one document tab is allowed per master";
			
	        String actMessage=checkValidationMessage(expMessage);
	        
	        System.out.println("********************************checkCreateDocumentTabCreationWithExistingDocumentTab***************************");
			
			System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
	   		{
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomizeMasterCloseBtn));
				  CustomizeMasterCloseBtn.click();
	   			 return true;
	   		}
	   		else
	   		{
	   		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CustomizeMasterCloseBtn));
			  CustomizeMasterCloseBtn.click();
	   			 return false;
	   		}
	    }
	
	  
	  @FindBy(xpath="//input[@id='txtNewTabCaption']")
		public static WebElement createTab_TabCaptionTxt;  

		@FindBy(xpath="//input[@id='txtNewTabName']")
		public static WebElement createTab_TabNameTxt;  
	  
	  
	  public static boolean checkCreateCreateTabInCustomizeMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	    {
	    
		  	Thread.sleep(2000);
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab));
	    	createTab.click();
		 	 
	    	Thread.sleep(2000);
	         
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab_CreateTab));
	    	createTab_CreateTab.click();
	                      
	    	Thread.sleep(2000);
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab_TabCaptionTxt));
	    	createTab_TabCaptionTxt.sendKeys("Create Tab");
	    	createTab_TabCaptionTxt.sendKeys(Keys.TAB); 

	    	String getTabName=createTab_TabNameTxt.getText();
	         
	    	         
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab_TabCaptionTxt));
	    	Thread.sleep(1000);
	         
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab_PopOkBtn));
	    	createTab_PopOkBtn.click();         
	         
	    	String expMessage="Tab Added Successfully";
	 		
	    	String actMessage=checkValidationMessage(expMessage);
	         
	    	System.out.println("********************************checkOkButtonWithAllInputInCreateCreateTab***************************");
	 		
	 		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
	 		
	 		if(actMessage.equalsIgnoreCase(expMessage))
	 		{
	 			
	 			return true;
	 		}
	 		else
	 		{
	 			
	 			return false;
	 		}

	    }
	  
	  @FindBy(xpath="(//*[@value='Save'])[5]")
		public static WebElement createTab_PopOkBtn;  
	    	
	  @FindBy(xpath="//*[@id='masterFields_list']/li[9]")
		public static WebElement customizeMasterCreateCreateTabNewlyCreateTabCaption; 	
	  
		@FindBy(xpath="//*[@id='btnAdd_Field_MasterCust']/a/i")
		public static WebElement customizeMasterCreateCreateTabAddButton; 	
		
		@FindBy(xpath="//*[@id='tab_HeaderDetails_8']")
		public static WebElement customizeMasterCreateCreateTabHeaderDetails; 	
		
		@FindBy(xpath="//*[@id='ExtraField_Tabs']/li[1]")
		public static WebElement extraFields_FieldDetailsTab;   
	   
	    public static boolean checkAddExtraFieldOptionInCreateCreateTabDisplayInCustomizeTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	    {
	    	//getAction().moveToElement(customizeMasterCreateCreateTabNewlyCreateTabCaption).build().perform();
	    	Thread.sleep(2000);
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCreateCreateTabNewlyCreateTabCaption));
	    	customizeMasterCreateCreateTabNewlyCreateTabCaption.click();
	      		    			 
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCreateCreateTabHeaderDetails));
	    	customizeMasterCreateCreateTabHeaderDetails.click();
	             
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCreateCreateTabAddButton));
	    	customizeMasterCreateCreateTabAddButton.click();

	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_FieldDetailsTab));
	    	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCreateCreateTabAddButton));
	    	customizeMasterCreateCreateTabAddButton.click();
	      		    			
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_FieldDetailsCaption));
	    	extraFields_FieldDetailsCaption.sendKeys("Extra Field One");
	             
	             
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_OkBtn));
	    	extraFields_OkBtn.click();
	    	Thread.sleep(2000);
	    	
	    	String expMessage="Field Added Successfully.";
	    	
	    	String actMessage=checkValidationMessage(expMessage);
	    	
	    	System.out.println("********************************checkOkButtonWithInputMandatoryFieldsCreateCreateTabDisplayInCustomizeTab***************************");
	    	
	    	System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
	    	
	    	if(actMessage.equalsIgnoreCase(expMessage))
	    	{
	    		return true;
	    	}
	    	else
	    	{
	    		return false;
	    	}
	    }	  
	  


	    
	    @FindBy(xpath="//*[contains(text(),'Document Tab')]")
	    public static WebElement createed_DocumentTab;
public boolean checkAfterTabCreationinAccountMasterPage() throws InterruptedException
{
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
    homeMenu.click();
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
    mastersMenu.click();
    Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
    accounts.click();
    
    Thread.sleep(3000);
    
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
	  masterNewBtn.click();
	  Thread.sleep(2000);
	  
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createed_DocumentTab));
	  boolean actresult=createed_DocumentTab.isDisplayed();
	  boolean expresult=true;
	  
	  System.err.println("Result is : "+createed_DocumentTab.getText());
	  
	  if(actresult==expresult)
	  {
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		  masterCloseBtn.click();
		  return true;
	  }
    
    
	  else
	  {
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		  masterCloseBtn.click();
		  return false;
	  }
    
}
	
	public AccountsPage(WebDriver driver)
    {
		PageFactory.initElements(driver, this);
    }
	
	

}
