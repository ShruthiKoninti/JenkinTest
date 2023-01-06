package com.focus.Pages;

import java.awt.AWTException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;

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
			
	        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		    mastersMenu.click();
		     
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
		    accounts.click();
		    	  
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		    masterNewBtn.click();
			 
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
			 
			String expMessage=excelReader.getCellData(xlSheetName, 8, 6);
				
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, xlSheetName, 8, 7, actMessage);
			 
			System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
			masterCloseBtn.click();
				
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
		 
		 String expMessage = excelReader.getCellData(xlSheetName, 10, 6);
			
		 String actMessage=checkValidationMessage(expMessage);
		 
		 excelReader.setCellData(xlfile, xlSheetName, 10, 7, actMessage);
		 
		 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		 
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
			    
		 String expaccountNewCreationName  =excelReader.getCellData(xlSheetName, 13, 6);
		 
		 System.out.println("************************checkUpdatedAccountDisplayInAccountMasterGridBeforeCustomization***********************");
		 
		 System.err.println("accountNewCreationName  : "+actaccountNewCreationName+" Value Expected : "+expaccountNewCreationName);
		 
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
		
		 int count = masterGridBodyName.size();
			
			for (int i = 0; i < count; i++) 
			{
				String data = masterGridBodyName.get(i).getText();
				
				System.err.println("*"+data+"*");
				
				if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 14, 5))) 
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
         
         String expMessage=excelReader.getCellData(xlSheetName, 14, 6);
 		
         String actMessage=checkValidationMessage(expMessage);
         
         excelReader.setCellData(xlfile, xlSheetName, 14, 7, actMessage);
         
         System.out.println("********************************checkDeleteAccountBeforeCustomization***************************");
			
		 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			
		 if(actMessage.equalsIgnoreCase(expMessage))
		 {
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			 closeBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 14, 8, resPass);
			 return true;
		 }
		 
		 else
		 {
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			 closeBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 14, 8, resFail);
			 return false;
		 }
	}
	
	
	///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	
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

	@FindBy(xpath="//span[@id='btnImportDataForAdvanceMasterPopUp']")
	public static WebElement advMasterImportDataBtn;;


	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[11]")
	private static List<WebElement> accNameList;

	
	public String actAccList=null;
	public String expAccList=null;
	
		
	public boolean checkAccountsListBeforeImportingDataFromExcel() throws InterruptedException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		   	homeMenu.click();
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		   	mastersMenu.click();
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
			     
		   	accounts.click();
			   
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
	      	
	  	if(advanceMasterLabel.getText().equalsIgnoreCase("Advance Master Import/Export----> Account"))
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
		 Thread.sleep(6000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterImportDataBtn));
		 advMasterImportDataBtn.click();
		 Thread.sleep(4000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterCloseBtn));
		 advMasterCloseBtn.click();
		 Thread.sleep(2000);
		 
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
		Thread.sleep(2000);
		
		
		 int accCount = accNameList.size();
			

		 ArrayList<String> acclist = new ArrayList<String>();

		 		for (int i = 0; i < accCount; i++)
		 		{
		 			 String data = accNameList.get(i).getText();
		 			acclist.add(data);
		 			
		 		}
		 		
		 		 actAccList=acclist.toString();
		 		 expAccList="[Pricebook Vendor Account Two, Pricebook Vendor Account One, Pricebook Account One, Pricebook Account Two, Pricebook Group Two, Pricebook Group One, Round off Exchange gain / loss, SALES, ASSETS, EXPENSES, CONTROL ACCOUNTS, REVENUE, EQUITIES, LIABILITIES, COGS POSTING ACC, BR COGS ACC INV, FIFO COGS ACC INV, WA COGS ACC INV, STD RATE COGS ACC INV, SR COGS POSTING ACC, SHORTAGE COGS POSTING ACC, EXCESS COGS POSTING ACC, VAT OUTPUT, VAT ADVANCE SALE, VAT ADVANCE PURCHASE, PURCHASE VARIANCE, VAT INPUT, AccountImport]";
		 		
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

System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

Thread.sleep(1000);

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

@FindBy(xpath="//*[@id='ol_treeNavigation']/li[2]/span")
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


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
masterAddGroupBtn.click();

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

Thread.sleep(2000);


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


String expMessage="Saved Successfully";

String actMessage=checkValidationMessage(expMessage);

System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

Thread.sleep(1000);

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

boolean s=accountGrp.isDisplayed();
if(s)
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


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
masterNewBtn.click();

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

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
masterCloseBtn.click();

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

Thread.sleep(2000);
String expMessage="Saved Successfully";

String actMessage=checkValidationMessage(expMessage);

System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

Thread.sleep(1000);

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


public static boolean checkAccountsSavedinGroupsandSubGroupsinAccountMaster() throws InterruptedException
{
getDriver().navigate().refresh();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));

accounts.click();

Thread.sleep(2000);


int accCount = accNameList.size();

ArrayList<String> acclist = new ArrayList<String>();

for (int i = 0; i < accCount; i++)
{
String data = accNameList.get(i).getText();
acclist.add(data);

}

String actAccList=acclist.toString();
String  expAccList="[Group Account, Pricebook Vendor Account Two, Pricebook Vendor Account One, Pricebook Account One, Pricebook Account Two, Pricebook Group Two, Pricebook Group One, Round off Exchange gain / loss, ASSETS, EXPENSES, CONTROL ACCOUNTS, REVENUE, EQUITIES, LIABILITIES, COGS POSTING ACC, BR COGS ACC INV, FIFO COGS ACC INV, WA COGS ACC INV, STD RATE COGS ACC INV, SR COGS POSTING ACC, SHORTAGE COGS POSTING ACC, EXCESS COGS POSTING ACC, VAT OUTPUT, VAT ADVANCE SALE, VAT ADVANCE PURCHASE, PURCHASE VARIANCE, VAT INPUT]";


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


System.out.println("Actual Accounts List:  "+actAccList);
System.out.println("Expected Accounts List:  "+expAccList);

System.out.println("Actual Sub Group Accounts List:  "+actGrpAccList);
System.out.println("Expected Sub Group Accounts List:  "+expGrpAccList);

System.out.println("Actual  Accounts in Sub Group1 List:  "+actSubGrpAccList);
System.out.println("Expected Accounts in Sub Group1 List:  "+expSubGrpAccList);

System.out.println("Actual  Accounts in Sub Group2 List:  "+actSubGrpAcc2List);
System.out.println("Expected Accounts in Sub Group2 List:  "+expSubGrpAcc2List);


if(actAccList.equalsIgnoreCase(expAccList) && actGrpAccList.equalsIgnoreCase(expGrpAccList)
&& actSubGrpAccList.equalsIgnoreCase(expSubGrpAccList) && actSubGrpAcc2List.equalsIgnoreCase(expSubGrpAcc2List))
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


public boolean checkDeleteGroupAccountinAccountMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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

	



public boolean checkOpenGroupAccountMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	

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
	
	
	System.out.println(actRow1List);
	System.out.println(expRow1List);	
	
	
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
	String actMessage=BaseEngine.restoreCompany("SanityMounted","Automation Company");
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
	

	
	
	
	public AccountsPage(WebDriver driver)
    {
		PageFactory.initElements(driver, this);
    }
	
	

}
