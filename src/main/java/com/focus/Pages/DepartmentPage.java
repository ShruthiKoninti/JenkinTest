package com.focus.Pages;

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class DepartmentPage extends BaseEngine 
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "DepartmentPage";
	

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


	public  boolean checkDepartmentMasterInMAsterMenu() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		System.out.println("***********checkDepartmentMasterInMAsterMenu Method Excutes **********");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
		departmentMenu.click();

		Thread.sleep(2000);
		
        String expected = excelReader.getCellData(xlSheetName, 8, 6);
    	
    	String actual = Boolean.toString(masterNewBtn.isDisplayed());
    	
    	System.out.println("Act Save Button is Displayed : "+actual);
    	System.out.println("Exp Save Button is Displayed : "+expected);   	
    	
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





	public static boolean checkSavingDepartmentInDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		System.out.println("********************************checkSavingUnitsInDepartmentMaster Method Excutes **********");

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

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 9, 6);
		String actMessage=checkValidationMessage(expMessage);

		Thread.sleep(8000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();

		System.out.println("actMessage  : " +actMessage);
		System.out.println("expMessage  : " +expMessage);
		
		Thread.sleep(4000);
		
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




	public static boolean checkEditingInSavedDepartmentINDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		Thread.sleep(8000);
		
		String actaccountNewCreationName = null;

		int count = masterDepartmentGridBodyName.size();

		for (int i = 0; i < count; i++) 
		{
			String data = masterDepartmentGridBodyName.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 11, 5))) 
			{
				getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
				actaccountNewCreationName=data;
				break;
			}
		}

		String expaccountNewCreationName  =excelReader.getCellData(xlSheetName, 11, 6);

		System.out.println("********************************checkEditingInSavedDepartmentINDepartmentMaster********************");

		System.out.println("Act Saved Department : " +actaccountNewCreationName);
		System.out.println("Exp Saved Department : " +expaccountNewCreationName);
		
		
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actaccountNewCreationName);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys(excelReader.getCellData(xlSheetName, 12, 5));
		nameTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys(excelReader.getCellData(xlSheetName, 13, 5));
		codeTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 12, 6);
		String actMessage=checkValidationMessage(expMessage);

		System.out.println("actMessage    : "+actMessage);
		System.out.println("expMessage    : "+expMessage);
				
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actaccountNewCreationName);
		
		Thread.sleep(8000);

		if(actaccountNewCreationName.equalsIgnoreCase(expaccountNewCreationName) && actMessage.equalsIgnoreCase(expMessage))
		{
			
			excelReader.setCellData(xlfile, xlSheetName, 11, 8, resPass);

			return true;
		}

		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 8, resFail);

			return false;
		}
	}




	public static boolean checkUpdatedDepartmentMasterDisplayInDepartment() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		String actaccountNewCreationName = null;

		int count = masterDepartmentGridBodyName.size();

		for (int i = 0; i < count; i++) 
		{
			String data = masterDepartmentGridBodyName.get(i).getText();

			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 14, 5))) 
			{
				if (masterGridBodyChkbox.get(i).isSelected()==false) 
				{
					masterGridBodyChkbox.get(i).click();
				}
				actaccountNewCreationName=data;
				break;
			}
		}

		String expaccountNewCreationName  =excelReader.getCellData(xlSheetName, 14, 6);

		System.out.println("************************checkUpdatedDepartmentMasterDisplayInDepartment***********************");

		System.err.println("actaccountNewCreationName  : " +actaccountNewCreationName);
		System.err.println("expaccountNewCreationName  : " +expaccountNewCreationName);
		
		 excelReader.setCellData(xlfile, xlSheetName, 14, 7, actaccountNewCreationName);

		if(actaccountNewCreationName.equalsIgnoreCase(expaccountNewCreationName))
		{ 
			 excelReader.setCellData(xlfile, xlSheetName, 14, 8, resPass);
			
			return true;
		}
		else
		{
			
			excelReader.setCellData(xlfile, xlSheetName, 14, 8, resFail);
			return false;
		}
	}






	public static boolean checkDeleteInDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		System.out.println("************************checkDeleteInDepartmentMaster  Method EXCUTES ***************");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));

		masterDeleteBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgOnDelete));
		System.out.println(getMsgOnDelete.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
		clickOnOkInDelete.click();

		System.out.println("********************************checkDeleteItemInItemMaster***************************");

		String expMessage= excelReader.getCellData(xlSheetName, 15, 6);
		String actMessage=checkValidationMessage(expMessage);
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actMessage);
		
		System.err.println("actMessage  : " +actMessage);
		System.err.println("expMessage  : " +expMessage);

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



	@FindBy(xpath="//*[@id='drpdownSlider']")
	public static WebElement toggleRibbon;

	@FindBy(xpath="//*[@id='btnAdvImportForMaster']")
	public static WebElement advanceMasterImportExportBtn;;

	@FindBy(xpath="//*[contains(text(),'Advance Master Import/Export')]")
	public static WebElement advanceMasterLabel;
					 
	@FindBy(xpath="//button[@id='btnFile']")
	public static WebElement FileUploadBtn;
		

	@FindBy(xpath="//*[@class='icon-font7 icon-close']")
	public static WebElement advMasterCloseBtn;


	@FindBy(xpath="//*[@id='btnImportDataForAdvanceMasterPopUp']")
	public static WebElement advMasterImportDataBtn;;


	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[10]")
	private static List<WebElement> masterGridBodyName;

	

	
	
	public boolean checkDepartmentsListBeforeImportingDataFromExcel() throws InterruptedException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		   	homeMenu.click();
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		   	mastersMenu.click();
		   	
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
			     
		   	departmentMenu.click();
			   
		   	Thread.sleep(4000);
		   	
		
		 
		 int accCount = masterGridBodyName.size();
		

		 ArrayList<String> acclist = new ArrayList<String>();

		 		for (int i = 0; i < accCount; i++)
		 		{
		 			 String data = masterGridBodyName.get(i).getText();
		 			acclist.add(data);
		 			
		 		}
		 		
		 		String actDeptList=acclist.toString();
		 		String  expDeptList="[CurrencyDepartment, DepartmentST, INDIA, DUBAI, AMERICA, SINGPORE, EUROPE]";
		 		
		 		System.out.println("Actual Departments List  :  " +actDeptList);
		 		System.out.println("Expected Departments List:  " +expDeptList);
		 		
		 		if(actDeptList.equalsIgnoreCase(expDeptList))
		 		{
		 			return true;
		 		}
		 		else
		 		{
		 			return false;
		 		}
		 		
	}

	public boolean checkAdvanceMasterImportExportDepartment() throws InterruptedException
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
	public boolean checkImportDatainAdvanceMasterImportExportDepartment() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
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
		
		 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\DepartmentImportFromExcel.exe");
		 Thread.sleep(5000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterImportDataBtn));
		 advMasterImportDataBtn.click();
		 Thread.sleep(10000);
		/* 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterCloseBtn));
		 advMasterCloseBtn.click();
		 Thread.sleep(2000);*/
		 
		 
		 String expMessage1="8 Records Imported Successfully";
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

	@FindBy(xpath="//*[@id='LandingGridBody']/tr[7]/td")
	public static List<WebElement> department7thRowList;

	public boolean checkDepartmentsListAfterImportingDataFromExcel() throws InterruptedException
	{
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	   	homeMenu.click();
	   	
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	   	mastersMenu.click();
	   	
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
		     
	   	departmentMenu.click();
		   
	   	Thread.sleep(6000);
	   	
	
	 
	 int accCount = masterGridBodyName.size();
	

	 ArrayList<String> acclist = new ArrayList<String>();

	 		for (int i = 0; i < accCount; i++)
	 		{
	 			 String data = masterGridBodyName.get(i).getText();
	 			acclist.add(data);
	 			
	 		}
	 		
	 		String actDeptList=acclist.toString();
	 		String  expDeptList="[CurrencyDepartment, DepartmentST, INDIA, DUBAI, AMERICA, SINGPORE, EUROPE, DepartmentImport]";
	 		
	 		System.out.println("Actual Departments List:  "+actDeptList);
	 		System.out.println("Expected Departments List:  "+expDeptList);
	 		
	 		if(actDeptList.equalsIgnoreCase(expDeptList))
	 		{
	 			return true;
	 		}
	 		else
	 		{
	 			return false;
	 		}}
	
	
public boolean checkImportedDepartmentDelete() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{/*
	Thread.sleep(2000);
	getDriver().navigate().refresh();
	Thread.sleep(4000);
	
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	   	homeMenu.click();
	   	Thread.sleep(2000);
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	   	mastersMenu.click();
	   	Thread.sleep(2000);
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
		     
	   	departmentMenu.click();*/
		   
	   	Thread.sleep(6000);
	
	
	 int count = masterGridBodyName.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = masterGridBodyName.get(i).getText();
			
			if (data.equalsIgnoreCase("DepartmentImport")) 
			{
				if (masterGridBodyChkbox.get(i).isSelected()==false) 
				{
					masterGridBodyChkbox.get(i).click();
					Thread.sleep(2000);
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
  
  System.out.println("********************************checkDeleteDepartmentBeforeCustomization***************************");
		
	 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		
	 if(actMessage.equalsIgnoreCase(expMessage))
	 {
		/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		 closeBtn.click();*/
		 return true;
	 }
	 
	 else
	 {
		/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		 closeBtn.click();*/
		 return false;
	 }
}

	
/////////////////////Adding Groups and subgroups//////////////////////////

public boolean checkSavingDepartmentGroupinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	getDriver().navigate().refresh();
	Thread.sleep(4000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	mastersMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
	departmentMenu.click();
	Thread.sleep(3000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
	masterAddGroupBtn.click();

	Thread.sleep(3000);

	System.out.println("********************************checkSavingDepartmentGroupDepartmentMaster*****************************");

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
	nameTxt.click();
	nameTxt.clear();
	nameTxt.sendKeys("Group Department");
	nameTxt.sendKeys(Keys.TAB);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
	codeTxt.click();
	codeTxt.clear();
	codeTxt.sendKeys("GroupDepartment");
	codeTxt.sendKeys(Keys.TAB);


	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();

	String expMessage="Saved Successfully";

	String actMessage=checkValidationMessage(expMessage);

	System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

	Thread.sleep(6000);

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

@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[10]")
private static List<WebElement> accNameList;

@FindBy(xpath="//*[@id='ol_treeNavigation']/li[1]")
public static WebElement DepartmentLabel;


@FindBy(xpath="//*[@id='ol_treeNavigation']/li[2]")
public static WebElement GroupAccountLabel;

public static boolean checkSavingDepartmentSubGroupinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
int accNameListCount = accNameList.size();

for(int i=0;i<accNameListCount;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("Group Department"))
{
Thread.sleep(2000);
getAction().doubleClick(accChkBoxList.get(i)).build().perform();
Thread.sleep(2000);
break;
}
}


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
masterAddGroupBtn.click();

System.out.println("********************************checkSavingDepartmentSubGroupDepartmentMaster*****************************");

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("SubGroup1 Department");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("SubGroup1Department");
codeTxt.sendKeys(Keys.TAB);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();

Thread.sleep(10000);


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
nameTxt.click();
nameTxt.clear();
nameTxt.sendKeys("SubGroup2 Department");
nameTxt.sendKeys(Keys.TAB);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
codeTxt.click();
codeTxt.clear();
codeTxt.sendKeys("SubGroup2Department");
codeTxt.sendKeys(Keys.TAB);



getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
saveBtn.click();

Thread.sleep(6000);

String expMessage="Saved Successfully";

String actMessage=checkValidationMessage(expMessage);

System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

Thread.sleep(6000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
masterCloseBtn.click();

Thread.sleep(6000);

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


@FindBy(xpath="//*[@id='lblnchecked']/input")
public static List<WebElement> accChkBoxList;


public boolean checkSavingDepartmentsinSubGroupsinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	
	Thread.sleep(4000);
	getDriver().navigate().refresh();
	Thread.sleep(4000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	mastersMenu.click();
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
	departmentMenu.click();
	Thread.sleep(4000);
	
	int accNameListCount3 = accNameList.size();

	for(int i=0;i<accNameListCount3;i++)
	{
	String data = accNameList.get(i).getText();
	Thread.sleep(2000);
	if(data.equals("Group Department"))
	{
	Thread.sleep(2000);
	getAction().doubleClick(accChkBoxList.get(i)).build().perform();
	Thread.sleep(2000);
	break;
	}
	}
	Thread.sleep(2000);
		int accNameListCount = accNameList.size();
		
		for(int i=0;i<accNameListCount;i++)
		{
		String data = accNameList.get(i).getText();
		Thread.sleep(2000);
		if(data.equals("SubGroup1 Department"))
		{
		Thread.sleep(2000);
		getAction().doubleClick(accChkBoxList.get(i)).build().perform();
		Thread.sleep(2000);
		break;
		}
		}
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		masterNewBtn.click();
		
		System.out.println("********************************checkSavingDepartmentSubGroupDepartmentMaster*****************************");
		Thread.sleep(6000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys("Department1");
		nameTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys("Department1");
		codeTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();
		
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupAccountLabel));
		GroupAccountLabel.click();
		Thread.sleep(4000);
		
		int accNameListCount1 = accNameList.size();
		
		for(int i=0;i<accNameListCount1;i++)
		{
		String data = accNameList.get(i).getText();
		Thread.sleep(2000);
		if(data.equals("SubGroup2 Department"))
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
		nameTxt.sendKeys("Department2");
		nameTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys("Department2");
		codeTxt.sendKeys(Keys.TAB);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		
		
		String expMessage="Saved Successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();
		Thread.sleep(6000);
		if(actMessage.equalsIgnoreCase(expMessage))
		{
		return true;
		}
		else
		{
		return false;
		}
}


public static boolean checkDepartmentsSavedinGroupsandSubGroupsinDepartmentMaster() throws InterruptedException
{
/*getDriver().navigate().refresh();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
departmentMenu.click();*/

	Thread.sleep(4000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentLabel));
	DepartmentLabel.click();
	Thread.sleep(2000);
	
int accCount = accNameList.size();

ArrayList<String> acclist = new ArrayList<String>();

for (int i = 0; i < accCount; i++)
{
String data = accNameList.get(i).getText();
acclist.add(data);

}

String actAccList=acclist.toString();
String  expAccList="[Group Department, CurrencyDepartment, DepartmentST, INDIA, DUBAI, AMERICA, SINGPORE, EUROPE]";


int accNameListCount = accNameList.size();

for(int i=0;i<accNameListCount;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("Group Department"))
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
String  expGrpAccList="[SubGroup2 Department, SubGroup1 Department]";


int accNameListCount1 = accNameList.size();
for(int i=0;i<accNameListCount1;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("SubGroup1 Department"))
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
String  expSubGrpAccList="[Department1]";
		


getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupAccountLabel));
GroupAccountLabel.click();
Thread.sleep(2000);

int accNameListCount3 = accNameList.size();
for(int i=0;i<accNameListCount3;i++)
{
String data = accNameList.get(i).getText();
Thread.sleep(2000);
if(data.equals("SubGroup2 Department"))
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
String  expSubGrpAcc2List="[Department2]";
		
System.out.println("********************************checkSavingGroupsandSubGroupDepartmentMaster*****************************");


System.out.println("Actual Departments List:  "+actAccList);
System.out.println("Expected Departments List:  "+expAccList);

System.out.println("Actual Sub Group Departments List:  "+actGrpAccList);
System.out.println("Expected Sub Group Departments List:  "+expGrpAccList);

System.out.println("Actual  Departments in Sub Group1 List:  "+actSubGrpAccList);
System.out.println("Expected Departments in Sub Group1 List:  "+expSubGrpAccList);

System.out.println("Actual  Departments in Sub Group2 List:  "+actSubGrpAcc2List);
System.out.println("Expected Departments in Sub Group2 List:  "+expSubGrpAcc2List);


if(actAccList.equalsIgnoreCase(expAccList) && actGrpAccList.equalsIgnoreCase(expGrpAccList)
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



public boolean checkDeleteGroupDepartmentinDepartmentMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException		

		{
			/*Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
		departmentMenu.click();*/
		Thread.sleep(6000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentLabel));
		DepartmentLabel.click();
		Thread.sleep(2000);
		
		int count1 = masterGridBodyName.size();
		
		for (int i = 0; i < count1; i++) 
		{
		String data = masterGridBodyName.get(i).getText();
		Thread.sleep(2000);
		
		if (data.equalsIgnoreCase("Group Department")) 
			{
				if (masterGridBodyChkbox.get(i).isSelected()==false) 
				{
					masterGridBodyChkbox.get(i).click();
					Thread.sleep(2000);
					break;
				}
			}
		}
		
		
		
		
		Thread.sleep(4000);	
		//new WebDriverWait(getDriver(), 500).until(ExpectedConditions.visibilityOf(masterDeleteBtn));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));
		
		masterDeleteBtn.click();
		
		Thread.sleep(4000);
		
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


	


///Adding ExtraField


		public boolean checkCreatingExtraFieldsinDepartmentMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
			departmentMenu.click();

			
			Thread.sleep(3000);
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
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCreateExtraFieldAddButton));
			customizeMasterCreateExtraFieldAddButton.click();
			Thread.sleep(2000);			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_FieldDetailsCaption));
			extraFields_FieldDetailsCaption.sendKeys("Department Extra Field ");
			extraFields_FieldDetailsCaption.sendKeys(Keys.TAB);
		         
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_OkBtn));
			extraFields_OkBtn.click();
		         
			String expMessage="Field Added Successfully.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			
			boolean actextraFields_CreatedByUserInHeaderDetailsTab =extraFields_CreatedByUserInHeaderDetailsTab.isDisplayed();
			
			
			boolean expextraFields_CreatedByUserInHeaderDetailsTab =true;
			
			System.out.println("***************************checkCreatedFieldDisplayInCustomizeTab*****************************");
			
			
			System.out.println("extraFields_CreatedByUserInHeaderDetailsTab : "+actextraFields_CreatedByUserInHeaderDetailsTab+" Value Expected : "+expextraFields_CreatedByUserInHeaderDetailsTab);
					 
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_SaveBtn));
			extraFields_SaveBtn.click();
			
			
			String expMessage1="Master updated successfully";
			
			String actMessage1=checkValidationMessage(expMessage1);
			
			System.out.println("********************************checkExtraFieldCreated***************************");
			
			System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			System.out.println("Message     : "+actMessage1+" Value Expected : "+expMessage1);
			
			if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) &&actextraFields_CreatedByUserInHeaderDetailsTab==expextraFields_CreatedByUserInHeaderDetailsTab)
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
		 
		 
			
			
		@FindBy(xpath="//*[@id='newMasterDiv0']//div[19]//label/a")
		public static WebElement createdextraFieldinGeneralTab;
			
			
		public boolean checkAddedExtraFieldinGeneralTabofDepartmentMaster() throws InterruptedException
		{
			
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
			  masterNewBtn.click();
			  Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdextraFieldinGeneralTab));
			  String actCreatedExtraField=createdextraFieldinGeneralTab.getText();
			  String expCreatedEXtraField="Department Extra Field";
			  
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
			


		@FindBy(xpath="(//*[@id='example']/tbody/tr[3]/td[1]/i[1])[1]")
		public static WebElement extraFields_EditOptionInHeaderDetailsTab;

		@FindBy(xpath="(//*[@id='example']/tbody/tr[3]/td[1]/i[2])[1]")
		public static WebElement extraFields_DeleteOptionInHeaderDetailsTab;


		public boolean checkEditandUpdateExtraFieldinCustomizeMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));
			masterRibbonToExpandOptions.click();

			Thread.sleep(2000);
			
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
			extraFields_FieldDetailsCaption.sendKeys("Department Extra Field Update");
			Thread.sleep(1500);
			extraFields_FieldDetailsCaption.sendKeys(Keys.TAB);	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_OkBtn));
			extraFields_OkBtn.click();
			Thread.sleep(2000);
		         
			String expMessage="Field Updated Successfully";
			
			String actMessage=checkValidationMessage(expMessage);
			
			Thread.sleep(2000);
			
			boolean actextraFields_CreatedByUserInHeaderDetailsTab =extraFields_CreatedByUserInHeaderDetailsTab.isDisplayed();
			
			
			boolean expextraFields_CreatedByUserInHeaderDetailsTab =true;
			
			System.out.println("***************************checkCreatedFieldDisplayInCustomizeTab*****************************");
			
			
			System.out.println("extraFields_CreatedByUserInHeaderDetailsTab : "+actextraFields_CreatedByUserInHeaderDetailsTab+" Value Expected : "+expextraFields_CreatedByUserInHeaderDetailsTab);
					 
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_SaveBtn));
			extraFields_SaveBtn.click();
			Thread.sleep(2000);
			
			String expMessage1="Master updated successfully";
			
			String actMessage1=checkValidationMessage(expMessage1);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_CloseBtn));
			extraFields_CloseBtn.click();
			
			System.out.println("********************************checkExtraFieldCreated***************************");
			
			System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			System.out.println("Message     : "+actMessage1+" Value Expected : "+expMessage1);
			
			if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) &&actextraFields_CreatedByUserInHeaderDetailsTab==expextraFields_CreatedByUserInHeaderDetailsTab)
			{
				
				
				
				
				return true;
			}
			else
			{
			
				return false;
			}
		}
			
			
		public boolean checkUpdatedExtraFieldinGeneralTabofDepartmentMaster() throws InterruptedException
		{
			  Thread.sleep(2000);
			
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
			  masterNewBtn.click();
			  Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdextraFieldinGeneralTab));
			  String actCreatedExtraField=createdextraFieldinGeneralTab.getText();
			  String expCreatedEXtraField="Department Extra Field Update";
			  
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
		     Thread.sleep(2000);
		     
		     String expMessage="Field deleted Successfully";
				
		     String actMessage=checkValidationMessage(expMessage);
		     
		    
		     
		     System.out.println("********************************checkDeleteDepartmentBeforeCustomization***************************");
				
			 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
				
			

			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_SaveBtn));
			extraFields_SaveBtn.click();
			 Thread.sleep(2000);
			
			String expMessage1="Master updated successfully";
			
			String actMessage1=checkValidationMessage(expMessage1);
			
			System.out.println("********************************checkExtraFieldCreated***************************");
			
			System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			System.out.println("Message     : "+actMessage1+" Value Expected : "+expMessage1);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_CloseBtn));
			extraFields_CloseBtn.click();
			
			if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) )
			{
				
				return true;
			}
			else
			{
			
				return false;
			}

		}

		@FindBy(xpath="//*[@id='newMasterDiv0']//div//label")
		public static List<WebElement> detailsList;


		public boolean checkDeletedExtraFieldinGeneralTabofDepartmentMaster() throws InterruptedException
		{
			  Thread.sleep(2000);
			  boolean result=true;
			
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
			  masterNewBtn.click();
			  Thread.sleep(2000);
			  
			  for(WebElement e:detailsList)
			  {
				  
				  if(e.getText().equals("Department Extra Field Update"))
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

		        Thread.sleep(2000);

		        String actUserInfo1=userNameDisplay.getText();

		        System.out.println("User Info  : "+actUserInfo1);

		        System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

		        /*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
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


		public boolean checkLogoutDepartmentPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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










	public DepartmentPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}

	
	public static String ActTime;
	

	public boolean  CreateExtraTabAndExtraFieldForCheckingTimeInDepartment() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException 
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
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
		departmentMenu.click();

		
		Thread.sleep(3000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));
		masterRibbonToExpandOptions.click();
				 
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCustamizemasterBtn));
		masterCustamizemasterBtn.click();
		
		Thread.sleep(3000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTabMenu));
		createTabMenu.click();
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createTab));
		createTab.click();
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(NewTabCaption));
		NewTabCaption.click();
		NewTabCaption.sendKeys("TimeFieldTab");
		NewTabCaption.sendKeys(Keys.TAB); 
		Thread.sleep(2000);	 
		
		ClickUsingJs(NewTab_SaveBut);
		
		String ExpTabAddmsg="Tab Added Successfully";
		String ActTabAddmsg=checkValidationMessage(ExpTabAddmsg);
		
		System.err.println("ActTabAddmsg is " +ActTabAddmsg);
		System.err.println("ExpTabAddmsg is " +ExpTabAddmsg);
		
		Thread.sleep(2500);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(TimeFieldTab));
		TimeFieldTab.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(TimeField_BodyDetailsTab));
		TimeField_BodyDetailsTab.click();
		Thread.sleep(1000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExtraFieldAdd));
		ExtraFieldAdd.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExtraFieldCaption));
		ExtraFieldCaption.click();
		Thread.sleep(1000);
		ExtraFieldCaption.sendKeys("Time1");
		Thread.sleep(1000);
		ExtraFieldCaption.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExtraField_Datatype));
		ExtraField_Datatype.click();
		Thread.sleep(1000);
		
		Select s1=new Select(ExtraField_Datatype);
		s1.selectByValue("5");
		Thread.sleep(1500);
		
		String ActFieldDataType=s1.getFirstSelectedOption().getText();
		String ExpFieldDataType="Time";
		
		System.err.println("ActFieldDataType is : " + ActFieldDataType);
		System.err.println("ExpFieldDataType is : " + ExpFieldDataType);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExtraField_Savebut));
		ExtraField_Savebut.click();
		Thread.sleep(1000);
		
		String ExpFieldAddmsg="Field Added Successfully.";
		String ActFieldAddmsg=checkValidationMessage(ExpFieldAddmsg);
		
		
		System.err.println("ActFieldAddmsg is " +ActFieldAddmsg);
		System.err.println("ExpFieldAddmsg is " +ExpFieldAddmsg);
		
		Thread.sleep(2500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExtraField_Closebut));
		ExtraField_Closebut.click();
		Thread.sleep(1000);
		

		
		if(ActTabAddmsg.equalsIgnoreCase(ExpTabAddmsg) && ActFieldAddmsg.equalsIgnoreCase(ExpFieldAddmsg)
				&& ActFieldDataType.equalsIgnoreCase(ExpFieldDataType))
		{
			return true;
		}
		else
		{
		
			return false;
		}
	}

	
	
	public boolean CheckPresentTimeInCreatedTabDepartment() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	{
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
		departmentMenu.click();
		Thread.sleep(2000);
		
		int count = masterDepartmentGridBodyName.size();

		for (int i = 0; i < count; i++) 
		{
			String data = masterDepartmentGridBodyName.get(i).getText();

			if (data.equalsIgnoreCase("India")) 
			{
				getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		int Count=MasterExtraFieldList.size();
		ArrayList<String> row1=new ArrayList<String>();
		for (int i=0;i<Count;i++)
		{
			String data=MasterExtraFieldList.get(i).getText();
			row1.add(data);
		}
		
		String ActMasterTabList=row1.toString();
		String ExtMasterTabList="[General, TimeFieldTab]";
		
		System.out.println("ActMasterTabList is : " + ActMasterTabList);
		System.out.println("ExtMasterTabList is : " + ExtMasterTabList);
		
		Thread.sleep(2000);
		
		Dep_TimeFieldTab.click();
		Thread.sleep(1000);
		
		Time1_Column.click();
		Thread.sleep(2500);
		
		ActTime=Time1_ColumnText.getAttribute("value");
		String ExpTime=getcurrentSystemTime();
		
		System.err.println("ActTime is : " +ActTime);
		System.err.println("ExpTime is : " +ExpTime);
		
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();

		String expMessage="Updated Successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		if(ActTime.equalsIgnoreCase(ExpTime) && actMessage.equalsIgnoreCase(expMessage)
				&& ActMasterTabList.equalsIgnoreCase(ExtMasterTabList))
		{
		return true;
		}
		else
		{
			return false;
			
		}
	}


	public boolean CheckEditingTimeInCreatedTabDepartment() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	{
		Thread.sleep(2500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
		departmentMenu.click();
		Thread.sleep(2000);
		
		int count = masterDepartmentGridBodyName.size();

		for (int i = 0; i < count; i++) 
		{
			String data = masterDepartmentGridBodyName.get(i).getText();

			if (data.equalsIgnoreCase("India")) 
			{
				getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		int Count=MasterExtraFieldList.size();
		ArrayList<String> row1=new ArrayList<String>();
		for (int i=0;i<Count;i++)
		{
			String data=MasterExtraFieldList.get(i).getText();
			row1.add(data);
		}
		
		String ActMasterTabList=row1.toString();
		String ExtMasterTabList="[General, TimeFieldTab]";
		
		System.out.println("ActMasterTabList is : " + ActMasterTabList);
		System.out.println("ExtMasterTabList is : " + ExtMasterTabList);
		
		
		Thread.sleep(2000);
		
		Dep_TimeFieldTab.click();
		Thread.sleep(1000);
		Time1_Column.click();
		Thread.sleep(1500);
		
		String actTime1=Time1_ColumnText.getAttribute("value");
		String expTime1=ActTime;
		
		System.err.println("actTime1 is : " + actTime1);
		System.err.println("expTime1 is : " + expTime1);
		
		for(int i=0;i<4;i++)
		{
			Time1_ColumnText.sendKeys(Keys.BACK_SPACE);
		}

		Thread.sleep(1500);
		
		Calendar calendar = Calendar.getInstance();
		System.err.println("Current Time " + new SimpleDateFormat("HH"+":"+"mm").format(Calendar.getInstance().getTime()));
		calendar.add(Calendar.MINUTE, 2);
		System.err.println("Edited Time " + new SimpleDateFormat("HH"+":"+"mm").format(calendar.getTime()));
		
		Time1_ColumnText.sendKeys(new SimpleDateFormat("HH"+":"+"mm").format(calendar.getTime()));
		
		Thread.sleep(1500);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();

		String expMessage="Updated Successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		Thread.sleep(2000);
		
		int count1 = masterDepartmentGridBodyName.size();

		for (int i = 0; i < count1; i++) 
		{
			String data = masterDepartmentGridBodyName.get(i).getText();

			if (data.equalsIgnoreCase("India")) 
			{
				getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		Dep_TimeFieldTab.click();
		Thread.sleep(1000);
		Time1_Column.click();
		Thread.sleep(1500);
		
		String ActEditedTime=Time1_ColumnText.getAttribute("value");
		String ExpEditedTime=new SimpleDateFormat("HH"+":"+"mm").format(calendar.getTime());
		
		System.err.println("ActEditedTime is : " + ActEditedTime);
		System.err.println("ExpEditedTime is : " + ExpEditedTime);
		
		if(ActMasterTabList.equalsIgnoreCase(ExtMasterTabList) && actTime1.equalsIgnoreCase(expTime1)
				&& actMessage.equalsIgnoreCase(expMessage) &&  ActEditedTime.equalsIgnoreCase(ExpEditedTime))
		{
		return true;
		}
		else
		{
			return false;
		}
		
		
	}




}
