package com.focus.Pages;

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
    	
    	System.out.println("Save Button is Displayed : "+actual+"  "+expected);
    	
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

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();

		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

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

		System.out.println("Saved Department : "+actaccountNewCreationName+"  Value Expected  "+expaccountNewCreationName);
		
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

		//Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();

		String expMessage=excelReader.getCellData(xlSheetName, 12, 6);

		String actMessage=checkValidationMessage(expMessage);

		System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
		
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actaccountNewCreationName);
		
		

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

		System.err.println("accountNewCreationName  : "+actaccountNewCreationName+" Value Expected : "+expaccountNewCreationName);
		
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

	@FindBy(xpath="//*[@id='viewheading']")
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
			   
		   	Thread.sleep(2000);
		   	
		
		 
		 int accCount = masterGridBodyName.size();
		

		 ArrayList<String> acclist = new ArrayList<String>();

		 		for (int i = 0; i < accCount; i++)
		 		{
		 			 String data = masterGridBodyName.get(i).getText();
		 			acclist.add(data);
		 			
		 		}
		 		
		 		String actDeptList=acclist.toString();
		 		String  expDeptList="[DepartmentST, INDIA, DUBAI, AMERICA, SINGPORE, EUROPE]";
		 		
		 		System.out.println("Actual Departments List:  "+actDeptList);
		 		System.out.println("Expected Departments List:  "+expDeptList);
		 		
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
	      	
	  	if(advanceMasterLabel.getText().equalsIgnoreCase("Advance Master Import/Export----> Department"))
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
		 Thread.sleep(4000);
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
		   
	   	Thread.sleep(2000);
	   	
	
	 
	 int accCount = masterGridBodyName.size();
	

	 ArrayList<String> acclist = new ArrayList<String>();

	 		for (int i = 0; i < accCount; i++)
	 		{
	 			 String data = masterGridBodyName.get(i).getText();
	 			acclist.add(data);
	 			
	 		}
	 		
	 		String actDeptList=acclist.toString();
	 		String  expDeptList="[DepartmentST, INDIA, DUBAI, AMERICA, SINGPORE, EUROPE, DepartmentImport]";
	 		
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
{
	
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	   	homeMenu.click();
	   	
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	   	mastersMenu.click();
	   	
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
		     
	   	departmentMenu.click();
		   
	   	Thread.sleep(2000);
	
	
	 int count = masterGridBodyName.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = masterGridBodyName.get(i).getText();
			
			if (data.equalsIgnoreCase("DepartmentImport")) 
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
  
  System.out.println("********************************checkDeleteDepartmentBeforeCustomization***************************");
		
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

	
/////////////////////Adding Groups and subgroups//////////////////////////

public boolean checkSavingDepartmentGroupinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
departmentMenu.click();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
masterAddGroupBtn.click();

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

@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[10]")
private static List<WebElement> accNameList;


@FindBy(xpath="//*[@id='ol_treeNavigation']/li[2]/span")
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

Thread.sleep(2000);


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


String expMessage="Saved Successfully";

String actMessage=checkValidationMessage(expMessage);

System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);

Thread.sleep(2000);

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


@FindBy(xpath="//*[@id='lblnchecked']/input")
public static List<WebElement> accChkBoxList;


public boolean checkSavingDepartmentsinSubGroupsinDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{

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


public static boolean checkDepartmentsSavedinGroupsandSubGroupsinDepartmentMaster() throws InterruptedException
{
getDriver().navigate().refresh();
Thread.sleep(2000);

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
departmentMenu.click();


int accCount = accNameList.size();

ArrayList<String> acclist = new ArrayList<String>();

for (int i = 0; i < accCount; i++)
{
String data = accNameList.get(i).getText();
acclist.add(data);

}

String actAccList=acclist.toString();
String  expAccList="[Group Department, DepartmentST, INDIA, DUBAI, AMERICA, SINGPORE, EUROPE]";


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



public boolean checkDeleteGroupDepartmentinDepartmentMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException		

{
getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
homeMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
mastersMenu.click();

getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
departmentMenu.click();
Thread.sleep(2000);

int count1 = masterGridBodyName.size();

for (int i = 0; i < count1; i++) 
{
String data = masterGridBodyName.get(i).getText();

if (data.equalsIgnoreCase("Group Department")) 
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


	
	
	








	public DepartmentPage(WebDriver driver) 
	{
		PageFactory.initElements(driver, this);
	}




}
