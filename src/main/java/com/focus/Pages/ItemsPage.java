package com.focus.Pages;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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

public class ItemsPage extends BaseEngine
{
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	private static String xlSheetName = "ItemsPage";

			public static void checkPopUpWindow() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
			{
			        try 
			        {
						getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(doNotShowCheckbox));
						doNotShowCheckbox.click();
						
						getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtnInDemoPopupScreen));
						closeBtnInDemoPopupScreen.click(); 
						
						System.err.println("POP UP DISPLAYED AND CLOSED SUCCESSFULLY");
						
					} 
			        catch (Exception e)
			        {
			        	System.err.println("NO POP UP DISPLAYED");
					}
			    	
			}
					 
					
			
			
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
			
			
			
		
		
		
		
		
		
		public  boolean checkItemMasterInMasterMenu() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	    {
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			System.out.println("***********checkItemMasterInMasterMenu Method Excutes **********");
			
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	      	homeMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	      	mastersMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
	    	homeMasterItem_ItemMenu.click();
	    	
	    	Thread.sleep(2000);
	    	
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
		
		
		
		
		
		public static boolean checkSavingItemInItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			System.out.println("********************************checkSavingItemInItemMaster********************");
			 
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
			 
			 //Thread.sleep(2000);
			 
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			 saveBtn.click();
			 
			 String expMessage=excelReader.getCellData(xlSheetName, 9, 6);
				
			 String actMessage=checkValidationMessage(expMessage);
			 
			 System.out.println("actMessage  : "+actMessage);
			 System.out.println("expMessage  : "+expMessage);
			 
			 
			 excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessage);
			 
			 Thread.sleep(8000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
			 masterCloseBtn.click();
				
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
		
		
		
		
		
		
		public static boolean checkEditingInSavedItemINItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			 System.out.println("********************************checkEditingInSavedItemINItemMaster*******************");
			 
			 Thread.sleep(8000);
			 
			 int count = masterItemGridBodyName.size();
				
			for (int i = 0; i < count; i++) 
			{
				String data = masterItemGridBodyName.get(i).getText();
				
				if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 11, 5))) 
				{
					getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
					break;
				}
			}
			    
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
			 
			 String expMessage=excelReader.getCellData(xlSheetName, 11, 6);
				
			 String actMessage=checkValidationMessage(expMessage);
			 
			 System.out.println("actMessage  : "+actMessage);
			 System.out.println("expMessage  : "+expMessage);
			 
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
		
		
		
		
		
		
	
		public static boolean checkUpdatedItemMasterDisplayInItemGridBeforeCustomization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			String actaccountNewCreationName  = null;
			
			 int count = masterItemGridBodyName.size();
				
				for (int i = 0; i < count; i++) 
				{
					String data = masterItemGridBodyName.get(i).getText();
					
					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 14, 5))) 
					{
						actaccountNewCreationName = data;
						break;
					}
				}
				    
			 String expaccountNewCreationName  =excelReader.getCellData(xlSheetName, 14, 6);
			 
			 System.out.println("************************checkUpdatedItemDisplayInItemMasterGridBeforeCustomization***********************");
			 
			 System.err.println("actaccountNewCreationName  : "+actaccountNewCreationName);
			 System.err.println("expaccountNewCreationName  : "+expaccountNewCreationName);
			 
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
	
		
		
		
		public static boolean checkDeleteItemInItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			
			 int count = masterItemGridBodyName.size();
				
				for (int i = 0; i < count; i++) 
				{
					String data = masterItemGridBodyName.get(i).getText();
					
					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 15, 5))) 
					{
						if (masterGridBodyChkbox.get(i).isSelected()==false) 
						{
							masterGridBodyChkbox.get(i).click();
							break;
						}
					}
				}
			
				
	         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDeleteBtn));
	         masterDeleteBtn.click();
	         
	         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getMsgOnDelete));
	         System.out.println(getMsgOnDelete.getText());
	         
	         //Thread.sleep(2000);
	         
	         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
	         clickOnOkInDelete.click();
	         
	         String expMessage=excelReader.getCellData(xlSheetName, 15, 6);
	 		
	         String actMessage=checkValidationMessage(expMessage);
	         
	         System.out.println("********************************checkDeleteItemInItemMaster***************************");
				
	         System.out.println("actMessage  : "+actMessage);
			 System.out.println("expMessage  : "+expMessage);
			 
			 excelReader.setCellData(xlfile, xlSheetName, 15, 7, actMessage);
				
			 if(actMessage.equalsIgnoreCase(expMessage))
			 {
				 /*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
				 closeBtn.click();
				*/ excelReader.setCellData(xlfile, xlSheetName, 15, 8, resPass);
				 return true;
			 }
			 
			 else
			 {
				/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
				 closeBtn.click();
				*/ excelReader.setCellData(xlfile, xlSheetName, 15, 8, resFail);
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
			

		@FindBy(xpath="//*[@id='divAdvMasterImportExport']/ul/li[2]/span/i")
		public static WebElement advMasterCloseBtn;
						 
	



		@FindBy(xpath="//*[@id='btnImportDataForAdvanceMasterPopUp']")
		public static WebElement advMasterImportDataBtn;;


		@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[12]")
		private static List<WebElement> accNameList;

		
		public String actAccList=null;
		public String expAccList=null;
		
		
		public boolean checkItemListBeforeImportingDataFromExcel() throws InterruptedException
		{
			
			
		  	Thread.sleep(3000);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			   	homeMenu.click();
			   	Thread.sleep(1000);
			   	
			   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
			   	mastersMenu.click();
			   	Thread.sleep(1000);
			   	
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		      	homeMasterItemMenu.click();
		      	Thread.sleep(1000);
		      	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
		    	homeMasterItem_ItemMenu.click();
		    	
				   
			   	Thread.sleep(3000);
			   	
			
			 
			 int itemCount = masterItemGridBodyName.size();
			

			 ArrayList<String> itemlist = new ArrayList<String>();

			 		for (int i = 0; i < itemCount; i++)
			 		{
			 			 String data = masterItemGridBodyName.get(i).getText();
			 			itemlist.add(data);
			 			
			 		}
			 		
			 		String actItemList=itemlist.toString();
			 		String  expItemList="[Item1, Item3, Item2, ITEMS GROUP, STOCK ITEM, BATCH WA ITEM, BATCH IGNORE EXP LIFO ITEM, BATCH FIFO ITEM, BATCH BR ITEM, BIN RAW MATERIAL ITEM, BIN FINISHED GOODS ITEM, BATCH BIN FINISHED GOODS ITEM, BATCH BIN WITH NO STOCK UPDATE ITEM, BATCH BIN WITH NO RESERVATION ITEM, RMA RITEM, RMA BATCH ITEM, WA COGS ITEM, FIFO COGS ITEM, Paste Item, Clone Item, BR COGS ITEM, STD RATE COGS ITEM]";
			 		
			 		System.out.println("Actual Item List:  "+actItemList);
			 		System.out.println("Expected Item List:  "+expItemList);
			 		
			 		if(actItemList.equalsIgnoreCase(expItemList))
			 		{
			 			return true;
			 		}
			 		else
			 		{
			 			return false;
			 		}
			 		
		}

		public boolean checkAdvanceMasterImportExportItem() throws InterruptedException
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
			 Thread.sleep(4000);
			 
			 String expMessage="Please load excel sheet";
			String actMessage=checkValidationMessage(expMessage);
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
			FileUploadBtn.click();
			Thread.sleep(4000);
			
			 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\ItemImportFromExcel.exe");
			 Thread.sleep(20000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterImportDataBtn));
			 advMasterImportDataBtn.click();
			 Thread.sleep(15000);
			 
			/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterCloseBtn));
			 advMasterCloseBtn.click();
			 Thread.sleep(2000);*/
			 
			 
			 String expMessage1="20 Records Imported Successfully.";
			 String actMessage1=checkValidationMessage(expMessage1);
			 
			 
		
			 
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

		public boolean checkItemListAfterImportingDataFromExcel() throws InterruptedException
		{
			Thread.sleep(4000);
			 	 
			
			
			 int itemCount = masterItemGridBodyName.size();
				

			 ArrayList<String> itemlist = new ArrayList<String>();

			 		for (int i = 0; i < itemCount; i++)
			 		{
			 			 String data = masterItemGridBodyName.get(i).getText();
			 			itemlist.add(data);
			 			
			 		}
			 		
			 		String actItemList=itemlist.toString();
			 		String  expItemList="[Item1, Item3, Item2, ITEMS GROUP, STOCK ITEM, BATCH WA ITEM, BATCH IGNORE EXP LIFO ITEM, BATCH FIFO ITEM, BATCH BR ITEM, BIN RAW MATERIAL ITEM, BIN FINISHED GOODS ITEM, BATCH BIN FINISHED GOODS ITEM, BATCH BIN WITH NO STOCK UPDATE ITEM, BATCH BIN WITH NO RESERVATION ITEM, RMA RITEM, RMA BATCH ITEM, WA COGS ITEM, FIFO COGS ITEM, Paste Item, Clone Item, BR COGS ITEM, STD RATE COGS ITEM, Item Import]";
			 		
			 		System.out.println("Actual Item List:  "+actItemList);
			 		System.out.println("Expected Item List:  "+expItemList);
			 		
			 		if(actItemList.equalsIgnoreCase(expItemList))
			 		{
			 			return true;
			 		}
			 		else
			 		{
			 			return false;
			 		}
			
			 		
			 		
		}

			
		public boolean checkImportedItemDelete() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			 int count = masterItemGridBodyName.size();
			 
			 for (int i = 0; i < count; i++) 
				{
					String data = masterItemGridBodyName.get(i).getText();
					
					if (data.equalsIgnoreCase("BR COGS ITEM")) 
					{
						if (masterGridBodyChkbox.get(i).isSelected()==true) 
						{
							masterGridBodyChkbox.get(i).click();
							break;
						}
					}
				}
			    
			    Thread.sleep(2000);	
				
				for (int i = 0; i < count; i++) 
				{
					String data = masterItemGridBodyName.get(i).getText();
					
					if (data.equalsIgnoreCase("Item Import")) 
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

		public boolean checkSavingItemGroupinItemMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{

			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		    homeMenu.click();
		    
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		    mastersMenu.click();
		      	
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		    homeMasterItemMenu.click();
		      	
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
		    homeMasterItem_ItemMenu.click();
		    Thread.sleep(4000);
		    
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
		    masterAddGroupBtn.click();
		    Thread.sleep(2000);
		    
			System.out.println("********************************checkSavingItemGroupItemMaster*****************************");
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys("Group Item");
			nameTxt.sendKeys(Keys.TAB);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			codeTxt.click();
			codeTxt.clear();
			codeTxt.sendKeys("GroupItem");
			codeTxt.sendKeys(Keys.TAB);
			 
		
			 
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			Thread.sleep(1000);
			
			String expMessage="Saved Successfully";
				
			String actMessage=checkValidationMessage(expMessage);
			 
			System.out.println("actMessage  : "+actMessage);
			System.out.println("expMessage  : "+expMessage);
			
			Thread.sleep(3000);
			
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




		@FindBy(xpath="//*[@id='ol_treeNavigation']/li[2]")
		public static WebElement GroupAccountLabel;

		public static boolean checkSavingItemSubGroupinItemMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			Thread.sleep(3000);
			
			int accNameListCount = accNameList.size();

			for(int i=0;i<accNameListCount;i++)
			{
				String data = accNameList.get(i).getText();
				Thread.sleep(2000);
				if(data.equals("Group Item"))
				{
					Thread.sleep(2000);
					getAction().doubleClick(accChkBoxList.get(i)).build().perform();
					Thread.sleep(2000);
					break;
				}
			}
			
			Thread.sleep(6000);
			    
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
			    masterAddGroupBtn.click();
				Thread.sleep(4000);
				 
				System.out.println("********************************checkSavingItemSubGroupItemMaster*****************************");
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
				nameTxt.click();
				nameTxt.clear();
				nameTxt.sendKeys("SubGroup1 Item");
				nameTxt.sendKeys(Keys.TAB);
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
				codeTxt.click();
				codeTxt.clear();
				codeTxt.sendKeys("SubGroup1Item");
				codeTxt.sendKeys(Keys.TAB);
				 
			
				 
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				Thread.sleep(3000);
				
				String expMessage1="Saved Successfully";
				
				String actMessage1=checkValidationMessage(expMessage1);
				 
				System.out.println("actMessage1  : "+actMessage1);
				System.out.println("expMessage1  : "+expMessage1);
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
				nameTxt.click();
				nameTxt.clear();
				nameTxt.sendKeys("SubGroup2 Item");
				nameTxt.sendKeys(Keys.TAB);
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
				codeTxt.click();
				codeTxt.clear();
				codeTxt.sendKeys("SubGroup2Item");
				codeTxt.sendKeys(Keys.TAB);
				 
				
				 
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				 
				String expMessage2="Saved Successfully";
					
				String actMessage2=checkValidationMessage(expMessage2);
				 
				System.out.println("actMessage2  : "+actMessage2);
				System.out.println("expMessage2  : "+expMessage2);
				
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
				masterCloseBtn.click();
					
				if(actMessage1.equalsIgnoreCase(expMessage1) && actMessage2.equalsIgnoreCase(expMessage2) )
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
			
			boolean ActAccountGrp=accountGrp.isDisplayed();
			boolean ExpAccountGrp=true;
			
			if(ActAccountGrp==ExpAccountGrp)
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
		
		
		public boolean checkSavingItemsinSubGroupsinItemMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			
			Thread.sleep(4000);
			int accNameListCount = accNameList.size();

			for(int i=0;i<accNameListCount;i++)
			{
				String data = accNameList.get(i).getText();
				Thread.sleep(2000);
				if(data.equals("SubGroup1 Item"))
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
			    Thread.sleep(2000);
				System.out.println("********************************checkSavingItemSubGroupItemMaster*****************************");
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
				nameTxt.click();
				nameTxt.clear();
				nameTxt.sendKeys("Item1");
				nameTxt.sendKeys(Keys.TAB);
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
				codeTxt.click();
				codeTxt.clear();
				codeTxt.sendKeys("Item11");
				codeTxt.sendKeys(Keys.TAB);
				 
			
				 
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				Thread.sleep(3500);
				
				String expMessage1="Saved Successfully";
				
				String actMessage1=checkValidationMessage(expMessage1);
				 
				System.out.println("actMessage 1 : "+actMessage1);
				System.out.println("expMessage 1 : "+expMessage1);
				
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
					if(data.equals("SubGroup2 Item"))
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
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
				nameTxt.click();
				nameTxt.clear();
				nameTxt.sendKeys("Item2");
				nameTxt.sendKeys(Keys.TAB);
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
				codeTxt.click();
				codeTxt.clear();
				codeTxt.sendKeys("Item21");
				codeTxt.sendKeys(Keys.TAB);
				 
				
				 
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				 
				String expMessage2="Saved Successfully";	
				String actMessage2=checkValidationMessage(expMessage2);
				 
				System.out.println("actMessage2  : "+actMessage2);
				System.out.println("expMessage2 : "+expMessage2);
				
				Thread.sleep(4000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
				masterCloseBtn.click();
					
				if(actMessage1.equalsIgnoreCase(expMessage1) && actMessage2.equalsIgnoreCase(expMessage2))
				{
					 return true;
				}
				else
				{
					 return false;
				}
		}

		@FindBy(xpath="//ol[@id='ol_treeNavigation']//li[1]")
		public static WebElement itemLabel1;
		
		public static boolean checkItemsSavedinGroupsandSubGroupsinItemMaster() throws InterruptedException
		{
			getDriver().navigate().refresh();
			Thread.sleep(2000);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			    homeMenu.click();
			    
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
			    mastersMenu.click();
			      	
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
			    homeMasterItemMenu.click();
			      	
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
			    homeMasterItem_ItemMenu.click();
			    Thread.sleep(2000);
			
			 Thread.sleep(4000);
			    //getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLabel1))	;
			   // itemLabel1.click();
			   // Thread.sleep(2000);

			   	
			
			   	int accCount = accNameList.size();
			
			   	ArrayList<String> acclist = new ArrayList<String>();

			 		for (int i = 0; i < accCount; i++)
			 		{
			 			 String data = accNameList.get(i).getText();
			 			acclist.add(data);
			 			
			 		}
			 		
			 		String actAccList=acclist.toString();
			 		String  expAccList="[Group Item, Item1, Item3, Item2, ITEMS GROUP, STOCK ITEM, BATCH WA ITEM, BATCH IGNORE EXP LIFO ITEM, BATCH FIFO ITEM, BATCH BR ITEM, BIN RAW MATERIAL ITEM, BIN FINISHED GOODS ITEM, BATCH BIN FINISHED GOODS ITEM, BATCH BIN WITH NO STOCK UPDATE ITEM, BATCH BIN WITH NO RESERVATION ITEM, RMA RITEM, RMA BATCH ITEM, WA COGS ITEM, FIFO COGS ITEM, BR COGS ITEM, STD RATE COGS ITEM]";
			 		

			 		int accNameListCount = accNameList.size();

			 		for(int i=0;i<accNameListCount;i++)
			 		{
			 			String data = accNameList.get(i).getText();
			 			Thread.sleep(2000);
			 			if(data.equals("Group Item"))
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
			 	String  expGrpAccList="[SubGroup2 Item, SubGroup1 Item]";
			 		 		
			 	
			 	int accNameListCount1 = accNameList.size();
		 		for(int i=0;i<accNameListCount1;i++)
		 		{
			 		String data = accNameList.get(i).getText();
			 		Thread.sleep(2000);
			 		if(data.equals("SubGroup1 Item"))
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
			 	String  expSubGrpAccList="[Item1]";
			 		 		 		
			 
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GroupAccountLabel));
				GroupAccountLabel.click();
				Thread.sleep(2000);
			 	
			 	int accNameListCount3 = accNameList.size();
		 		for(int i=0;i<accNameListCount3;i++)
		 		{
			 		String data = accNameList.get(i).getText();
			 		Thread.sleep(2000);
			 		if(data.equals("SubGroup2 Item"))
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
			 	String  expSubGrpAcc2List="[Item2]";
			 		 		 		
			 	System.out.println("********************************checkSavingGroupsandSubGroupItemMaster*****************************");
			 	
			 	
			 	System.out.println("Actual Items List  :  "+actAccList);
		 		System.out.println("Expected Items List:  "+expAccList);
			 	
		 		System.out.println("Actual Sub Group Items List  :  "+actGrpAccList);
			 	System.out.println("Expected Sub Group Items List:  "+expGrpAccList);
			 	
				System.out.println("Actual  Items in Sub Group1 List :  "+actSubGrpAccList);
			 	System.out.println("Expected Items in Sub Group1 List:  "+expSubGrpAccList);
			 	
			 	System.out.println("Actual  Items in Sub Group2 List :  "+actSubGrpAcc2List);
			 	System.out.println("Expected Items in Sub Group2 List:  "+expSubGrpAcc2List);
			 	
			 	
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
					closeBtn.click();
*/			 		return false;
			 	}

		}
	

		public boolean checkDeleteGroupIteminItemMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			Thread.sleep(1500);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			    homeMenu.click();
			    
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
			    mastersMenu.click();
			      	
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
			    homeMasterItemMenu.click();
			      	
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
			    homeMasterItem_ItemMenu.click();
			    Thread.sleep(2000);
			
			 //Thread.sleep(2000);
			    
			   // getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLabel1))	;
			   // itemLabel1.click();
			    //Thread.sleep(2000);
		
			 int count1 = accNameList.size();
				
				for (int i = 0; i < count1; i++) 
				{
					String data = accNameList.get(i).getText();
					
					if (data.equalsIgnoreCase("Group Item")) 
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
			
		System.out.println("actMessage  : "+actMessage);
		System.out.println("expMessage  : "+expMessage);
			
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

///Adding ExtraField
		
		
		public boolean checkCreatingExtraFieldsinItemMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
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
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
	    	homeMasterItem_ItemMenu.click();
	    	
	    	Thread.sleep(2000);
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));
			masterRibbonToExpandOptions.click();
					 
			 Thread.sleep(3000);
			
			ClickUsingJs(masterCustamizemasterBtn);
			
			Thread.sleep(6000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterGeneralField));
			masterGeneralField.click();
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterGeneralHeaderDetails));
			masterGeneralHeaderDetails.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeMasterCreateExtraFieldAddButton));
			customizeMasterCreateExtraFieldAddButton.click();
			Thread.sleep(2000);			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_FieldDetailsCaption));
			extraFields_FieldDetailsCaption.sendKeys("Item Extra Field ");
			extraFields_FieldDetailsCaption.sendKeys(Keys.TAB);
		         
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_OkBtn));
			extraFields_OkBtn.click();
		         
			String expMessage="Field Added Successfully.";
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("actMessage  : "+actMessage);
			System.out.println("expMessage  : "+expMessage);
			
			boolean actextraFields_CreatedByUserInHeaderDetailsTab =extraFields_CreatedByUserInHeaderDetailsTab.isDisplayed();
			boolean expextraFields_CreatedByUserInHeaderDetailsTab =true;
			
			System.out.println("***************************checkCreatedFieldDisplayInCustomizeTab*****************************");
			
			
			System.out.println("actextraFields_CreatedByUserInHeaderDetailsTab : "+actextraFields_CreatedByUserInHeaderDetailsTab);
			System.out.println("expextraFields_CreatedByUserInHeaderDetailsTab : "+expextraFields_CreatedByUserInHeaderDetailsTab);		 
			
			System.err.println("Created Extra field is : "+extraFields_CreatedByUserInHeaderDetailsTab.getText());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_SaveBtn));
			extraFields_SaveBtn.click();
			
			
			String expMessage1="Master updated successfully";
			String actMessage1=checkValidationMessage(expMessage1);
			
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
			
				return false;
			}
		}
		 
		 
			
			
		@FindBy(xpath="//a[text()='Item Extra Field ']")
		public static WebElement createdextraFieldinGeneralTab;
		
		@FindBy(xpath="//a[text()='Item Extra Field Update ']")
		public static WebElement createdUpdatedextraFieldinGeneralTab;
			
			
		public boolean checkAddedExtraFieldinGeneralTabofItemMaster() throws InterruptedException
		{
			Thread.sleep(2000);
			
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	      	homeMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	      	mastersMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
	    	homeMasterItem_ItemMenu.click();
	    	
	    	Thread.sleep(4000);
	    
			
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
			  masterNewBtn.click();
			  Thread.sleep(8000);
			  
			
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdextraFieldinGeneralTab));
			  String actCreatedExtraField=createdextraFieldinGeneralTab.getText();
			  String expCreatedEXtraField="Item Extra Field";
			  
			  System.out.println("actCreatedExtraField  : "+actCreatedExtraField);
			  System.out.println("expCreatedEXtraField  : "+expCreatedEXtraField);
			  
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
			


		@FindBy(xpath="(//*[contains(text(),'Item Extra Field')])[5]/..//i[1]")
		public static WebElement extraFields_EditOptionInHeaderDetailsTab;

		@FindBy(xpath="(//*[contains(text(),'Item Extra Field Update')])[5]/..//i[2]")
		public static WebElement extraFields_DeleteOptionInHeaderDetailsTab;


		public boolean checkEditandUpdateExtraFieldinCustomizeMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));
			masterRibbonToExpandOptions.click();
			Thread.sleep(2000);	 
			
			getAction().moveToElement(masterCustamizemasterBtn).build().perform();
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
			extraFields_FieldDetailsCaption.sendKeys("Item Extra Field Update");
			Thread.sleep(1500);
			extraFields_FieldDetailsCaption.sendKeys(Keys.TAB);	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_OkBtn));
			extraFields_OkBtn.click();
			Thread.sleep(4000);
		         
			String expMessage="Field Updated Successfully";
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("actMessage  : "+actMessage);
			System.out.println("expMessage  : "+expMessage);
			
			Thread.sleep(2000);
			
			
			
			boolean actextraFields_CreatedByUserInHeaderDetailsTab =extraFields_CreatedByUserInHeaderDetailsTab.isDisplayed();
			boolean expextraFields_CreatedByUserInHeaderDetailsTab =true;
			
			System.out.println("***************************checkCreatedFieldDisplayInCustomizeTab*****************************");
			
			
			System.out.println("actextraFields_CreatedByUserInHeaderDetailsTab : "+actextraFields_CreatedByUserInHeaderDetailsTab);
			System.out.println("expextraFields_CreatedByUserInHeaderDetailsTab : "+expextraFields_CreatedByUserInHeaderDetailsTab);
					 
			System.err.println("Updated Extra field is : "+extraFields_CreatedByUserInHeaderDetailsTab.getText());
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_SaveBtn));
			extraFields_SaveBtn.click();
			Thread.sleep(4000);
			
			String expMessage1="Master updated successfully";
			String actMessage1=checkValidationMessage(expMessage1);
			
			System.out.println("actMessage1  : "+actMessage1);
			System.out.println("expMessage1  : "+expMessage1);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_CloseBtn));
			extraFields_CloseBtn.click();
			
			
			if(actMessage.equalsIgnoreCase(expMessage) && actMessage1.equalsIgnoreCase(expMessage1) &&actextraFields_CreatedByUserInHeaderDetailsTab==expextraFields_CreatedByUserInHeaderDetailsTab)
			{
				
				
				return true;
			}
			else
			{
			
				return false;
			}
		}
			
			
		public boolean checkUpdatedExtraFieldinGeneralTabofItemMaster() throws InterruptedException
		{
			 
			
			  Thread.sleep(2000);
				
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		      	homeMenu.click();
		      	
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		      	mastersMenu.click();
		      	
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		      	homeMasterItemMenu.click();
		      	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
		    	homeMasterItem_ItemMenu.click();
		    	
		    	Thread.sleep(4000);
			  
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
			  masterNewBtn.click();
			  Thread.sleep(2000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createdUpdatedextraFieldinGeneralTab));
			  String actCreatedExtraField=createdUpdatedextraFieldinGeneralTab.getText();
			  String expCreatedEXtraField="Item Extra Field Update";
			  
			  System.out.println("actCreatedExtraField   : "+actCreatedExtraField);
			  System.out.println("expCreatedEXtraField   : "+expCreatedEXtraField);
				
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
			

			
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));
			masterRibbonToExpandOptions.click();
			
			Thread.sleep(3000);

			ClickUsingJs(masterCustamizemasterBtn);
			
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
		     
		     String expMessage="Field deleted Successfully";
		     String actMessage=checkValidationMessage(expMessage);
		     
		     System.out.println("actMessage  : "+actMessage);
		     System.out.println("expMessage  : "+expMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(extraFields_SaveBtn));
			extraFields_SaveBtn.click();
			
			
			String expMessage1="Master updated successfully";
			String actMessage1=checkValidationMessage(expMessage1);
			
			System.out.println("actMessage1  : "+actMessage1);
			System.out.println("expMessage1  : "+expMessage1);
			
			
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


		public boolean checkDeletedExtraFieldinGeneralTabofItemMaster() throws InterruptedException
		{
			Thread.sleep(2000);
			
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	      	homeMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	      	mastersMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
	    	homeMasterItem_ItemMenu.click();
	    	
	    	Thread.sleep(4000);
			  boolean result=true;
			
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
			  masterNewBtn.click();
			  Thread.sleep(2000);
			  
			  for(WebElement e:detailsList)
			  {
				  
				  if(e.getText().equals("Item Extra Field Update"))
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

		 public boolean checkEditedItemInCustomizeDisplayColumns() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException 
		  {
			  Thread.sleep(1500);
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		      homeMenu.click();
		      	
		      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		      mastersMenu.click();
		      	
		      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		      homeMasterItemMenu.click();
		      	
		      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
		      homeMasterItem_ItemMenu.click();
		    	
		      Thread.sleep(2000);
		      
		      CustomizeMasterButton.click();
		      
		      Thread.sleep(1500);
		      
		      StandardFields.click();
		      Thread.sleep(1500);
		      Customize_Field.click();
		      Thread.sleep(1500);
		      
		      Select s1=new Select(Customize_Field);
		      s1.selectByValue("sCode");
		      
		      
		      Customize_Field.sendKeys(Keys.TAB);
		  
		      
		      Thread.sleep(2500);
		      
		      String ActField=s1.getFirstSelectedOption().getText();
		      String ExpField="sCode";
		      
		      String ActHeader=Customize_Header.getAttribute("value");
		      String ExpHeader="Code";
		      
		      System.err.println("ActField is : " + ActField);
		      System.err.println("ExpField is : " + ExpField);
		      
		      System.err.println("ActHeader is : " + ActHeader);
		      System.err.println("ExpHeader is: "  + ExpHeader);
		    		  
		      Thread.sleep(2500);
		      
		      Customize_OKBut.click();
		      
		      Thread.sleep(1000);
		      
		      Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\CustomizeCode.exe");
			  
			  Thread.sleep(7000);
			  
			  Customize_OKBut.click();
			  Thread.sleep(1500);
			  
			  accountSearchTextArea.click();
			  Thread.sleep(1000);
			  accountSearchTextArea.sendKeys("RMA RITEM");
			  Thread.sleep(1000);
			  accountSearchTextArea.sendKeys(Keys.ENTER);
			  Thread.sleep(1000);
			  
			  Item_selectCheckbox.click();
			  Thread.sleep(1000);
			  Item_Edit.click();
			  Thread.sleep(1000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			  nameTxt.click();
			  nameTxt.clear();
			  nameTxt.sendKeys("RMA RITEM1");
			  nameTxt.sendKeys(Keys.TAB);
				 
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			  codeTxt.click();
			  codeTxt.clear();
			  codeTxt.sendKeys("RMA RITEM");
			  codeTxt.sendKeys(Keys.TAB);
				 
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			  saveBtn.click();
				 
			  String expMessage="Updated Successfully";		
			  String actMessage=checkValidationMessage(expMessage);
				 
			  System.out.println("actMessage  : "+actMessage);
			  System.out.println("expMessage  : "+expMessage);
			  
			  Item_SearchText.click();
			  Thread.sleep(1000);
			  Item_SearchText.sendKeys("RMA RITEM");
			  Thread.sleep(1000);
			  
				int count=SearchItem_List.size();
				ArrayList<String> row1=new ArrayList<String>();
				for (int i=0;i<count;i++)
				{
					String data=SearchItem_List.get(i).getText();
					row1.add(data);
				}
				
				String exprow1="[RMA RITEM1]";
				String actrow1= row1.toString();
				
				System.out.println("actrow1 : " + actrow1);
				System.out.println("exprow1 : " + exprow1);
				
				settingCustomizeDisplayColumnsasBefore();
			  
			  if(actrow1.equalsIgnoreCase(exprow1))
			  {
				  return true;
			  }
			  else 
			  {
				  return false;
			  }
				
			}
		 
		public static boolean settingCustomizeDisplayColumnsasBefore() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException 
		{
			CustomizeMasterButton.click();
		      
		     Thread.sleep(1500);
		     
		      Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\CustomizeCode.exe");
			  
			  Thread.sleep(7000);
			  
			  Customize_OKBut.click();
			  Thread.sleep(1500);
		     
			  accountSearchTextArea.click();
			  Thread.sleep(1000);
			  accountSearchTextArea.sendKeys("RMA RITEM");
			  Thread.sleep(1000);
			  accountSearchTextArea.sendKeys(Keys.ENTER);
			  Thread.sleep(1000);
			  
			  Item_selectCheckbox.click();
			  Thread.sleep(1000);
			  Item_Edit.click();
			  Thread.sleep(1000);
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			  nameTxt.click();
			  nameTxt.clear();
			  nameTxt.sendKeys("RMA RITEM");
			  nameTxt.sendKeys(Keys.TAB);
				 
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			  codeTxt.click();
			  codeTxt.clear();
			  codeTxt.sendKeys("RMA RITEM");
			  codeTxt.sendKeys(Keys.TAB);
				 
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			  saveBtn.click();
				 
			  String expMessage="Updated Successfully";		
			  String actMessage=checkValidationMessage(expMessage);
				 
			  System.out.println("actMessage  : "+actMessage);
			  System.out.println("expMessage  : "+expMessage);
			  
			  if(actMessage.equalsIgnoreCase(expMessage))
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
		       // companyLogo.click();

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


		public boolean checkLogoutItemsPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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

		 @FindBy(xpath="//span[@id='spnHeaderText']")
	     private static WebElement item_ItemLabel;
		 
		 
		 
		 @FindBy(xpath="//*[@id='divBtnGroup0']")
			public static WebElement generalTab1;
			
			@FindBy(xpath="//*[@id='divBtnGroup1']")
			public static WebElement unitsTab1;
			
			@FindBy(xpath="//*[@id='divBtnGroup2']")
			public static WebElement settingsTab1;
			
			@FindBy(xpath="//*[@id='divBtnGroup3']")
			public static WebElement classificationTab1;
			
			@FindBy(xpath="//*[@id='divBtnGroup4']")
			public static WebElement otherDetailsTab1;
			
			@FindBy(xpath="//*[@id='divBtnGroup5']")
			public static WebElement replenishmentTab1;
			
		
			@FindBy(xpath="//*[@id='iProductType']")
			public static WebElement productTypeSelect;
			
			
			@FindBy(xpath="//*[@id='fReorderLevel']")
			public static WebElement reorderLevel;
			
			@FindBy(xpath="//*[@id='iBinCapacity']")
			public static WebElement binCapacity;
			
			@FindBy(xpath="//*[@id='iValuationMethod']")
			public static WebElement valuationMethodSelect;
			
			@FindBy(xpath="//*[@id='iProductMake']")
			public static WebElement itemMakeSelect;
			
			
			@FindBy(xpath="//*[@id='btnMasterCopy']")
			public static WebElement itemCopyBtn;
					
			
			@FindBy(xpath="//*[@id='btnMasterPaste']")
			public static WebElement itemPasteBtn;
			
			//Units tab
			
			@FindBy(xpath="//*[@id='iDefaultBaseUnit']")
			public static WebElement itemDefaultBaseUnitTxt;
			
			@FindBy(xpath="//*[@id='iDefaultSalesUnit']")
			public static WebElement itemDefaultSalesUnitTxt;
			
			@FindBy(xpath="//*[@id='iDefaultPurchaseUnit']")
			public static WebElement itemDefaultPurchaseUnitTxt;
			
			@FindBy(xpath="//*[@id='fLength']")
			public static WebElement itemLengthUnitTxt;
			
			@FindBy(xpath="//*[@id='fWidth']")
			public static WebElement itemWindthUnitTxt;
			
			@FindBy(xpath="//*[@id='fHeight']")
			public static WebElement itemHeightUnitTxt;
			
			@FindBy(xpath="//*[@id='fWeight']")
			public static WebElement itemWeightUnitTxt;
			
			
			//Settings Tab
			
			@FindBy(xpath="//*[@id='fStandardCost']")
			public static WebElement itemStandardCostTxt;
			
			@FindBy(xpath="//*[@id='fOverheadCost']")
			public static WebElement itemOverheadCostTxt;
			
			@FindBy(xpath="//*[@id='fIndirectCost']")
			public static WebElement itemIndirectCostTxt;
			
			@FindBy(xpath="//*[@id='fProfit']")
			public static WebElement itemProfitTxt;
			
			@FindBy(xpath="//*[@id='iDontshowproductexpireddays']")
			public static WebElement itemExpiredDaysTxt;
			
			@FindBy(xpath="//*[@id='iTaxCode']")
			public static WebElement itemTaxCodeTxt;
			
			@FindBy(xpath="//*[@id='TaxCategory']")
			public static WebElement itemSelectTaxCategoryTxt;
			
			//Classification
			
			@FindBy(xpath="//*[@id='iABCCodeValue']")
			public static WebElement itemSelectABCCodeValueTxt;
			
			@FindBy(xpath="//*[@id='iABCCodeMargin']")
			public static WebElement itemSelectABCCodeMarginTxt;
			
			@FindBy(xpath="//*[@id='iABCCodeRevenue']")
			public static WebElement itemSelectPortalRightsTxt;
			
			@FindBy(xpath="//*[@id='iABCCodeCarryingCost']")
			public static WebElement itemSelectABCCodeCarryingTxt;
			
			@FindBy(xpath="//*[@id='sBarcode1']")
			public static WebElement itemBarcode1Txt;
			
			@FindBy(xpath="//*[@id='sBarcode2']")
			public static WebElement itemBarcode2Txt;
			
			@FindBy(xpath="//*[@id='sBarcode3']")
			public static WebElement itemBarcode3Txt;
			
			//Other Details
			
			@FindBy(xpath="//*[@id='iCostOfIssueAccount']")
			public static WebElement itemCostofIssueAccountTxt;
			
			@FindBy(xpath="//*[@id='iStocksAccount']")
			public static WebElement itemStocksAccountTxt;
			
			@FindBy(xpath="//*[@id='iSalesAccount']")
			public static WebElement itemSalesAccountTxt;
			
			@FindBy(xpath="//*[@id='iWIPAccount']")
			public static WebElement itemWIPAccountTxt;
			
			@FindBy(xpath="//*[@id='iCostofShortageStockAC']")
			public static WebElement itemCostofShortageStockAccTxt;
			
			@FindBy(xpath="//*[@id='iCostofExcessStockAC']")
			public static WebElement itemCostofExcessStockAccTxt;
			
			@FindBy(xpath="//*[@id='iCostofSaleReturnAC']")
			public static WebElement itemCostofSalesReturnAccTxt;
			

			@FindBy(xpath="//*[@id='iPurchaseVarianceAC']")
			public static WebElement itemPurchaseVarianceAccTxt;
			
			@FindBy(xpath="//*[@id='VATExpenseAccount']")
			public static WebElement itemVATExpenseAccTxt;
			
			@FindBy(xpath="//*[@id='iProductionLeadTime']")
			public static WebElement itemProductionLeadTimeTxt;
			
			//Replenishment
			
			@FindBy(xpath="//*[@id='iDefaultReplenishment']")
			public static WebElement itemSelectDefaultReplenidshmentTxt;
			
			@FindBy(xpath="//*[@id='iManufacturePolicy']")
			public static WebElement itemSelectManufacturePolicyTxt;
			

			@FindBy(xpath="//*[@id='iPackingBOM']")
			public static WebElement itemPackingBOMTxt;
			
			@FindBy(xpath="//*[@id='iBOM']")
			public static WebElement itemBOMTxt;
			
			@FindBy(xpath="//*[@id='iFlushing']")
			public static WebElement itemSelectFlushingTxt;
			

		public boolean checkCopyandPasteofAnItem() throws InterruptedException
		{
			getDriver().navigate().refresh();
			Thread.sleep(2000);
			
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	      	homeMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	      	mastersMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
	    	homeMasterItem_ItemMenu.click();
	    	
			
			Thread.sleep(2000);
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_ItemLabel));
		      	
		int accNameListCount = accNameList.size();

		for(int i=0;i<accNameListCount;i++)
		{
			String data = accNameList.get(i).getText();
			Thread.sleep(2000);
			if(data.equals("BR COGS ITEM"))
			{
				Thread.sleep(2000);
				getAction().doubleClick(accChkBoxList.get(i)).build().perform();
				Thread.sleep(2000);
				break;
			}
		}
		
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalTab1));
		
		
		System.out.println("Copying the item");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCopyBtn));
		itemCopyBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterCloseBtn));
		masterCloseBtn.click();
		
		Thread.sleep(4000);
		System.out.println("Closing the item");
		
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		 masterNewBtn.click();
		 Thread.sleep(2000);
		 
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemPasteBtn));
		 itemPasteBtn.click();
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalTab1));
		 
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		String actItemName=nameTxt.getAttribute("value");
		String expItemName="BR COGS ITEM";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		String actItemCode=codeTxt.getAttribute("value");
		String expItemCode="BR COGS ITEM";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(productTypeSelect));
		Select s1=new Select(productTypeSelect);
		String actProductType=s1.getFirstSelectedOption().getText();
		String expProductType="Raw material";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reorderLevel));
		String actReorderLevelTxt=reorderLevel.getAttribute("value");
		String expReorderLevelTxt="0.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binCapacity));
		String actBinCapacity=binCapacity.getAttribute("value");
		String expBinCapacity="1.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valuationMethodSelect));
		Select s2=new Select(valuationMethodSelect);
		String actValuationMethod=s2.getFirstSelectedOption().getText();
		String expValuationMethod="Batch Rate";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemMakeSelect));
		Select s3=new Select(itemMakeSelect);
		String actProductMake=s3.getFirstSelectedOption().getText();
		String expProductMake="Item";
		
		
		
		System.err.println("Item Text in Details Tab");
		System.err.println("Name	  		:" +	"	Actual " +actItemName		 +	"	Expected "	+expItemName);
		System.err.println("Code			:" +	"	Actual " +actItemCode		 +	"	Expected "	+expItemCode);
		System.err.println("Item Type		:" +	"	Actual " +actProductType	 +	"	Expected "	+expProductType);
		System.err.println("Reorder Level	:" +	"	Actual " +actReorderLevelTxt +	"	Expected "	+expReorderLevelTxt);
		System.err.println("Bin Capacity	:" +	"	Actual " +actBinCapacity	 +	"	Expected "	+expBinCapacity);
		System.err.println("Valuation Method:" +	"	Actual " +actValuationMethod +	"	Expected "	+expValuationMethod);
		System.err.println("Product Make	:" +	"	Actual " +actProductMake	 +	"	Expected "	+expProductMake);
		if(actItemName.equalsIgnoreCase(expItemName) && actItemCode.equalsIgnoreCase(expItemCode) && actProductType.equalsIgnoreCase(expProductType) &&actReorderLevelTxt.equalsIgnoreCase(expReorderLevelTxt)
				&& actBinCapacity.equalsIgnoreCase(expBinCapacity) && actValuationMethod.equalsIgnoreCase(expValuationMethod) && actProductMake.equalsIgnoreCase(expProductMake))
		
		{
		
		return true;
		}
		else
		{
			return false;
		}
		}
		
	
		
	public boolean checkUnitsTabDataAfterPasteItem() throws InterruptedException
	{
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitsTab1));
		 unitsTab1.click();
		 Thread.sleep(3000);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultBaseUnitTxt));
		String actDefaultBase=itemDefaultBaseUnitTxt.getAttribute("value");
		String expDefaultBase="Pcs";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultSalesUnitTxt));
		String actDefaultSales=itemDefaultSalesUnitTxt.getAttribute("value");
		String expDefaultSales="Pcs";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultPurchaseUnitTxt));
		String actDefaultPurchase=itemDefaultPurchaseUnitTxt.getAttribute("value");
		String expDefaultPurchase="Dozs";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLengthUnitTxt));
		String actLenghtTxt=itemLengthUnitTxt.getAttribute("value");
		String expLenghtTxt="0.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWindthUnitTxt));
		String actWidthTxt=itemWindthUnitTxt.getAttribute("value");
		String expWidthTxt="0.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemHeightUnitTxt));
		String actHeightTxt=itemHeightUnitTxt.getAttribute("value");
		String expHeightTxt="0.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWeightUnitTxt));
		String actWeightTxt=itemWeightUnitTxt.getAttribute("value");
		String expWeightTxt="0.00";
		
		
		
		System.err.println("Item Text in Units Tab");
		System.err.println("*******************************************************************************************************************************************");
		System.err.println("Default Base Unit		"		+		"	Actual		"	+	actDefaultBase			+		"	Expected		"	+	expDefaultBase);
		System.err.println("Deafult Sales Unit		"		+		"	Actual		"	+	actDefaultSales			+		"	Expected		"	+	expDefaultSales);
		System.err.println("Default Purchase Unit	"		+		"	Actual		"	+	actDefaultPurchase		+		"	Expected		"	+	expDefaultPurchase);
		System.err.println("Length					"		+		"	Actual		"	+	actLenghtTxt			+		"	Expected		"	+	expLenghtTxt);
		System.err.println("Width					"		+		"	Actual		"	+	actWidthTxt				+		"	Expected		"	+	expWidthTxt);
		System.err.println("Height					"		+		"	Actual		"	+	actHeightTxt			+		"	Expected		"	+	expHeightTxt);
		System.err.println("Weight					"		+		"	Actual		"	+	actWeightTxt			+		"	Expected		"	+	expWeightTxt);
		
		
		if(actDefaultBase.equalsIgnoreCase(expDefaultBase) && actDefaultSales.equalsIgnoreCase(expDefaultSales) && actDefaultPurchase.equalsIgnoreCase(expDefaultPurchase)
				&& actLenghtTxt.equalsIgnoreCase(expLenghtTxt) && actWidthTxt.equalsIgnoreCase(expWidthTxt) && actHeightTxt.equalsIgnoreCase(expHeightTxt)
				&& actWeightTxt.equalsIgnoreCase(expWeightTxt))
		{
			return true;
		}
		
		else
		{
			return false;
		}
		
		
	}
		
		
	public boolean checkSettingsTabDataAfterPasteItem() throws InterruptedException
	{
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsTab1));
		 settingsTab1.click();
		 Thread.sleep(3000);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemStandardCostTxt));
		String actStandardCostTxt=itemStandardCostTxt.getAttribute("value");
		String expStandardCostTxt="0.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemOverheadCostTxt));
		String actOverheadCostTxt=itemOverheadCostTxt.getAttribute("value");
		String expOverheadCostTxt="0.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemIndirectCostTxt));
		String actIndirectCostTxt=itemIndirectCostTxt.getAttribute("value");
		String expIndirectCostTxt="0.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemProfitTxt));
		String actProfitTxt=itemProfitTxt.getAttribute("value");
		String expProfitTxt="0.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemExpiredDaysTxt));
		String actExpiredDaysTxt=itemExpiredDaysTxt.getAttribute("value");
		String expExpiredDaysTxt="0";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemTaxCodeTxt));
		String actTaxCodeTxt=Boolean.toString(itemTaxCodeTxt.getAttribute("value").isEmpty());
		String expTaxCodeTxt="True";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectTaxCategoryTxt));
		Select s=new Select(itemSelectTaxCategoryTxt);
		String actTaxCategoryTxt=s.getFirstSelectedOption().getText();
		String expTaxCategoryTxt="Taxable";
		
		
		
		System.err.println("Item Text in Settings Tab");
		System.err.println("*******************************************************************************************************************************************");
		System.err.println("Standard Cost 				"		+		"	Actual		"	+	actStandardCostTxt			+		"	Expected		"	+	expStandardCostTxt);
		System.err.println("Overhead Cost				"		+		"	Actual		"	+	actOverheadCostTxt			+		"	Expected		"	+	expOverheadCostTxt);
		System.err.println("Indirect Cost %				"		+		"	Actual		"	+	actIndirectCostTxt			+		"	Expected		"	+	expIndirectCostTxt);
		System.err.println("Profit %					"		+		"	Actual		"	+	actProfitTxt				+		"	Expected		"	+	expProfitTxt);
		System.err.println("Expiring Items				"		+		"	Actual		"	+	actExpiredDaysTxt			+		"	Expected		"	+	expExpiredDaysTxt);
		System.err.println("Tax Code					"		+		"	Actual		"	+	actTaxCodeTxt				+		"	Expected		"	+	expTaxCodeTxt);
		System.err.println("Tax Category				"		+		"	Actual		"	+	actTaxCategoryTxt			+		"	Expected		"	+	expTaxCategoryTxt);
	
		
		if(actStandardCostTxt.equalsIgnoreCase(expStandardCostTxt) &&  actOverheadCostTxt.equalsIgnoreCase(expOverheadCostTxt) && actIndirectCostTxt.equalsIgnoreCase(expIndirectCostTxt)
			&& actProfitTxt.equalsIgnoreCase(expProfitTxt) && actExpiredDaysTxt.equalsIgnoreCase(expExpiredDaysTxt) && actTaxCodeTxt.equalsIgnoreCase(expTaxCodeTxt)
			&& actTaxCategoryTxt.equalsIgnoreCase(expTaxCategoryTxt))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	
	public boolean checkClasifictationTabDataAfterPasteItem() throws InterruptedException
	{
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(classificationTab1));
		 classificationTab1.click();
		 Thread.sleep(3000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectABCCodeValueTxt));
			Select s=new Select(itemSelectABCCodeValueTxt);
			String actABCCodeTxt=s.getFirstSelectedOption().getText();
			String expABCCodeTxt="A";
		
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectABCCodeMarginTxt));
			Select s1=new Select(itemSelectABCCodeMarginTxt);
			String actABCMarginTxt=s1.getFirstSelectedOption().getText();
			String expABCMarginTxt="A";
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectPortalRightsTxt));
			Select s2=new Select(itemSelectPortalRightsTxt);
			String actPortalRightsTxt=s2.getFirstSelectedOption().getText();
			String expPortalRightsTxt="A";
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectABCCodeCarryingTxt));
			Select s3=new Select(itemSelectABCCodeCarryingTxt);
			String actABCCarryingTxt=s3.getFirstSelectedOption().getText();
			String expABCCarryingTxt="A";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBarcode1Txt));
		String actBarcode1Txt=Boolean.toString(itemBarcode1Txt.getAttribute("value").isEmpty());
		String expBarcode1Txt="True";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBarcode2Txt));
		String actBarcode2Txt=Boolean.toString(itemBarcode2Txt.getAttribute("value").isEmpty());
		String expBarcode2Txt="True";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBarcode3Txt));
		String actBarcode3Txt=Boolean.toString(itemBarcode3Txt.getAttribute("value").isEmpty());
		String expBarcode3Txt="True";
		
		
		
		System.err.println("Item Text in Classification Tab");
		System.err.println("*******************************************************************************************************************************************");
		System.err.println("ABC-Code Value			"		+		"	Actual		"	+	actABCCodeTxt					+		"	Expected		"	+	expABCCodeTxt);
		System.err.println("ABC-Code Margin			"		+		"	Actual		"	+	actABCMarginTxt					+		"	Expected		"	+	expABCMarginTxt);
		System.err.println("Portal Rights			"		+		"	Actual		"	+	actPortalRightsTxt				+		"	Expected		"	+	expPortalRightsTxt);
		System.err.println("ABCCarryingCost			"		+		"	Actual		"	+	actABCCarryingTxt				+		"	Expected		"	+	expABCCarryingTxt);
		System.err.println("Barcode1				"		+		"	Actual		"	+	actBarcode1Txt					+		"	Expected		"	+	expBarcode1Txt);
		System.err.println("Barcode2				"		+		"	Actual		"	+	actBarcode2Txt					+		"	Expected		"	+	expBarcode2Txt);
		System.err.println("Barcode3				"		+		"	Actual		"	+	actBarcode3Txt					+		"	Expected		"	+	expBarcode3Txt);
	
		if(actABCCodeTxt.equalsIgnoreCase(expABCCodeTxt) && actABCMarginTxt.equalsIgnoreCase(expABCMarginTxt) && actABCCarryingTxt.equalsIgnoreCase(expABCCarryingTxt)
				&& actPortalRightsTxt.equalsIgnoreCase(expPortalRightsTxt) && actBarcode1Txt.equalsIgnoreCase(expBarcode1Txt) && actBarcode2Txt.equalsIgnoreCase(expBarcode2Txt)
				&& actBarcode3Txt.equalsIgnoreCase(expBarcode3Txt))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	
	}
	
	
	
	public boolean checkOtherDetailsTabDataAfterPasteItem() throws InterruptedException
	{
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(otherDetailsTab1));
		 otherDetailsTab1.click();
		 Thread.sleep(3000);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofIssueAccountTxt));
		String actCostofIssueTxt=itemCostofIssueAccountTxt.getAttribute("value");
		String expCostofIssueTxt="COGS POSTING ACC";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemStocksAccountTxt));
		String actStocksAccountTxt=itemStocksAccountTxt.getAttribute("value");
		String expStocksAccountTxt="BR COGS ACC INV";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSalesAccountTxt));
		String actSalesAccountTxt=itemSalesAccountTxt.getAttribute("value");
		String expSalesAccountTxt="Sales - Computers";	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWIPAccountTxt));
		String actWIPAccountTxt=Boolean.toString(itemWIPAccountTxt.getAttribute("value").isEmpty());
		String expWIPAccountTxt="true";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofShortageStockAccTxt));
		String actCostofShortageStockAccTxt=itemCostofShortageStockAccTxt.getAttribute("value");
		String expCostofShortageStockAccTxt="SHORTAGE COGS POSTING ACC";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofExcessStockAccTxt));
		String actCostofExcessStockAccTxt=itemCostofExcessStockAccTxt.getAttribute("value");
		String expCostofExcessStockAccTxt="EXCESS COGS POSTING ACC";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofSalesReturnAccTxt));
		String actCostofSalesReturnAccTxt=itemCostofSalesReturnAccTxt.getAttribute("value");
		String expCostofSalesReturnAccTxt="SR COGS POSTING ACC";	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemPurchaseVarianceAccTxt));
		String actPurchaseVarianceAccTxt=Boolean.toString(itemPurchaseVarianceAccTxt.getAttribute("value").isEmpty());
		String expPurchaseVarianceAccTxt="true";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemVATExpenseAccTxt));
		String actVATExpenseAccTxt=Boolean.toString(itemVATExpenseAccTxt.getAttribute("value").isEmpty());
		String expVATExpenseAccTxt="true";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemProductionLeadTimeTxt));
		String actProductionLeadTimeTxt=itemProductionLeadTimeTxt.getAttribute("value");
		String expProductionLeadTimeTxt="0";	
		
		
		System.err.println("Item Text in Other Details Tab");
		System.err.println("********************************************************************************************************************************************************************************");
		System.err.println("Cost of Issue Account			"		+		"	Actual		"	+	actCostofIssueTxt						+		"	Expected		"	+	expCostofIssueTxt);
		System.err.println("Stocks Account					"		+		"	Actual		"	+	actStocksAccountTxt						+		"	Expected		"	+	expStocksAccountTxt);
		System.err.println("Sales Account					"		+		"	Actual		"	+	actSalesAccountTxt						+		"	Expected		"	+	expSalesAccountTxt);
		System.err.println("WIP Account						"		+		"	Actual		"	+	actWIPAccountTxt						+		"	Expected		"	+	expWIPAccountTxt);
		System.err.println("Cost of Shortage Stock A/C		"		+		"	Actual		"	+	actCostofShortageStockAccTxt			+		"	Expected		"	+	expCostofShortageStockAccTxt);
		System.err.println("Cost of Excess Stock A/C		"		+		"	Actual		"	+	actCostofExcessStockAccTxt				+		"	Expected		"	+	expCostofExcessStockAccTxt);
		System.err.println("Cost of Sale Return A/C			"		+		"	Actual		"	+	actCostofSalesReturnAccTxt				+		"	Expected		"	+	expCostofSalesReturnAccTxt);
		System.err.println("Purchase Varience A/C			"		+		"	Actual		"	+	actPurchaseVarianceAccTxt				+		"	Expected		"	+	expPurchaseVarianceAccTxt);
		System.err.println("VAT Expense Account				"		+		"	Actual		"	+	actVATExpenseAccTxt						+		"	Expected		"	+	expVATExpenseAccTxt);
		System.err.println("Production Lead Time			"		+		"	Actual		"	+	actProductionLeadTimeTxt				+		"	Expected		"	+	expProductionLeadTimeTxt);
	
		if(actCostofIssueTxt.equalsIgnoreCase(expCostofIssueTxt) && actStocksAccountTxt.equalsIgnoreCase(expStocksAccountTxt) &&actSalesAccountTxt.equalsIgnoreCase(expSalesAccountTxt)
				&&actWIPAccountTxt.equalsIgnoreCase(expWIPAccountTxt) && actCostofShortageStockAccTxt.equalsIgnoreCase(expCostofShortageStockAccTxt) && actCostofExcessStockAccTxt.equalsIgnoreCase(expCostofExcessStockAccTxt)
				&& actCostofSalesReturnAccTxt.equalsIgnoreCase(expCostofSalesReturnAccTxt) && actPurchaseVarianceAccTxt.equalsIgnoreCase(expPurchaseVarianceAccTxt)
				&& actVATExpenseAccTxt.equalsIgnoreCase(expVATExpenseAccTxt) &&  actProductionLeadTimeTxt.equalsIgnoreCase(expProductionLeadTimeTxt))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	
	}
	
	
	public boolean checkReplenishmentTabDataAfterPasteItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(replenishmentTab1));
		 replenishmentTab1.click();
		 Thread.sleep(3000);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectDefaultReplenidshmentTxt));
		Select s=new Select(itemSelectDefaultReplenidshmentTxt);
		String actDefaultReplenidshmentTxt=s.getFirstSelectedOption().getText();
		String expDefaultReplenidshmentTxt="Purchase";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectManufacturePolicyTxt));
		Select s1=new Select(itemSelectManufacturePolicyTxt);
		String actManufacturePolicyTxt=s1.getFirstSelectedOption().getText();
		String expManufacturePolicyTxt="Made to stock";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemPackingBOMTxt));
		String actPackingBOMTxt=Boolean.toString(itemPackingBOMTxt.getAttribute("value").isEmpty());
		String expPackingBOMTxt="true";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBOMTxt));
		String actBOMTxt=Boolean.toString(itemBOMTxt.getAttribute("value").isEmpty());
		String expBOMtxt="true";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectFlushingTxt));
		Select s2=new Select(itemSelectFlushingTxt);
		String actFlushingTxt=s2.getFirstSelectedOption().getText();
		String expFlushingTxt="Manual";
		
		System.err.println("Item Text in Replenishment Tab");
		System.err.println("********************************************************************************************************************************************************************************");
		System.err.println("Default Replenishment	"		+		"	Actual		"	+	actDefaultReplenidshmentTxt						+		"	Expected		"	+	expDefaultReplenidshmentTxt);
		System.err.println("Manufacture Policy		"		+		"	Actual		"	+	actManufacturePolicyTxt							+		"	Expected		"	+	expManufacturePolicyTxt);
		System.err.println("Packing BOM				"		+		"	Actual		"	+	actPackingBOMTxt								+		"	Expected		"	+	expPackingBOMTxt);
		System.err.println("BOM						"		+		"	Actual		"	+	actBOMTxt										+		"	Expected		"	+	expBOMtxt);
		System.err.println("Flushing				"		+		"	Actual		"	+	actFlushingTxt									+		"	Expected		"	+	expFlushingTxt);
		
		System.out.println("*********Saving Item*************");
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		saveBtn.click();
		 
		 String expMessage="Code is unique";
		 String actMessage=checkValidationMessage(expMessage);
		 
		 System.out.println("actMessage  : "+actMessage);
		 System.out.println("expMessage  : "+expMessage);
	
		if(actDefaultReplenidshmentTxt.equalsIgnoreCase(expDefaultReplenidshmentTxt) && actManufacturePolicyTxt.equalsIgnoreCase(expManufacturePolicyTxt)
				&& actPackingBOMTxt.equalsIgnoreCase(expPackingBOMTxt) && actBOMTxt.equalsIgnoreCase(expBOMtxt) && actFlushingTxt.equalsIgnoreCase(expFlushingTxt)
				&& actMessage.equalsIgnoreCase(expMessage))
		{
			
			
			return true;
		}
		else
		{
			return false;
		}
		
		
		
	}
	
	public static boolean CheckSavingItemAfterPaste() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalTab1));
		 generalTab1.click();
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Paste Item");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("Paste Item");
		 codeTxt.sendKeys(Keys.TAB);
		 
		 //Thread.sleep(2000);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Saved Successfully";
		 String actMessage=checkValidationMessage(expMessage);
		 
		 System.out.println("actMessage  : "+actMessage);
		 System.out.println("expMessage  : "+expMessage);
		 
		 Thread.sleep(8000);
		 
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
		
	  @FindBy(xpath="//*[@id='btnClone']")
	    public static WebElement itemCloneHome;
	
	public boolean checkItemCloneOptionFromHome() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{

		 getDriver().navigate().refresh();
			Thread.sleep(2000);
			
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	      	homeMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	      	mastersMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
	    	homeMasterItem_ItemMenu.click();
	    	
	    	Thread.sleep(2000);
	    	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_ItemLabel));
      	
		int accNameListCount = accNameList.size();

		for(int i=0;i<accNameListCount;i++)
		{
			String data = accNameList.get(i).getText();
			Thread.sleep(2000);
			if(data.equals("BR COGS ITEM"))
			{
				Thread.sleep(2000);
				if (masterGridBodyChkbox.get(i).isSelected()==false) 
				{
					masterGridBodyChkbox.get(i).click();
					break;
				}
			}
		}
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCloneHome));
		 itemCloneHome.click();
		
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalTab1));
		 Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Clone Item");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("Clone Item");
		 codeTxt.sendKeys(Keys.TAB);
		 
		 Thread.sleep(2000);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage1="Saved Successfully";
		 String actMessage1=checkValidationMessage(expMessage1);
		 
		 System.out.println("actMessage1  : "+actMessage1);
		 System.out.println("expMessage1  : "+expMessage1);
		 
		 
		 Thread.sleep(2000);
		 
			/*
			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * masterCloseBtn)); masterCloseBtn.click();
			 */
			
		 if(actMessage1.equalsIgnoreCase(expMessage1))
		 {
			 return true;
		 }
		 
		 else
		 {
			 return false;
		 }
	
	}
	    
	 public boolean checkCloneIteminDetailsTabAfterSave() throws InterruptedException
	 {
		 Thread.sleep(6000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	      	homeMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	      	mastersMenu.click();
	      	
	      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
	      	homeMasterItemMenu.click();
	      	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
	    	homeMasterItem_ItemMenu.click();
	    	
	    	Thread.sleep(2000);
		 	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_ItemLabel));
	      	
			int accNameListCount = accNameList.size();

			for(int i=0;i<accNameListCount;i++)
			{
				String data = accNameList.get(i).getText();
				Thread.sleep(2000);
				if(data.equals("Clone Item"))
				{
					Thread.sleep(2000);
					getAction().doubleClick(accChkBoxList.get(i)).build().perform();
					Thread.sleep(2000);
					break;
				}
			}
			
			Thread.sleep(8000);
			
			
		
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalTab1));
			 generalTab1.click();
			 Thread.sleep(2000);;
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			String actItemName=nameTxt.getAttribute("value");
			String expItemName="Clone Item";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			String actItemCode=codeTxt.getAttribute("value");
			String expItemCode="Clone Item";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(productTypeSelect));
			Select s1=new Select(productTypeSelect);
			String actProductType=s1.getFirstSelectedOption().getText();
			String expProductType="Raw material";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reorderLevel));
			String actReorderLevelTxt=reorderLevel.getAttribute("value");
			String expReorderLevelTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binCapacity));
			String actBinCapacity=binCapacity.getAttribute("value");
			String expBinCapacity="1.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valuationMethodSelect));
			Select s2=new Select(valuationMethodSelect);
			String actValuationMethod=s2.getFirstSelectedOption().getText();
			String expValuationMethod="Batch Rate";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemMakeSelect));
			Select s3=new Select(itemMakeSelect);
			String actProductMake=s3.getFirstSelectedOption().getText();
			String expProductMake="Item";
			
			
			
			System.err.println("Item Text in Details Tab");
			System.err.println("Name	"			+		"	Actual		"	+	actItemName			+		"	Expected		"	+	expItemName);
			System.err.println("Code	"			+		"	Actual		"	+	actItemCode			+		"	Expected		"	+	expItemCode);
			System.err.println("Item Type	"		+		"	Actual		"	+	actProductType		+		"	Expected		"	+	expProductType);
			System.err.println("Reorder Level"		+		"	Actual		"	+	actReorderLevelTxt	+		"	Expected		"	+	expReorderLevelTxt);
			System.err.println("Bin Capacity"		+		"	Actual		"	+	actBinCapacity		+		"	Expected		"	+	expBinCapacity);
			System.err.println("Valuation Method"	+		"	Actual		"	+	actValuationMethod	+		"	Expected		"	+	expValuationMethod);
			System.err.println("Product Make"		+		"	Actual		"	+	actProductMake		+		"	Expected		"	+	expProductMake);
			
			if(actItemName.equalsIgnoreCase(expItemName) && actItemCode.equalsIgnoreCase(expItemCode) && actProductType.equalsIgnoreCase(expProductType) &&actReorderLevelTxt.equalsIgnoreCase(expReorderLevelTxt)
					&& actBinCapacity.equalsIgnoreCase(expBinCapacity) && actValuationMethod.equalsIgnoreCase(expValuationMethod) && actProductMake.equalsIgnoreCase(expProductMake))
			
			{
			
			return true;
			}
			else
			{
				return false;
			}
	 
		 
	 }	
	 
	 @FindBy(xpath="//*[@id='btnProformaExportForAdvanceFormPopUp']")
		public static WebElement AdvMasterExportBtn;
		
		@FindBy(xpath="//*[@id='AdvMasterExportModalLabel']")
		public static WebElement ExportLabel;
		
		
		@FindBy(xpath="//*[@id='ulMasterExportTab']//li[contains(text(),'Units')]")
		public static WebElement ExportUnitsTab;
		
		@FindBy(xpath="//*[@id='ulMasterExportTab']//li[contains(text(),'Settings')]")
		public static WebElement ExportSettingsTab;
		
		@FindBy(xpath="//*[@id='ulMasterExportTab']//li[contains(text(),'Classification')]")
		public static WebElement ExportClassificationTab;
		
		@FindBy(xpath="//*[@id='ulMasterExportTab']//li[contains(text(),'Other Details')]")
		public static WebElement ExportOtherDetailsTab;
		
		@FindBy(xpath="//*[@id='ulMasterExportTab']//li[contains(text(),'Other Details')]")
		public static WebElement ExportReplenishmentTab;
		
		@FindBy(xpath="//*[@id='ulMasterExportTab']//li[contains(text(),'Other Details')]")
		public static WebElement ExportReorderTab;
		
		@FindBy(xpath="//input[@id='btnMasterAdvanceExportSelectAll']")
		public static WebElement ExportSelectAllChkBoxSelected;
		
		@FindBy(xpath="//input[@id='btnMasterAdvanceExportSelectAll']//..//span")
		public static WebElement ExportSelectAllChkBox;
		
		@FindBy(xpath="//*[@id='btnMasterAdvanceExport']")
		public static WebElement ExportBtn;
	 
	 
	 public boolean checkAdvanceMasterExportItem() throws InterruptedException, AWTException, IOException
		{
		 	Thread.sleep(8000);
		  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
		   	toggleRibbon.click();
		   	Thread.sleep(2000);
		   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
		   	advanceMasterImportExportBtn.click();
		   	
			Thread.sleep(2000);
			
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterExportBtn));
		 	AdvMasterExportBtn.click();
		 	Thread.sleep(2000);
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportLabel));
		 	System.out.println("Advance Master import/Export Label  "+ExportLabel.getText());
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportUnitsTab));
		 	ExportUnitsTab.click();
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportSelectAllChkBox));
		 	ExportSelectAllChkBox.click();
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportSettingsTab));
		 	ExportSettingsTab.click();
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportSelectAllChkBox));
		 	ExportSelectAllChkBox.click();
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportClassificationTab));
		 	ExportClassificationTab.click();
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportSelectAllChkBox));
		 	ExportSelectAllChkBox.click();
		 	Thread.sleep(2000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportOtherDetailsTab));
		 	ExportOtherDetailsTab.click();
		 	Thread.sleep(3000);
		 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportSelectAllChkBox));
		 	ExportSelectAllChkBox.click();
		 	Thread.sleep(3000);
		 	
			/*
			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * ExportReplenishmentTab)); ExportReplenishmentTab.click(); Thread.sleep(2000);
			 * 
			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * ExportSelectAllChkBox)); ExportSelectAllChkBox.click(); Thread.sleep(2000);
			 */
		 	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportBtn));
			ExportBtn.click();
			Thread.sleep(10000);
			
			File Efile1 = new File(getBaseDir() + "\\autoIt\\ExportFiles\\ItemExport.xlsx");

			if (Efile1.exists()) {
				Efile1.delete();
			}
			
			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);

			Thread.sleep(8000);

			robot.keyPress(KeyEvent.VK_TAB);
			robot.keyRelease(KeyEvent.VK_TAB);
			robot.keyPress(KeyEvent.VK_TAB);
			robot.keyRelease(KeyEvent.VK_TAB);
			
			Thread.sleep(4000);
			
			  robot.keyPress(KeyEvent.VK_ENTER); robot.keyRelease(KeyEvent.VK_ENTER);
			 
			Thread.sleep(8000);
			// robot.keyPress(KeyEvent.VK_ENTER); robot.keyRelease(KeyEvent.VK_ENTER);
			// Thread.sleep(2000);
			 
			robot.keyPress(KeyEvent.VK_DOWN);
			robot.keyPress(KeyEvent.VK_DOWN);
			robot.keyPress(KeyEvent.VK_DOWN);
			robot.keyPress(KeyEvent.VK_DOWN);
			robot.keyPress(KeyEvent.VK_DOWN);
			robot.keyPress(KeyEvent.VK_DOWN);
			robot.keyRelease(KeyEvent.VK_DOWN);
			robot.keyRelease(KeyEvent.VK_DOWN);
			robot.keyRelease(KeyEvent.VK_DOWN);
			robot.keyRelease(KeyEvent.VK_DOWN);
			robot.keyRelease(KeyEvent.VK_DOWN);
			robot.keyRelease(KeyEvent.VK_DOWN);

			Thread.sleep(5000);
			
			robot.keyPress(KeyEvent.VK_ENTER); robot.keyRelease(KeyEvent.VK_ENTER);
			Thread.sleep(4000);
			
			robot.keyPress(KeyEvent.VK_ENTER); robot.keyRelease(KeyEvent.VK_ENTER);
			 Thread.sleep(4000);
			 
			 Thread.sleep(4000);
				robot.keyPress(KeyEvent.VK_F12);
				robot.keyRelease(KeyEvent.VK_F12);
				 Thread.sleep(4000);
				robot.keyPress(KeyEvent.VK_LEFT);
				robot.keyRelease(KeyEvent.VK_LEFT);
				 Thread.sleep(2000);
				 
				 robot.keyPress(KeyEvent.VK_ENTER); robot.keyRelease(KeyEvent.VK_ENTER);
				 Thread.sleep(2000);
				 robot.keyPress(KeyEvent.VK_F12);
					robot.keyRelease(KeyEvent.VK_F12);
					 Thread.sleep(2000);
				 
			Runtime.getRuntime().exec(getBaseDir() + "\\autoIt\\scripts\\ItemExport.exe");

			Thread.sleep(8000);

			robot.keyPress(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);
			Thread.sleep(4000);
			
			
			robot.keyPress(KeyEvent.VK_ALT);
			robot.keyPress(KeyEvent.VK_F4);
			
			robot.keyRelease(KeyEvent.VK_ALT);
			robot.keyRelease(KeyEvent.VK_F4);
			Thread.sleep(3000);
			
			robot.keyPress(KeyEvent.VK_ALT);
			robot.keyPress(KeyEvent.VK_F4);
			
			robot.keyRelease(KeyEvent.VK_ALT);
			robot.keyRelease(KeyEvent.VK_F4);
			Thread.sleep(3000);
			
			ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());

			int actOpenWindowsCount = getDriver().getWindowHandles().size();
			int expOpenWindowsCount = 2;

			System.out.println(
					"Number of Windows  : " + actOpenWindowsCount + "  Value Expected  " + expOpenWindowsCount);

			getDriver().switchTo().window(newTabs.get(1)).close();
			Thread.sleep(1000);
			
			getDriver().switchTo().window(newTabs.get(0));
			Thread.sleep(1000);
			
			return true;
		 	
			
		}
		
		
		public boolean checkValidatingExcelFile() throws IOException, EncryptedDocumentException, InvalidFormatException
		{
			
			 boolean result=false;
	        try
	        {
	            FileInputStream excellFile1 = new FileInputStream(new File("E:\\FocusXSanity\\FocusAI\\autoIt\\ExportFiles\\ItemExport.xlsx"));
	            FileInputStream excellFile2 = new FileInputStream(new File("E:\\FocusXSanity\\FocusAI\\autoIt\\ImportFiles\\ItemExport.xlsx"));

	            XSSFWorkbook workbook1 = new XSSFWorkbook(excellFile1);
	            XSSFWorkbook workbook2 = new XSSFWorkbook(excellFile2);

	            XSSFSheet sheet1 = workbook1.getSheetAt(0);
	            XSSFSheet sheet2 = workbook2.getSheetAt(0);
	               
	            
	             result=compareTwoSheets(sheet1, sheet2);
	            excellFile1.close();
	            excellFile2.close();
	            
	        }
	        
	        catch (Exception e)
	        {
	            e.printStackTrace();
	        }
	           
	        if(result)
	        {
	            System.out.println("\nTwo Excelsheets are Equal");
	            return true;
	        }
	        else
	        {
	            System.out.println("\nTwo Excelsheets are Not Equal");
	            return false;
	        }     
	           
	           
	        }
	       
			
			
	   
		
		  public static boolean compareTwoSheets(XSSFSheet sheet1, XSSFSheet sheet2)
		    {
		        int firstRow1 = sheet1.getFirstRowNum();
		       XSSFRow r=sheet1.getRow(0);
		       int col=r.getLastCellNum();
		        int lastRow1 = sheet2.getLastRowNum();
		        boolean equalSheets = true;
		        System.out.println("Number of Rows:" + lastRow1);
		        System.out.println("Number of Columns:" + col);
		        ArrayList<String> ColumnsName = new ArrayList<String>();
		        for(int i=0;i<col;i++)
		        {
		        
		        	System.err.println(sheet2.getRow(3).getCell(i));
		        	
		        }
		       
		        
		        for(int i=firstRow1; i <= lastRow1; i++)
		        {
		            /*System.out.print("___________________________");
		            System.out.println("\nComparing Row "+i);
		            System.out.println("___________________________");*/
		            XSSFRow row1 = sheet1.getRow(i);
		            XSSFRow row2 = sheet2.getRow(i);
		            if(!compareTwoRows(row1, row2))
		            {
		                equalSheets = false;
		               // System.out.println(" Row "+i+" | Not Equal");
		            }
		            else
		            {
		               // System.out.println(" Row "+i+" | Equal");
		            }
		        }
		        return equalSheets;
		        
		    }
		    
		  public static boolean compareTwoRows(XSSFRow row1, XSSFRow row2)
		    {
		        if((row1 == null) && (row2 == null))
		        {
		            return true;
		        }
		        else if((row1 == null) || (row2 == null))
		        {
		            return false;
		        }
		        int firstCell1 = row1.getFirstCellNum();
		        int lastCell1 = row1.getLastCellNum();
		        boolean equalRows = true;

		        for(int i=firstCell1; i <= lastCell1; i++)
		        {
		            XSSFCell cell1 = row1.getCell(i);
		          System.err.println( cell1);
		          
		            XSSFCell cell2 = row2.getCell(i);
		            System.err.println( cell1);
		            if(!compareTwoCells(cell1, cell2))
		            {
		                equalRows = false;
		               
		               
		               // System.err.println("Cell "+i+" | Not Equal");
		            }
		            else
		            {
		            	
		            
		             //   System.out.println("Cell "+i+" | Equal");
		            }
		            System.out.println();
		        }
		        return equalRows;
		    }

		    public static boolean compareTwoCells(XSSFCell cell1, XSSFCell cell2)
		    {
		    	
		        if((cell1 == null) && (cell2 == null))
		        {
		            return true;
		        }
		        else if((cell1 == null) || (cell2 == null))
		        {
		            return false;
		        }
		        
		        boolean equalCells = false;
		        int type1 = cell1.getCellType();
		        int type2 = cell2.getCellType();
		        if (type1 == type2)
		        {
		            if (cell1.getCellStyle().equals(cell2.getCellStyle()))
		            {
		                switch (cell1.getCellType())
		                {
		                    case HSSFCell.CELL_TYPE_FORMULA:
		                        if (cell1.getCellFormula().equals(cell2.getCellFormula()))
		                        {
		                            equalCells = true;
		                        }
		                        break;
		                    case HSSFCell.CELL_TYPE_NUMERIC:
		                        if (cell1.getNumericCellValue() == cell2.getNumericCellValue())
		                        {
		                            equalCells = true;
		                        }
		                        break;
		                    case HSSFCell.CELL_TYPE_STRING:
		                        if (cell1.getStringCellValue().equals(cell2.getStringCellValue()))
		                        {
		                            equalCells = true;
		                        }
		                        break;
		                    case HSSFCell.CELL_TYPE_BLANK:
		                        if (cell2.getCellType() == HSSFCell.CELL_TYPE_BLANK)
		                        {
		                            equalCells = true;
		                        }
		                        break;
		                    case HSSFCell.CELL_TYPE_BOOLEAN:
		                        if (cell1.getBooleanCellValue() == cell2.getBooleanCellValue())
		                        {
		                            equalCells = true;
		                        }
		                        break;
		                    case HSSFCell.CELL_TYPE_ERROR:
		                        if (cell1.getErrorCellValue() == cell2.getErrorCellValue())
		                        {
		                            equalCells = true;
		                        }
		                        break;
		                    default:
		                        if (cell1.getStringCellValue().equals(cell2.getStringCellValue()))
		                        {
		                            equalCells = true;
		                        }
		                        break;
		                }
		            }
		            else
		            {
		                return false;
		            }
		        }
		        else
		        {
		            return false;
		        }
		        return equalCells;
		    }


		    
		    @FindBy(xpath="//*[@id='ddlBulkImport']")
		    public static WebElement itemBulkSelect;
		    
		   
		    
		    @FindBy(xpath="//i[@class='icon-close']")
		    public static WebElement AdvMasterCloseBtn;
		    
		    @FindBy(xpath="//h1")
		    public static WebElement serverHeaderMsg;;
		    
		    @FindBy(xpath="(//*[@id='innerDivAdvMasterExportImport']/div[3]//span)[1]")
		    public static WebElement itemBulkBtn;
		    
		   
		    
		    @FindBy(xpath="//*[@id='btnImportMastersBulkData']")
		    public static WebElement itemBulkImport;
		    
		    @FindBy(xpath="//*[@id='btnExportMastersBulkData']")
		    public static WebElement itemBulkExport;;
		    
		  
		    
		public boolean checkImportDatawithEnableBulkOption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
			 Thread.sleep(2000);
			 
			 String expMessage1 = null,actMessage1=null;
		
			   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBulkSelect));
			   	itemBulkSelect.click();;
		    	 Select s=new Select(itemBulkSelect);
		    	 s.selectByVisibleText("Yes");
		    	 Thread.sleep(2000);
		    	 
				
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
			FileUploadBtn.click();
			Thread.sleep(2000);
			
			 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportFromExcel.exe");
			 Thread.sleep(5000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBulkBtn));
			 itemBulkBtn.click();
			 Thread.sleep(2000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBulkImport));
			 itemBulkImport.click();
			 Thread.sleep(2000);
			 
			 try {
				 
				 expMessage1="1 Records Imported Successfully";
				 actMessage1=checkValidationMessage(expMessage1);
				
				 
			 }
			 catch (Exception e) {
				
				 System.err.println("Server Error");
				 
				 
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
				 ServerErrorPopupCloseBtn.click();
				 Thread.sleep(2000);
				 
			}
			 Thread.sleep(2000);
			 
			/* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterCloseBtn));
			 AdvMasterCloseBtn.click();
			 Thread.sleep(2000);*/
			 
			 
			 if(actMessage1.equalsIgnoreCase(expMessage1))
			 {
				 return true;
			 }
			 else
			 {
				 return false;
			 }
			
			
		}
		
		
		
		 
		 ///Item Set
		 
		 
		 @FindBy(xpath="//*[@id='btnNewSets']")
		 public static WebElement item_NewSet;
		 
		 @FindBy(xpath="//*[@id='optProductSets']")
		 public static WebElement item_NewSetName;
		 
		 @FindBy(xpath="//*[@id='txtCode']")
		 public static WebElement item_NewSetCode;
		 
		 @FindBy(xpath="//*[@id='txtAlias']")
		 public static WebElement item_NewSetAlias;
		 
		 @FindBy(xpath="//*[@id='Grid_ProductSet_body']/tr[1]//td")
		 public static List<WebElement> item_NewSet1stRowList;
		 
		 @FindBy(xpath="//*[@id='Grid_ProductSet_body']/tr[1]/td[2]")
		 public static WebElement item_NewSet1stRow1stCol;
		 
		 @FindBy(xpath="//*[@id='Grid_ProductSet_body']/tr[1]/td[3]")
		 public static WebElement item_NewSet1stRow2ndCol;
		 
		 @FindBy(xpath="//*[@id='Grid_ProductSet_body']/tr[1]/td[5]")
		 public static WebElement item_NewSet1stRow4thCol;
		 
		 @FindBy(xpath="//*[@id='Grid_ProductSet_body']/tr[1]/td[6]")
		 public static WebElement item_NewSet1stRow5thCol;
		 
		 @FindBy(xpath="//*[@id='optProduct']")
		 public static WebElement item_NewSetItemTxt;
		 
		 @FindBy(xpath="//*[@id='Grid_ProductSet_control_heading_ctrl_4']")
		 public static WebElement item_NewSetQuantityTxt;
		 
		 @FindBy(xpath="//*[@id='divMessageModalBody']//label")
		 public static WebElement item_NewSetSaveMsgTxt;
		 
		// @FindBy(xpath="//*[@id='divMessageModalFooter']//input")
		 @FindBy(xpath="//*[@id='divMessageModalBody']//input")
		 public static WebElement item_NewSetOKBtn;
		 
		 
		 @FindBy(xpath="//*[@id='btnMasterSaveClick']")
		 public static WebElement item_NewSetSaveBtn;
		 
		 @FindBy(xpath="(//*[@id='btnDelete'])[2]")
		 public static WebElement item_NewSetDeleteBtn;
		 
		 
		 
		 
		 
		 public boolean checkSavingItemSetinItemMaster() throws InterruptedException
		 {
			 Thread.sleep(2000);
				
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		      	homeMenu.click();
		      	Thread.sleep(2000);
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		      	mastersMenu.click();
		      	
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		      	homeMasterItemMenu.click();
		      	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
		    	homeMasterItem_ItemMenu.click();
		    	
		    	Thread.sleep(5000);
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSet));
		    	item_NewSet.click();
		    	Thread.sleep(40000);
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetName));
		    	item_NewSetName.click();
		    	item_NewSetName.sendKeys("Set1");
		    	item_NewSetName.sendKeys(Keys.TAB);
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetCode));
		    	item_NewSetCode.click();
		    	item_NewSetCode.sendKeys("Set1");
		    	item_NewSetCode.sendKeys(Keys.TAB);
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetAlias));
		    	item_NewSetAlias.click();
		    	item_NewSetAlias.sendKeys("Set1");
		    	item_NewSetAlias.sendKeys(Keys.TAB);
		    	
		    	Thread.sleep(4000);
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSet1stRow1stCol));
		    	item_NewSet1stRow1stCol.click();
		    	item_NewSetItemTxt.sendKeys("Item1");
		    	Thread.sleep(4000);
		    	item_NewSetItemTxt.sendKeys(Keys.TAB);
		    	Thread.sleep(6000);
		    	
		    	String actItemCode=item_NewSet1stRow2ndCol.getText();
		    	String expItemCode="Item1";
		    	System.out.println("Actual Item Code	"	+	actItemCode);
		    	System.out.println("Expeced Item Code	"	+	expItemCode);
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSet1stRow4thCol));
		    	item_NewSet1stRow4thCol.click();
		    	
		    	item_NewSetQuantityTxt.sendKeys("3");
		    	item_NewSetQuantityTxt.sendKeys(Keys.TAB);
		    	
		    	String actItemRate=item_NewSet1stRow5thCol.getText();
		    	String expItemRate="11.26";
		    	System.out.println("Actual Item Rate	"	+	actItemRate);
		    	System.out.println("Expected Item Rate	"	+	expItemRate);
		    	
		    	Thread.sleep(2000);

		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetSaveBtn));
		    	item_NewSetSaveBtn.click();
		    	Thread.sleep(2000);
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetSaveMsgTxt));
		    	
		    	String actSaveMsg=item_NewSetSaveMsgTxt.getText();
		    	String expSaveMsg="Details saved successfully.";
		    	System.out.println("Actual Save Msg			"	+	actSaveMsg);
		    	System.out.println("Expected Save Msg		"	+	expSaveMsg);
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetOKBtn));
		    	item_NewSetOKBtn.click();
		    	
		    	if(actItemCode.equalsIgnoreCase(expItemCode) && actItemRate.equalsIgnoreCase(expItemRate) && actSaveMsg.equalsIgnoreCase(expSaveMsg))
		    	{
		    		return true;
		    	}
		    	
		    	else
		    	{
		    		return false;
		    	}
		    	
		    	
		    	
		 }
		 
		 
		 public boolean checkVerifyingtheItemSetinItemMaster() throws InterruptedException
		 {
			 Thread.sleep(10000);	
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetName));
		    	item_NewSetName.click();
		    	item_NewSetName.sendKeys("Set1");
		    	 Thread.sleep(4000);
		    	item_NewSetName.sendKeys(Keys.TAB);
		    	 Thread.sleep(10000);
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetCode));
		    	String actSetCodeTxt=item_NewSetCode.getAttribute("value");
		    	String expSetCodeTxt="Set1";
		    	
		    	System.out.println("Actual Set Code			"		+	actSetCodeTxt);
		    	System.out.println("Expected  Set Code		"		+	expSetCodeTxt);
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetAlias));
		    	String actSetAliasTxt=item_NewSetAlias.getAttribute("value");
		    	String expSetAliasTxt="Set1";
		    	System.out.println("Actual Set Alias		"	+	actSetAliasTxt);
		    	System.out.println("Expected Set Alias		"	+	expSetAliasTxt);
		    	
		    	int newSet1stRowListCount = item_NewSet1stRowList.size();
		    	ArrayList<String> newSet1stRowListArray = new ArrayList<String>();
		    	for(int i=0;i<newSet1stRowListCount;i++)
		    	{
		    		String data = item_NewSet1stRowList.get(i).getText();
		    		if(!data.isEmpty())
		    		{
		    			newSet1stRowListArray.add(data);
		    		}
		    		
		    	}
		    	String actNewSetList = newSet1stRowListArray.toString();
		    	String expNewSetList = "[1, Item1, Item1, 3.00, 11.26]";
		    	
		    	System.out.println("Actual Item List	"	+	actNewSetList);
		    	System.out.println("Expected Item List	"	+	expNewSetList);
		    	
		    	if(actSetCodeTxt.equalsIgnoreCase(expSetCodeTxt) && actSetAliasTxt.equalsIgnoreCase(expSetAliasTxt) &&actNewSetList.equalsIgnoreCase(expNewSetList) )
		    	{
		    		return true;
		    	}
		    	else
		    	{
		    		return false;
		    	}
		 }
		 
		 public boolean checkDeletingItemSetinItemMaster() throws InterruptedException
		 {
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetDeleteBtn));
			 item_NewSetDeleteBtn.click();
			 Thread.sleep(8000);
			 
			
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetSaveMsgTxt));
		    	
		    	String actDeleteConfirmMsg=item_NewSetSaveMsgTxt.getText();
		    	String expDeleteConfirmMsg="Are you sure to Delete?";
		    	System.out.println("Actual Delete Confirm  Msg			"	+	actDeleteConfirmMsg);
		    	System.out.println("Expected Delete Confirm  Msg		"	+	expDeleteConfirmMsg);
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetOKBtn));
		    	item_NewSetOKBtn.click();
		    	
		    	
			    	Thread.sleep(8000);
			    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetSaveMsgTxt));
			    	
			    	String actDeleteConfirmMsg1=item_NewSetSaveMsgTxt.getText();
			    	String expDeleteConfirmMsg1="Deleted successfully";
			    	System.out.println("Actual Delete Confirm  Msg			"	+	actDeleteConfirmMsg1);
			    	System.out.println("Expected Delete Confirm  Msg		"	+	expDeleteConfirmMsg1);
			    	
			    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item_NewSetOKBtn));
			    	item_NewSetOKBtn.click();
			    	
			    	if(actDeleteConfirmMsg.equalsIgnoreCase(expDeleteConfirmMsg) && actDeleteConfirmMsg1.equalsIgnoreCase(expDeleteConfirmMsg1))
			    	{
			    		return true;
			    	}
			    	else
			    	{
			    		return false;
			    	}
			 
		 }
		 
		
		//Bulk Import
		 
		 
		 public  void setCellData(String filePath,int sheetNo, int rowNum, int CellNum,String data,int newRowNo) throws IOException, EncryptedDocumentException, InvalidFormatException
		    {
			  if(newRowNo==0)	
			  {
			  this.filePath=filePath;
			  	 FileInputStream fip=new FileInputStream(filePath);
				   workbook =WorkbookFactory.create(fip);
				   
				    sheet=workbook.getSheetAt(sheetNo);
				   
				    row=sheet.getRow(rowNum);
			   
			    row.createCell(CellNum).setCellValue(data);
			    FileOutputStream fop=new FileOutputStream(filePath);
				
				   workbook.write(fop);
				   
				   fip.close();
				   
				   fop.close();
			  }
			  else
			  {

				  	this.filePath=filePath;
				  	 FileInputStream fip=new FileInputStream(filePath);
					   workbook =WorkbookFactory.create(fip);
					   
					    sheet=workbook.getSheetAt(sheetNo);
					   
					    row=sheet.createRow(rowNum);
				   
				    row.createCell(CellNum).setCellValue(data);
				    FileOutputStream fop=new FileOutputStream(filePath);
					
					   workbook.write(fop);
					   
					   fip.close();
					   
					   fop.close();
				    
			     
			    
			  }
		     
		    }
		  
		 
		 @FindBy(xpath="(//span[contains(text(),'Bulk')])[3]")
		 public static WebElement BulkBtn;
		 
		 @FindBy(xpath="//*[@id='btnImportMastersBulkData']")
		 public static WebElement BulkImportBtn;
		 
		 @FindBy(xpath="//*[@id='btnExportMastersBulkData']")
		 public static WebElement BulkExportBtn;
		 
		 
		 
		 @FindBy(xpath="//*[@id='ddlMasterAdvanceExportType']")
		 public static WebElement ExportTypeSelect;
		 
		 @FindBy(xpath="//*[@id='ddlBulkImport']")
		 public static WebElement BulkImportOptionSelect;
		 
		 
		 public boolean checkBulkExportAndImportinItemMaster() throws InterruptedException, AWTException, IOException
		 {
			 Thread.sleep(2000);
			 getDriver().navigate().refresh();
				Thread.sleep(4000);
				
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		      	homeMenu.click();
		      	Thread.sleep(2000);
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		      	mastersMenu.click();
		      	Thread.sleep(2000);
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		      	homeMasterItemMenu.click();
		      	Thread.sleep(2000);
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
		    	homeMasterItem_ItemMenu.click();
		    	
			 Thread.sleep(4000);
			  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
			   	toggleRibbon.click();
			   	
			   	Thread.sleep(2000);
			   	
			   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
			   	advanceMasterImportExportBtn.click();
			 	Thread.sleep(4000);
			 	
			 			 	
			 	click(BulkBtn);
			 	Thread.sleep(2000);
			 	ClickUsingJs(BulkExportBtn);
			 	Thread.sleep(2000);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportUnitsTab));
			 	ExportUnitsTab.click();
			 	Thread.sleep(2000);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportSelectAllChkBox));
			 	ExportSelectAllChkBox.click();
			 	Thread.sleep(2000);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportSettingsTab));
			 	ExportSettingsTab.click();
			 	Thread.sleep(2000);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportSelectAllChkBox));
			 	ExportSelectAllChkBox.click();
			 	Thread.sleep(2000);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportClassificationTab));
			 	ExportClassificationTab.click();
			 	Thread.sleep(2000);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportSelectAllChkBox));
			 	ExportSelectAllChkBox.click();
			 	Thread.sleep(2000);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportOtherDetailsTab));
			 	ExportOtherDetailsTab.click();
			 	Thread.sleep(3000);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportSelectAllChkBox));
			 	ExportSelectAllChkBox.click();
			 	Thread.sleep(3000);
			 	
				
			 	ClickUsingJs(ExportTypeSelect);
			 	Thread.sleep(2000);
			 	
			 	Select s=new Select(ExportTypeSelect);
			 	s.selectByVisibleText("Only Proforma");
			 	Thread.sleep(2000);
			 	
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExportBtn));
				ExportBtn.click();
				Thread.sleep(10000);
				
				
				Select s1=new Select(BulkImportOptionSelect);
			 	s1.selectByVisibleText("Yes");
			 	Thread.sleep(2000);
				
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
				FileUploadBtn.click();
				Thread.sleep(6000);
				
				 //Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportFromExcel.exe");
				 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportFromExcel.exe");
				 Thread.sleep(8000);
				 
				 click(BulkBtn);
				 Thread.sleep(2000);
				 ClickUsingJs(BulkImportBtn);
				 Thread.sleep(10000);
			 	
				 String expMessage1=null;
				 String actMessage1=null;
				 
				 try {
					 
					 
					 expMessage1="Records Imported Successfully";
					 actMessage1=checkValidationMessage(expMessage1);
					 
				 }
				 catch (Exception e) {
					 System.err.println("Server Error");
					 
					 
					 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
					 ServerErrorPopupCloseBtn.click();
					 Thread.sleep(2000);
					 
					 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterCloseBtn));
					 AdvMasterCloseBtn.click();
					 Thread.sleep(2000);
					 
					
				}
		
				 if(actMessage1.equalsIgnoreCase(expMessage1))
				 {
					 return true;
				 }
				 else
				 {
					 return false;
				 }
			
		 }
		 
		 
		 
		 public boolean checkItemsAfterBulkImportinItemMaster() throws InterruptedException
		 {	
			 Thread.sleep(2000);
			 int accNameListCount = accNameList.size();
			 
			 for(int i=0;i<accNameListCount;i++)		
					
				{
					String data = accNameList.get(i).getText();
					Thread.sleep(2000);
					if(data.equals("Steel"))
					{
						Thread.sleep(2000);
						getAction().doubleClick(accChkBoxList.get(i)).build().perform();
						Thread.sleep(2000);
						break;
					}
				}
			 Thread.sleep(2000);

			 generalTab1.click();
			 Thread.sleep(2000);;
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			String actItemName=nameTxt.getAttribute("value");
			String expItemName="Steel";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			String actItemCode=codeTxt.getAttribute("value");
			String expItemCode="Steel";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(productTypeSelect));
			Select s1=new Select(productTypeSelect);
			String actProductType=s1.getFirstSelectedOption().getText();
			String expProductType="Raw material";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reorderLevel));
			String actReorderLevelTxt=reorderLevel.getAttribute("value");
			String expReorderLevelTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binCapacity));
			String actBinCapacity=binCapacity.getAttribute("value");
			String expBinCapacity="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valuationMethodSelect));
			Select s2=new Select(valuationMethodSelect);
			String actValuationMethod=s2.getFirstSelectedOption().getText();
			String expValuationMethod="Weighted Avg";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemMakeSelect));
			Select s3=new Select(itemMakeSelect);
			String actProductMake=s3.getFirstSelectedOption().getText();
			String expProductMake="Item";
			
			
			
			System.err.println("Item Text in Details Tab");
			System.err.println("Name	"			+		"	Actual		"	+	actItemName			+		"	Expected		"	+	expItemName);
			System.err.println("Code	"			+		"	Actual		"	+	actItemCode			+		"	Expected		"	+	expItemCode);
			System.err.println("Item Type	"		+		"	Actual		"	+	actProductType		+		"	Expected		"	+	expProductType);
			System.err.println("Reorder Level"		+		"	Actual		"	+	actReorderLevelTxt	+		"	Expected		"	+	expReorderLevelTxt);
			System.err.println("Bin Capacity"		+		"	Actual		"	+	actBinCapacity		+		"	Expected		"	+	expBinCapacity);
			System.err.println("Valuation Method"	+		"	Actual		"	+	actValuationMethod	+		"	Expected		"	+	expValuationMethod);
			System.err.println("Product Make"		+		"	Actual		"	+	actProductMake		+		"	Expected		"	+	expProductMake);
			
			
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitsTab1));
			 unitsTab1.click();
			 Thread.sleep(3000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultBaseUnitTxt));
			String actDefaultBase=itemDefaultBaseUnitTxt.getAttribute("value");
			String expDefaultBase="Dozs";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultSalesUnitTxt));
			String actDefaultSales=itemDefaultSalesUnitTxt.getAttribute("value");
			String expDefaultSales="Dozs";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultPurchaseUnitTxt));
			String actDefaultPurchase=itemDefaultPurchaseUnitTxt.getAttribute("value");
			String expDefaultPurchase="Dozs";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLengthUnitTxt));
			String actLenghtTxt=itemLengthUnitTxt.getAttribute("value");
			String expLenghtTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWindthUnitTxt));
			String actWidthTxt=itemWindthUnitTxt.getAttribute("value");
			String expWidthTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemHeightUnitTxt));
			String actHeightTxt=itemHeightUnitTxt.getAttribute("value");
			String expHeightTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWeightUnitTxt));
			String actWeightTxt=itemWeightUnitTxt.getAttribute("value");
			String expWeightTxt="0.00";
			
			
			
			System.err.println("Item Text in Units Tab");
			System.err.println("*******************************************************************************************************************************************");
			System.err.println("Default Base Unit		"		+		"	Actual		"	+	actDefaultBase			+		"	Expected		"	+	expDefaultBase);
			System.err.println("Deafult Sales Unit		"		+		"	Actual		"	+	actDefaultSales			+		"	Expected		"	+	expDefaultSales);
			System.err.println("Default Purchase Unit	"		+		"	Actual		"	+	actDefaultPurchase		+		"	Expected		"	+	expDefaultPurchase);
			System.err.println("Length					"		+		"	Actual		"	+	actLenghtTxt			+		"	Expected		"	+	expLenghtTxt);
			System.err.println("Width					"		+		"	Actual		"	+	actWidthTxt				+		"	Expected		"	+	expWidthTxt);
			System.err.println("Height					"		+		"	Actual		"	+	actHeightTxt			+		"	Expected		"	+	expHeightTxt);
			System.err.println("Weight					"		+		"	Actual		"	+	actWeightTxt			+		"	Expected		"	+	expWeightTxt);
			
			click(masterCloseBtn);
			
			if(actDefaultBase.equalsIgnoreCase(expDefaultBase) && actDefaultSales.equalsIgnoreCase(expDefaultSales) && actDefaultPurchase.equalsIgnoreCase(expDefaultPurchase)
					&& actLenghtTxt.equalsIgnoreCase(expLenghtTxt) && actWidthTxt.equalsIgnoreCase(expWidthTxt) && actHeightTxt.equalsIgnoreCase(expHeightTxt)
					&& actWeightTxt.equalsIgnoreCase(expWeightTxt)&& actItemName.equalsIgnoreCase(expItemName) && actItemCode.equalsIgnoreCase(expItemCode) && actProductType.equalsIgnoreCase(expProductType) &&actReorderLevelTxt.equalsIgnoreCase(expReorderLevelTxt)
					&& actBinCapacity.equalsIgnoreCase(expBinCapacity) && actValuationMethod.equalsIgnoreCase(expValuationMethod) && actProductMake.equalsIgnoreCase(expProductMake))
			{
				return true;
			}
			
			else
			{
				return false;
			}
			 
			 
		 }
		 
		 
		 public boolean checkUpdatingDetilsTabDataforExistingItemthroughBulkImport()  throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		 {
			
			    
			    xlfile=getBaseDir()+"\\autoIt\\ImportFiles\\ItemBulkImportNew.xlsx";
		
		
			    setCellData(xlfile, 0, 1, 6, "Raw Materials",0);
				setCellData(xlfile, 0, 1, 7, "Plastic",0);
				setCellData(xlfile, 0, 1, 8, "LG",0);
				setCellData(xlfile, 0, 1, 9, "Angola",0);
				
				setCellData(xlfile, 0, 2, 6, "Computers",0);
				setCellData(xlfile, 0, 2, 7, "Analog",0);
				setCellData(xlfile, 0, 2, 8, "Dell",0);
			
				setCellData(xlfile, 0, 2, 9, "Benin",0);
			    
				/*
				 * row=sheet.getRow(1); setCellData(6, "Raw Materials"); setCellData(7,
				 * "Plastic"); setCellData(8, "LG"); setCellData(9, "Angola");
				 * 
				 * row=sheet.getRow(2); setCellData(6, "Computers"); setCellData(7, "Analog");
				 * setCellData(8, "Dell"); setCellData(9, "Benin");
				 */
				
				 Thread.sleep(2000);
				  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
				   	toggleRibbon.click();
				   	
				   	Thread.sleep(2000);
				   	
				   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
				   	advanceMasterImportExportBtn.click();
				 	Thread.sleep(4000);
				 	Select s1=new Select(BulkImportOptionSelect);
				 	s1.selectByVisibleText("Yes");
				 	Thread.sleep(2000);
					
				 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
					FileUploadBtn.click();
					Thread.sleep(2000);
					
					 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportFromExcel.exe");
					 Thread.sleep(5000);
					 
					 click(BulkBtn);
					 click(BulkImportBtn);
					 Thread.sleep(2000);
				 	
					 String expMessage1=null;
					 String actMessage1=null;
					 
					 try {
						 
						 
						 expMessage1="Records Imported Successfully";
						 actMessage1=checkValidationMessage(expMessage1);
						 
					 }
					 catch (Exception e) {
						 System.err.println("Server Error");
						 
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
						 ServerErrorPopupCloseBtn.click();
						 Thread.sleep(2000);
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterCloseBtn));
						 AdvMasterCloseBtn.click();
						 Thread.sleep(2000);
						 
						
					}
			
					 if(actMessage1.equalsIgnoreCase(expMessage1))
					 {
						 return true;
					 }
					 else
					 {
						 return false;
					 }

		 }
		 
		 
		 @FindBy(xpath="//*[@id='ItemCategory']")
		 public static WebElement itemCategoryTxt;
		 
		 @FindBy(xpath="//*[@id='ItemSubCategory']")
		 public static WebElement itemSubCategoryTxt;
		 
		 @FindBy(xpath="//*[@id='Brand']")
		 public static WebElement itemBrandTxt;
		 
		 @FindBy(xpath="//*[@id='CountryOfOrigin']")
		 public static WebElement itemCountryOriginTxt;
		 
		 
		 public boolean checkImportedDatainDetailsTabAfterUpdatingExistingItem() throws InterruptedException
		 {
			 Thread.sleep(5000);
			 int accNameListCount = accNameList.size();
			 
			 for(int i=0;i<accNameListCount;i++)		
					
				{
					String data = accNameList.get(i).getText();
					Thread.sleep(2000);
					if(data.equals("Steel"))
					{
						Thread.sleep(2000);
						getAction().doubleClick(accChkBoxList.get(i)).build().perform();
						Thread.sleep(2000);
						break;
					}
				}
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalTab1));
			 generalTab1.click();
			 Thread.sleep(2000);;
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			String actItemName=nameTxt.getAttribute("value");
			String expItemName="Steel";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			String actItemCode=codeTxt.getAttribute("value");
			String expItemCode="Steel";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(productTypeSelect));
			Select s1=new Select(productTypeSelect);
			String actProductType=s1.getFirstSelectedOption().getText();
			String expProductType="Raw material";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reorderLevel));
			String actReorderLevelTxt=reorderLevel.getAttribute("value");
			String expReorderLevelTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binCapacity));
			String actBinCapacity=binCapacity.getAttribute("value");
			String expBinCapacity="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valuationMethodSelect));
			Select s2=new Select(valuationMethodSelect);
			String actValuationMethod=s2.getFirstSelectedOption().getText();
			String expValuationMethod="Weighted Avg";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemMakeSelect));
			Select s3=new Select(itemMakeSelect);
			String actProductMake=s3.getFirstSelectedOption().getText();
			String expProductMake="Item";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCategoryTxt));
			String actItemCategory=itemCategoryTxt.getAttribute("value");
			String expItemCategory="Raw Materials";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSubCategoryTxt));
			String actItemSubCategory=itemSubCategoryTxt.getAttribute("value");
			String expItemSubCategory="Plastic";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBrandTxt));
			String actBrand=itemBrandTxt.getAttribute("value");
			String expBrand="LG";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCountryOriginTxt));
			String actCountryOrigin=itemCountryOriginTxt.getAttribute("value");
			String expCountryOrigin="";
			
			
			System.err.println("Item Text in Details Tab");
			System.err.println("Name	"					+		"	Actual		"	+	actItemName					+		"	Expected		"	+	expItemName);
			System.err.println("Code	"					+		"	Actual		"	+	actItemCode					+		"	Expected		"	+	expItemCode);
			System.err.println("Item Type	"				+		"	Actual		"	+	actProductType				+		"	Expected		"	+	expProductType);
			System.err.println("Reorder Level"				+		"	Actual		"	+	actReorderLevelTxt			+		"	Expected		"	+	expReorderLevelTxt);
			System.err.println("Bin Capacity"				+		"	Actual		"	+	actBinCapacity				+		"	Expected		"	+	expBinCapacity);
			System.err.println("Valuation Method"			+		"	Actual		"	+	actValuationMethod			+		"	Expected		"	+	expValuationMethod);
			System.err.println("Product Make"				+		"	Actual		"	+	actProductMake				+		"	Expected		"	+	expProductMake);
			System.err.println("Item Category"				+		"	Actual		"	+	actItemCategory				+		"	Expected		"	+	expItemCategory);
			System.err.println("Item SubCategory"			+		"	Actual		"	+	actItemSubCategory			+		"	Expected		"	+	expItemSubCategory);
			System.err.println("Brand"						+		"	Actual		"	+	actBrand					+		"	Expected		"	+	expBrand);
			System.err.println("Country Origin"				+		"	Actual		"	+	actCountryOrigin			+		"	Expected		"	+	expCountryOrigin);
			
			
			click(masterCloseBtn);
			Thread.sleep(2000);
			if(actItemName.equalsIgnoreCase(expItemName) && actItemCode.equalsIgnoreCase(expItemCode) && actProductType.equalsIgnoreCase(expProductType) &&actReorderLevelTxt.equalsIgnoreCase(expReorderLevelTxt)
					&& actBinCapacity.equalsIgnoreCase(expBinCapacity) && actValuationMethod.equalsIgnoreCase(expValuationMethod) && actProductMake.equalsIgnoreCase(expProductMake)
					&& actItemCategory.equalsIgnoreCase(expItemCategory) && actItemSubCategory.equalsIgnoreCase(expItemSubCategory)
					&& actBrand.equalsIgnoreCase(expBrand) && actCountryOrigin.equalsIgnoreCase(expCountryOrigin))
			{
				return true;
			}
			
			else
			{
				return false;
			}
			
		 }
		 
		 
		 
		 public boolean checkImportedDatainDetailsTabAfterUpdatingExistingItem1() throws InterruptedException
		 {
			 
			 int accNameListCount = accNameList.size();
			 
			 for(int i=0;i<accNameListCount;i++)		
					
				{
					String data = accNameList.get(i).getText();
					Thread.sleep(2000);
					if(data.equals("QLED"))
					{
						Thread.sleep(2000);
						getAction().doubleClick(accChkBoxList.get(i)).build().perform();
						Thread.sleep(2000);
						break;
					}
				}
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalTab1));
			 generalTab1.click();
			 Thread.sleep(2000);;
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			String actItemName=nameTxt.getAttribute("value");
			String expItemName="QLED";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			String actItemCode=codeTxt.getAttribute("value");
			String expItemCode="QLED";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(productTypeSelect));
			Select s1=new Select(productTypeSelect);
			String actProductType=s1.getFirstSelectedOption().getText();
			String expProductType="Raw material";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reorderLevel));
			String actReorderLevelTxt=reorderLevel.getAttribute("value");
			String expReorderLevelTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binCapacity));
			String actBinCapacity=binCapacity.getAttribute("value");
			String expBinCapacity="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valuationMethodSelect));
			Select s2=new Select(valuationMethodSelect);
			String actValuationMethod=s2.getFirstSelectedOption().getText();
			String expValuationMethod="Weighted Avg";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemMakeSelect));
			Select s3=new Select(itemMakeSelect);
			String actProductMake=s3.getFirstSelectedOption().getText();
			String expProductMake="Item";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCategoryTxt));
			String actItemCategory=itemCategoryTxt.getAttribute("value");
			String expItemCategory="Computers";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSubCategoryTxt));
			String actItemSubCategory=itemSubCategoryTxt.getAttribute("value");
			String expItemSubCategory="Analog";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBrandTxt));
			String actBrand=itemBrandTxt.getAttribute("value");
			String expBrand="Dell";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCountryOriginTxt));
			String actCountryOrigin=itemCountryOriginTxt.getAttribute("value");
			String expCountryOrigin="";
			
			
			System.err.println("Item Text in Details Tab");
			System.err.println("Name	"					+		"	Actual		"	+	actItemName					+		"	Expected		"	+	expItemName);
			System.err.println("Code	"					+		"	Actual		"	+	actItemCode					+		"	Expected		"	+	expItemCode);
			System.err.println("Item Type	"				+		"	Actual		"	+	actProductType				+		"	Expected		"	+	expProductType);
			System.err.println("Reorder Level"				+		"	Actual		"	+	actReorderLevelTxt			+		"	Expected		"	+	expReorderLevelTxt);
			System.err.println("Bin Capacity"				+		"	Actual		"	+	actBinCapacity				+		"	Expected		"	+	expBinCapacity);
			System.err.println("Valuation Method"			+		"	Actual		"	+	actValuationMethod			+		"	Expected		"	+	expValuationMethod);
			System.err.println("Product Make"				+		"	Actual		"	+	actProductMake				+		"	Expected		"	+	expProductMake);
			System.err.println("Item Category"				+		"	Actual		"	+	actItemCategory				+		"	Expected		"	+	expItemCategory);
			System.err.println("Item SubCategory"			+		"	Actual		"	+	actItemSubCategory			+		"	Expected		"	+	expItemSubCategory);
			System.err.println("Brand"						+		"	Actual		"	+	actBrand					+		"	Expected		"	+	expBrand);
			System.err.println("Country Origin"				+		"	Actual		"	+	actCountryOrigin			+		"	Expected		"	+	expCountryOrigin);
			
			
			click(masterCloseBtn);
			Thread.sleep(2000);
			
			
			if(actItemName.equalsIgnoreCase(expItemName) && actItemCode.equalsIgnoreCase(expItemCode) && actProductType.equalsIgnoreCase(expProductType) &&actReorderLevelTxt.equalsIgnoreCase(expReorderLevelTxt)
					&& actBinCapacity.equalsIgnoreCase(expBinCapacity) && actValuationMethod.equalsIgnoreCase(expValuationMethod) && actProductMake.equalsIgnoreCase(expProductMake)
					&& actItemCategory.equalsIgnoreCase(expItemCategory) && actItemSubCategory.equalsIgnoreCase(expItemSubCategory)
					&& actBrand.equalsIgnoreCase(expBrand) && actCountryOrigin.equalsIgnoreCase(expCountryOrigin))
			{
				return true;
			}
			
			else
			{
				return false;
			}
			
		 }
		 
		 
		 
		 public boolean checkUpdatingDetilsTabDataforNewItemthroughBulkImport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		 {
			 

			
			 xlfile=getBaseDir()+"\\autoIt\\ImportFiles\\ItemBulkImportNew.xlsx";
				
			
		
				setCellData(xlfile, 0, 3, 0, "TV",1);
				setCellData(xlfile, 0, 3, 1, "TV",0);
				
				setCellData(xlfile, 0, 3, 2, "Raw Material", 0);
				setCellData(xlfile, 0, 3, 3, "Weighted Avg", 0);
				setCellData(xlfile, 0, 3, 4, "False", 0);
				setCellData(xlfile, 0, 3, 6, "Televisions", 0);
				setCellData(xlfile, 0, 3, 7, "Plasma", 0);
				setCellData(xlfile, 0, 3, 8, "Samsung", 0);
				setCellData(xlfile, 0, 3, 9, "Canada", 0);
				 
				
				
				
				 Thread.sleep(2000);
				  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
				   	toggleRibbon.click();
				   	
				   	Thread.sleep(2000);
				   	
				   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
				   	advanceMasterImportExportBtn.click();
				 	Thread.sleep(4000);
				 	Select s1=new Select(BulkImportOptionSelect);
				 	s1.selectByVisibleText("Yes");
				 	Thread.sleep(2000);
					
				 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
					FileUploadBtn.click();
					Thread.sleep(2000);
					
					 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportFromExcel.exe");
					 Thread.sleep(5000);
					 
					 click(BulkBtn);
					 click(BulkImportBtn);
					 Thread.sleep(2000);
				 	
					 String expMessage1=null;
					 String actMessage1=null;
					 
					 try {
						 
						 
						 expMessage1="Records Imported Successfully";
						 actMessage1=checkValidationMessage(expMessage1);
						 
					 }
					 catch (Exception e) {
						 System.err.println("Server Error");
						 
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
						 ServerErrorPopupCloseBtn.click();
						 Thread.sleep(2000);
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterCloseBtn));
						 AdvMasterCloseBtn.click();
						 Thread.sleep(2000);
						 
						
					}
			
					 if(actMessage1.equalsIgnoreCase(expMessage1))
					 {
						 return true;
					 }
					 else
					 {
						 return false;
					 }

		 }
		 
		 public boolean checkImportedDatainDetailsTabAfterUpdatingNewItem() throws InterruptedException
		 {
			 
			 

			 
			 int accNameListCount = accNameList.size();
			 
			 for(int i=0;i<accNameListCount;i++)		
					
				{
					String data = accNameList.get(i).getText();
					Thread.sleep(2000);
					if(data.equals("TV"))
					{
						Thread.sleep(2000);
						getAction().doubleClick(accChkBoxList.get(i)).build().perform();
						Thread.sleep(2000);
						break;
					}
				}
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalTab1));
			 generalTab1.click();
			 Thread.sleep(2000);;
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			String actItemName=nameTxt.getAttribute("value");
			String expItemName="TV";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
			String actItemCode=codeTxt.getAttribute("value");
			String expItemCode="TV";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(productTypeSelect));
			Select s1=new Select(productTypeSelect);
			String actProductType=s1.getFirstSelectedOption().getText();
			String expProductType="Raw material";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reorderLevel));
			String actReorderLevelTxt=reorderLevel.getAttribute("value");
			String expReorderLevelTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binCapacity));
			String actBinCapacity=binCapacity.getAttribute("value");
			String expBinCapacity="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valuationMethodSelect));
			Select s2=new Select(valuationMethodSelect);
			String actValuationMethod=s2.getFirstSelectedOption().getText();
			String expValuationMethod="Weighted Avg";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemMakeSelect));
			Select s3=new Select(itemMakeSelect);
			String actProductMake=s3.getFirstSelectedOption().getText();
			String expProductMake="Item";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCategoryTxt));
			String actItemCategory=itemCategoryTxt.getAttribute("value");
			String expItemCategory="Televisions";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSubCategoryTxt));
			String actItemSubCategory=itemSubCategoryTxt.getAttribute("value");
			String expItemSubCategory="Plasma";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBrandTxt));
			String actBrand=itemBrandTxt.getAttribute("value");
			String expBrand="Samsung";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCountryOriginTxt));
			String actCountryOrigin=itemCountryOriginTxt.getAttribute("value");
			String expCountryOrigin="";
			
			
			System.err.println("Item Text in Details Tab");
			System.err.println("Name	"					+		"	Actual		"	+	actItemName					+		"	Expected		"	+	expItemName);
			System.err.println("Code	"					+		"	Actual		"	+	actItemCode					+		"	Expected		"	+	expItemCode);
			System.err.println("Item Type	"				+		"	Actual		"	+	actProductType				+		"	Expected		"	+	expProductType);
			System.err.println("Reorder Level"				+		"	Actual		"	+	actReorderLevelTxt			+		"	Expected		"	+	expReorderLevelTxt);
			System.err.println("Bin Capacity"				+		"	Actual		"	+	actBinCapacity				+		"	Expected		"	+	expBinCapacity);
			System.err.println("Valuation Method"			+		"	Actual		"	+	actValuationMethod			+		"	Expected		"	+	expValuationMethod);
			System.err.println("Product Make"				+		"	Actual		"	+	actProductMake				+		"	Expected		"	+	expProductMake);
			System.err.println("Item Category"				+		"	Actual		"	+	actItemCategory				+		"	Expected		"	+	expItemCategory);
			System.err.println("Item SubCategory"			+		"	Actual		"	+	actItemSubCategory			+		"	Expected		"	+	expItemSubCategory);
			System.err.println("Brand"						+		"	Actual		"	+	actBrand					+		"	Expected		"	+	expBrand);
			System.err.println("Country Origin"				+		"	Actual		"	+	actCountryOrigin			+		"	Expected		"	+	expCountryOrigin);
			
			
			click(masterCloseBtn);
			Thread.sleep(2000);
			if(actItemName.equalsIgnoreCase(expItemName) && actItemCode.equalsIgnoreCase(expItemCode) && actProductType.equalsIgnoreCase(expProductType) &&actReorderLevelTxt.equalsIgnoreCase(expReorderLevelTxt)
					&& actBinCapacity.equalsIgnoreCase(expBinCapacity) && actValuationMethod.equalsIgnoreCase(expValuationMethod) && actProductMake.equalsIgnoreCase(expProductMake)
					&& actItemCategory.equalsIgnoreCase(expItemCategory) && actItemSubCategory.equalsIgnoreCase(expItemSubCategory)
					&& actBrand.equalsIgnoreCase(expBrand) && actCountryOrigin.equalsIgnoreCase(expCountryOrigin))
			{
				return true;
			}
			
			else
			{
				return false;
			}
			
		 
		 }
		 
		 
		 public boolean checkUpdatingUnitsTabDataforExistingItemthroughBulkImport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		 {
			 xlfile=getBaseDir()+"\\autoIt\\ImportFiles\\ItemBulkImportNew.xlsx";
			
				setCellData(xlfile, 0, 1, 10, "Dozs",0);
				setCellData(xlfile, 0, 1, 11, "Dozs",0);
				setCellData(xlfile, 0, 1, 12, "Dozs",0);
				
				
				setCellData(xlfile, 0, 3, 10, "Pcs",0);
				setCellData(xlfile, 0, 3, 11, "Pcs",0);
				setCellData(xlfile, 0, 3, 12, "Pcs",0);
				
				
				 Thread.sleep(5500);
				  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
				   	toggleRibbon.click();
				   	
				   	Thread.sleep(2000);
				   	
				   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
				   	advanceMasterImportExportBtn.click();
				 	Thread.sleep(4000);
				 	Select s1=new Select(BulkImportOptionSelect);
				 	s1.selectByVisibleText("Yes");
				 	Thread.sleep(2000);
					
				 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
					FileUploadBtn.click();
					Thread.sleep(2000);
					
					 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportFromExcel.exe");
					 Thread.sleep(5000);
					 
					 click(BulkBtn);
					 click(BulkImportBtn);
					 Thread.sleep(2000);
				 	
					 String expMessage1=null;
					 String actMessage1=null;
					 
					 try {
						 
						 
						 expMessage1="Records Imported Successfully";
						 actMessage1=checkValidationMessage(expMessage1);
						 
					 }
					 catch (Exception e) {
						 System.err.println("Server Error");
						 
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
						 ServerErrorPopupCloseBtn.click();
						 Thread.sleep(2000);
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterCloseBtn));
						 AdvMasterCloseBtn.click();
						 Thread.sleep(2000);
						 
						
					}
			
					 if(actMessage1.equalsIgnoreCase(expMessage1))
					 {
						 return true;
					 }
					 else
					 {
						 return false;
					 }

		 
		 }
		 
		 
		 public boolean checkUpdatingUnitsTabDataforNewItemthroughBulkImport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		 {
			
			 xlfile=getBaseDir()+"\\autoIt\\ImportFiles\\ItemBulkImportNew.xlsx";
		
				setCellData(xlfile, 0, 4, 0, "Mobile",1);
				setCellData(xlfile, 0, 4, 1, "Mobile",0);
				setCellData(xlfile, 0, 4, 2, "Raw Material",0);
				setCellData(xlfile, 0, 4, 3, "Weighted Avg",0);
				setCellData(xlfile, 0, 4, 4, "False",0);
				setCellData(xlfile, 0, 4, 6, "Accessories",0);
				setCellData(xlfile, 0, 4, 7, "Watches",0);
				setCellData(xlfile, 0, 4, 8, "Samsung",0);
				setCellData(xlfile, 0, 4, 9, "India",0);
				setCellData(xlfile, 0, 4, 10, "Pcs",0);
				setCellData(xlfile, 0, 4, 11, "Dozs",0);
				setCellData(xlfile, 0, 4, 12, "Pcs",0);
				
				 Thread.sleep(2000);
				  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
				   	toggleRibbon.click();
				   	
				   	Thread.sleep(2000);
				   	
				   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
				   	advanceMasterImportExportBtn.click();
				 	Thread.sleep(4000);
				 	Select s1=new Select(BulkImportOptionSelect);
				 	s1.selectByVisibleText("Yes");
				 	Thread.sleep(2000);
					
				 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
					FileUploadBtn.click();
					Thread.sleep(2000);
					
					 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportFromExcel.exe");
					 Thread.sleep(7000);
					 
					 click(BulkBtn);
					 Thread.sleep(2000);
					 click(BulkImportBtn);
					 Thread.sleep(2000);
				 	
					 String expMessage1=null;
					 String actMessage1=null;
					 
					 try {
						 
						 
						 expMessage1="Records Imported Successfully";
						 actMessage1=checkValidationMessage(expMessage1);
						 
					 }
					 catch (Exception e) {
						 System.err.println("Server Error");
						 
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
						 ServerErrorPopupCloseBtn.click();
						 Thread.sleep(2000);
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterCloseBtn));
						 AdvMasterCloseBtn.click();
						 Thread.sleep(2000);
						 
						
					}
			
					 if(actMessage1.equalsIgnoreCase(expMessage1))
					 {
						 return true;
					 }
					 else
					 {
						 return false;
					 }

		 
		 }
		 
		 
		 public boolean checkImportedDatainUnitsTabAfterUpdatingExistingItem() throws InterruptedException
		 {

			 int accNameListCount = accNameList.size();
			 
			 for(int i=0;i<accNameListCount;i++)		
					
				{
					String data = accNameList.get(i).getText();
					Thread.sleep(2000);
					if(data.equals("TV"))
					{
						Thread.sleep(2000);
						getAction().doubleClick(accChkBoxList.get(i)).build().perform();
						Thread.sleep(2000);
						break;
					}
				}
			 
			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitsTab1));
			 unitsTab1.click();
			 Thread.sleep(3000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultBaseUnitTxt));
			String actDefaultBase=itemDefaultBaseUnitTxt.getAttribute("value");
			String expDefaultBase="Pcs";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultSalesUnitTxt));
			String actDefaultSales=itemDefaultSalesUnitTxt.getAttribute("value");
			String expDefaultSales="Pcs";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultPurchaseUnitTxt));
			String actDefaultPurchase=itemDefaultPurchaseUnitTxt.getAttribute("value");
			String expDefaultPurchase="Pcs";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLengthUnitTxt));
			String actLenghtTxt=itemLengthUnitTxt.getAttribute("value");
			String expLenghtTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWindthUnitTxt));
			String actWidthTxt=itemWindthUnitTxt.getAttribute("value");
			String expWidthTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemHeightUnitTxt));
			String actHeightTxt=itemHeightUnitTxt.getAttribute("value");
			String expHeightTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWeightUnitTxt));
			String actWeightTxt=itemWeightUnitTxt.getAttribute("value");
			String expWeightTxt="0.00";
			
			
			
			System.err.println("Item Text in Units Tab");
			System.err.println("*******************************************************************************************************************************************");
			System.err.println("Default Base Unit		"		+		"	Actual		"	+	actDefaultBase			+		"	Expected		"	+	expDefaultBase);
			System.err.println("Deafult Sales Unit		"		+		"	Actual		"	+	actDefaultSales			+		"	Expected		"	+	expDefaultSales);
			System.err.println("Default Purchase Unit	"		+		"	Actual		"	+	actDefaultPurchase		+		"	Expected		"	+	expDefaultPurchase);
			System.err.println("Length					"		+		"	Actual		"	+	actLenghtTxt			+		"	Expected		"	+	expLenghtTxt);
			System.err.println("Width					"		+		"	Actual		"	+	actWidthTxt				+		"	Expected		"	+	expWidthTxt);
			System.err.println("Height					"		+		"	Actual		"	+	actHeightTxt			+		"	Expected		"	+	expHeightTxt);
			System.err.println("Weight					"		+		"	Actual		"	+	actWeightTxt			+		"	Expected		"	+	expWeightTxt);
			
			
			click(masterCloseBtn);
			Thread.sleep(2000);
			
			if(actDefaultBase.equalsIgnoreCase(expDefaultBase) && actDefaultSales.equalsIgnoreCase(expDefaultSales) && actDefaultPurchase.equalsIgnoreCase(expDefaultPurchase)
					&& actLenghtTxt.equalsIgnoreCase(expLenghtTxt) && actWidthTxt.equalsIgnoreCase(expWidthTxt) && actHeightTxt.equalsIgnoreCase(expHeightTxt)
					&& actWeightTxt.equalsIgnoreCase(expWeightTxt))
			{
				return true;
			}
			
			else
			{
				return false;
			}
			
			
		
		 }
		 
		 
		 public boolean checkImportedDatainUnitsTabAfterUpdatingExistingItem1() throws InterruptedException
		 {

			 int accNameListCount = accNameList.size();
			 
			 for(int i=0;i<accNameListCount;i++)		
					
				{
					String data = accNameList.get(i).getText();
					Thread.sleep(2000);
					if(data.equals("Steel"))
					{
						Thread.sleep(2000);
						getAction().doubleClick(accChkBoxList.get(i)).build().perform();
						Thread.sleep(2000);
						break;
					}
				}
			 
			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitsTab1));
			 unitsTab1.click();
			 Thread.sleep(3000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultBaseUnitTxt));
			String actDefaultBase=itemDefaultBaseUnitTxt.getAttribute("value");
			String expDefaultBase="Dozs";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultSalesUnitTxt));
			String actDefaultSales=itemDefaultSalesUnitTxt.getAttribute("value");
			String expDefaultSales="Dozs";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultPurchaseUnitTxt));
			String actDefaultPurchase=itemDefaultPurchaseUnitTxt.getAttribute("value");
			String expDefaultPurchase="Dozs";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLengthUnitTxt));
			String actLenghtTxt=itemLengthUnitTxt.getAttribute("value");
			String expLenghtTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWindthUnitTxt));
			String actWidthTxt=itemWindthUnitTxt.getAttribute("value");
			String expWidthTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemHeightUnitTxt));
			String actHeightTxt=itemHeightUnitTxt.getAttribute("value");
			String expHeightTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWeightUnitTxt));
			String actWeightTxt=itemWeightUnitTxt.getAttribute("value");
			String expWeightTxt="0.00";
			
			
			
			System.err.println("Item Text in Units Tab");
			System.err.println("*******************************************************************************************************************************************");
			System.err.println("Default Base Unit		"		+		"	Actual		"	+	actDefaultBase			+		"	Expected		"	+	expDefaultBase);
			System.err.println("Deafult Sales Unit		"		+		"	Actual		"	+	actDefaultSales			+		"	Expected		"	+	expDefaultSales);
			System.err.println("Default Purchase Unit	"		+		"	Actual		"	+	actDefaultPurchase		+		"	Expected		"	+	expDefaultPurchase);
			System.err.println("Length					"		+		"	Actual		"	+	actLenghtTxt			+		"	Expected		"	+	expLenghtTxt);
			System.err.println("Width					"		+		"	Actual		"	+	actWidthTxt				+		"	Expected		"	+	expWidthTxt);
			System.err.println("Height					"		+		"	Actual		"	+	actHeightTxt			+		"	Expected		"	+	expHeightTxt);
			System.err.println("Weight					"		+		"	Actual		"	+	actWeightTxt			+		"	Expected		"	+	expWeightTxt);
			
			
			click(masterCloseBtn);
			Thread.sleep(2000);
			if(actDefaultBase.equalsIgnoreCase(expDefaultBase) && actDefaultSales.equalsIgnoreCase(expDefaultSales) && actDefaultPurchase.equalsIgnoreCase(expDefaultPurchase)
					&& actLenghtTxt.equalsIgnoreCase(expLenghtTxt) && actWidthTxt.equalsIgnoreCase(expWidthTxt) && actHeightTxt.equalsIgnoreCase(expHeightTxt)
					&& actWeightTxt.equalsIgnoreCase(expWeightTxt))
			{
				return true;
			}
			
			else
			{
				return false;
			}
		 }
		 
		 
		 public boolean checkImportedDatainUnitsTabAfterUpdatingNewItem() throws InterruptedException
		 {
			 
			 Thread.sleep(5000);
			 int accNameListCount = accNameList.size();
			 
			 for(int i=0;i<accNameListCount;i++)		
					
				{
					String data = accNameList.get(i).getText();
					Thread.sleep(2000);
					if(data.equals("Mobile"))
					{
						Thread.sleep(2000);
						getAction().doubleClick(accChkBoxList.get(i)).build().perform();
						Thread.sleep(2000);
						break;
					}
				}
			 
			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitsTab1));
			 unitsTab1.click();
			 Thread.sleep(3000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultBaseUnitTxt));
			String actDefaultBase=itemDefaultBaseUnitTxt.getAttribute("value");
			String expDefaultBase="Pcs";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultSalesUnitTxt));
			String actDefaultSales=itemDefaultSalesUnitTxt.getAttribute("value");
			String expDefaultSales="Dozs";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemDefaultPurchaseUnitTxt));
			String actDefaultPurchase=itemDefaultPurchaseUnitTxt.getAttribute("value");
			String expDefaultPurchase="Pcs";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLengthUnitTxt));
			String actLenghtTxt=itemLengthUnitTxt.getAttribute("value");
			String expLenghtTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWindthUnitTxt));
			String actWidthTxt=itemWindthUnitTxt.getAttribute("value");
			String expWidthTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemHeightUnitTxt));
			String actHeightTxt=itemHeightUnitTxt.getAttribute("value");
			String expHeightTxt="0.00";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWeightUnitTxt));
			String actWeightTxt=itemWeightUnitTxt.getAttribute("value");
			String expWeightTxt="0.00";
			
			
			
			System.err.println("Item Text in Units Tab");
			System.err.println("*******************************************************************************************************************************************");
			System.err.println("Default Base Unit		"		+		"	Actual		"	+	actDefaultBase			+		"	Expected		"	+	expDefaultBase);
			System.err.println("Deafult Sales Unit		"		+		"	Actual		"	+	actDefaultSales			+		"	Expected		"	+	expDefaultSales);
			System.err.println("Default Purchase Unit	"		+		"	Actual		"	+	actDefaultPurchase		+		"	Expected		"	+	expDefaultPurchase);
			System.err.println("Length					"		+		"	Actual		"	+	actLenghtTxt			+		"	Expected		"	+	expLenghtTxt);
			System.err.println("Width					"		+		"	Actual		"	+	actWidthTxt				+		"	Expected		"	+	expWidthTxt);
			System.err.println("Height					"		+		"	Actual		"	+	actHeightTxt			+		"	Expected		"	+	expHeightTxt);
			System.err.println("Weight					"		+		"	Actual		"	+	actWeightTxt			+		"	Expected		"	+	expWeightTxt);
			
			click(masterCloseBtn);
			Thread.sleep(2000);
			if(actDefaultBase.equalsIgnoreCase(expDefaultBase) && actDefaultSales.equalsIgnoreCase(expDefaultSales) && actDefaultPurchase.equalsIgnoreCase(expDefaultPurchase)
					&& actLenghtTxt.equalsIgnoreCase(expLenghtTxt) && actWidthTxt.equalsIgnoreCase(expWidthTxt) && actHeightTxt.equalsIgnoreCase(expHeightTxt)
					&& actWeightTxt.equalsIgnoreCase(expWeightTxt))
			{
				return true;
			}
			
			else
			{
				return false;
			}

		 }
		 
		 
		 
		 public boolean checkUpdatingOtherDetailsTabforExistingItemthroughBulkImport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		 {
			
			 xlfile=getBaseDir()+"\\autoIt\\ImportFiles\\ItemBulkImportNew.xlsx";
				
			
		
				setCellData(xlfile, 0, 4, 31, "COGS POSTING ACC",0);
				setCellData(xlfile, 0, 4, 32, "STD RATE COGS ACC INV",0);
				setCellData(xlfile, 0, 4, 33, "Sales - Computers",0);
				setCellData(xlfile, 0, 4, 35, "SHORTAGE COGS POSTING ACC",0);
				setCellData(xlfile, 0, 4, 36, "EXCESS COGS POSTING ACC",0);
				setCellData(xlfile, 0, 4, 37, "SR COGS POSTING ACC",0);
				setCellData(xlfile, 0, 4, 38, "PURCHASE VARIANCE",0);
				
				 Thread.sleep(2000);
				  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
				   	toggleRibbon.click();
				   	
				   	Thread.sleep(2000);
				   	
				   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
				   	advanceMasterImportExportBtn.click();
				 	Thread.sleep(4000);
				 	Select s1=new Select(BulkImportOptionSelect);
				 	s1.selectByVisibleText("Yes");
				 	Thread.sleep(2000);
					
				 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
					FileUploadBtn.click();
					Thread.sleep(2000);
					
					 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportFromExcel.exe");
					 Thread.sleep(5000);
					 
					 click(BulkBtn);
					 Thread.sleep(2000);
					 click(BulkImportBtn);
					 Thread.sleep(2000);
				 	
					 String expMessage1=null;
					 String actMessage1=null;
					 
					 try {
						 
						 
						 expMessage1="Records Imported Successfully";
						 actMessage1=checkValidationMessage(expMessage1);
						 
					 }
					 catch (Exception e) {
						 System.err.println("Server Error");
						 
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
						 ServerErrorPopupCloseBtn.click();
						 Thread.sleep(2000);
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterCloseBtn));
						 AdvMasterCloseBtn.click();
						 Thread.sleep(2000);
						 
						
					}
			
					 if(actMessage1.equalsIgnoreCase(expMessage1))
					 {
						 return true;
					 }
					 else
					 {
						 return false;
					 }

		 }
		 
		 
		 public boolean checkImportedDatainOtherDetailsTabAfterUpdatingExistingItem() throws InterruptedException
		 {

			 int accNameListCount = accNameList.size();
			 
			 for(int i=0;i<accNameListCount;i++)		
					
				{
					String data = accNameList.get(i).getText();
					Thread.sleep(2000);
					if(data.equals("Mobile"))
					{
						Thread.sleep(2000);
						getAction().doubleClick(accChkBoxList.get(i)).build().perform();
						Thread.sleep(2000);
						break;
					}
				}
			 

			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(otherDetailsTab1));
			 otherDetailsTab1.click();
			 Thread.sleep(3000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofIssueAccountTxt));
			String actCostofIssueTxt=itemCostofIssueAccountTxt.getAttribute("value");
			String expCostofIssueTxt="COGS POSTING ACC";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemStocksAccountTxt));
			String actStocksAccountTxt=itemStocksAccountTxt.getAttribute("value");
			String expStocksAccountTxt="STD RATE COGS ACC INV";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSalesAccountTxt));
			String actSalesAccountTxt=itemSalesAccountTxt.getAttribute("value");
			String expSalesAccountTxt="Sales-Computers";	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWIPAccountTxt));
			String actWIPAccountTxt=Boolean.toString(itemWIPAccountTxt.getAttribute("value").isEmpty());
			String expWIPAccountTxt="true";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofShortageStockAccTxt));
			String actCostofShortageStockAccTxt=itemCostofShortageStockAccTxt.getAttribute("value");
			String expCostofShortageStockAccTxt="SHORTAGE COGS POSTING ACC";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofExcessStockAccTxt));
			String actCostofExcessStockAccTxt=itemCostofExcessStockAccTxt.getAttribute("value");
			String expCostofExcessStockAccTxt="EXCESS COGS POSTING ACC";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofSalesReturnAccTxt));
			String actCostofSalesReturnAccTxt=itemCostofSalesReturnAccTxt.getAttribute("value");
			String expCostofSalesReturnAccTxt="SR COGS POSTING ACC";	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemPurchaseVarianceAccTxt));
			String actPurchaseVarianceAccTxt=itemPurchaseVarianceAccTxt.getAttribute("value");
			String expPurchaseVarianceAccTxt="PURCHASE VARIANCE";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemVATExpenseAccTxt));
			String actVATExpenseAccTxt=Boolean.toString(itemVATExpenseAccTxt.getAttribute("value").isEmpty());
			String expVATExpenseAccTxt="true";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemProductionLeadTimeTxt));
			String actProductionLeadTimeTxt=itemProductionLeadTimeTxt.getAttribute("value");
			String expProductionLeadTimeTxt="0";	
			
			
			System.err.println("Item Text in Other Details Tab");
			System.err.println("********************************************************************************************************************************************************************************");
			System.err.println("Cost of Issue Account				"		+		"	Actual		"	+	actCostofIssueTxt						+		"	Expected		"	+	expCostofIssueTxt);
			System.err.println("Stocks Account						"		+		"	Actual		"	+	actStocksAccountTxt						+		"	Expected		"	+	expStocksAccountTxt);
			System.err.println("Sales Account						"		+		"	Actual		"	+	actSalesAccountTxt						+		"	Expected		"	+	expSalesAccountTxt);
			System.err.println("WIP Account							"		+		"	Actual		"	+	actWIPAccountTxt						+		"	Expected		"	+	expWIPAccountTxt);
			System.err.println("Cost of Shortage Stock A/C			"		+		"	Actual		"	+	actCostofShortageStockAccTxt			+		"	Expected		"	+	expCostofShortageStockAccTxt);
			System.err.println("Cost of Excess Stock A/C			"		+		"	Actual		"	+	actCostofExcessStockAccTxt				+		"	Expected		"	+	expCostofExcessStockAccTxt);
			System.err.println("Cost of Sale Return A/C				"		+		"	Actual		"	+	actCostofSalesReturnAccTxt				+		"	Expected		"	+	expCostofSalesReturnAccTxt);
			System.err.println("Purchase Varience A/C				"		+		"	Actual		"	+	actPurchaseVarianceAccTxt				+		"	Expected		"	+	expPurchaseVarianceAccTxt);
			System.err.println("VAT Expense Account					"		+		"	Actual		"	+	actVATExpenseAccTxt						+		"	Expected		"	+	expVATExpenseAccTxt);
			System.err.println("Production Lead Time				"		+		"	Actual		"	+	actProductionLeadTimeTxt				+		"	Expected		"	+	expProductionLeadTimeTxt);
		
			click(masterCloseBtn);
			Thread.sleep(2000);
			
			if(actCostofIssueTxt.equalsIgnoreCase(expCostofIssueTxt) && actStocksAccountTxt.equalsIgnoreCase(expStocksAccountTxt) &&actSalesAccountTxt.equalsIgnoreCase(expSalesAccountTxt)
					&&actWIPAccountTxt.equalsIgnoreCase(expWIPAccountTxt) && actCostofShortageStockAccTxt.equalsIgnoreCase(expCostofShortageStockAccTxt) && actCostofExcessStockAccTxt.equalsIgnoreCase(expCostofExcessStockAccTxt)
					&& actCostofSalesReturnAccTxt.equalsIgnoreCase(expCostofSalesReturnAccTxt) && actPurchaseVarianceAccTxt.equalsIgnoreCase(expPurchaseVarianceAccTxt)
					&& actVATExpenseAccTxt.equalsIgnoreCase(expVATExpenseAccTxt) &&  actProductionLeadTimeTxt.equalsIgnoreCase(expProductionLeadTimeTxt))
			{
				return true;
			}
			else
			{
				return false;
			}
			
		
		 }
		 
		 
		 public boolean checkUpdatingOtherDetailsTabDataforNewItemthroughBulkImport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		 {
			 

			
			 xlfile=getBaseDir()+"\\autoIt\\ImportFiles\\ItemBulkImportNew.xlsx";
				
		
		
				setCellData(xlfile, 0, 5, 0, "Table",1);
				setCellData(xlfile, 0, 5, 1, "Table",0);
				setCellData(xlfile, 0, 5, 2, "Raw Material",0);
				setCellData(xlfile, 0, 5, 3, "Weighted Avg",0);
				setCellData(xlfile, 0, 5, 4, "False",0);
				setCellData(xlfile, 0, 5, 6, "Accessories",0);
				setCellData(xlfile, 0, 5, 7, "Furniture",0);
				setCellData(xlfile, 0, 5, 8, "Samsung",0);
				setCellData(xlfile, 0, 5, 9, "India",0);
				setCellData(xlfile, 0, 5, 10, "Pcs",0);
				setCellData(xlfile, 0, 5, 11, "Dozs",0);
				setCellData(xlfile, 0, 5, 12, "Pcs",0);
				setCellData(xlfile, 0, 5, 31, "COGS POSTING ACC",0);
				setCellData(xlfile, 0, 5, 32, "STD RATE COGS ACC INV",0);
				setCellData(xlfile, 0, 5, 33, "Sales - Computers",0);
				setCellData(xlfile, 0, 5, 35, "SHORTAGE COGS POSTING ACC",0);
				setCellData(xlfile, 0, 5, 36, "EXCESS COGS POSTING ACC",0);
				setCellData(xlfile, 0, 5, 37, "SR COGS POSTING ACC",0);
				setCellData(xlfile, 0, 5, 38, "PURCHASE VARIANCE",0);
				
				 Thread.sleep(2000);
				  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
				   	toggleRibbon.click();
				   	
				   	Thread.sleep(2000);
				   	
				   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
				   	advanceMasterImportExportBtn.click();
				 	Thread.sleep(4000);
				 	Select s1=new Select(BulkImportOptionSelect);
				 	s1.selectByVisibleText("Yes");
				 	Thread.sleep(2000);
					
				 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
					FileUploadBtn.click();
					Thread.sleep(2000);
					
					 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportFromExcel.exe");
					 Thread.sleep(7000);
					 
					 click(BulkBtn);
					 Thread.sleep(2000);
					 click(BulkImportBtn);
					 Thread.sleep(2000);
				 	
					 String expMessage1=null;
					 String actMessage1=null;
					 
					 try {
						 
						 
						 expMessage1="Records Imported Successfully";
						 actMessage1=checkValidationMessage(expMessage1);
						 
					 }
					 catch (Exception e) {
						 System.err.println("Server Error");
						 
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
						 ServerErrorPopupCloseBtn.click();
						 Thread.sleep(2000);
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterCloseBtn));
						 AdvMasterCloseBtn.click();
						 Thread.sleep(2000);
						 
						
					}
			
					 if(actMessage1.equalsIgnoreCase(expMessage1))
					 {
						 return true;
					 }
					 else
					 {
						 return false;
					 }

		 }
		 
		 
		 
		 public boolean checkImportedDatainOtherDetailsTabAfterUpdatingNewItem() throws InterruptedException
		 {

			 int accNameListCount = accNameList.size();
			 
			 for(int i=0;i<accNameListCount;i++)		
					
				{
					String data = accNameList.get(i).getText();
					Thread.sleep(2000);
					if(data.equals("Table"))
					{
						Thread.sleep(2000);
						getAction().doubleClick(accChkBoxList.get(i)).build().perform();
						Thread.sleep(2000);
						break;
					}
				}
			 

			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(otherDetailsTab1));
			 otherDetailsTab1.click();
			 Thread.sleep(3000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofIssueAccountTxt));
			String actCostofIssueTxt=itemCostofIssueAccountTxt.getAttribute("value");
			String expCostofIssueTxt="COGS POSTING ACC";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemStocksAccountTxt));
			String actStocksAccountTxt=itemStocksAccountTxt.getAttribute("value");
			String expStocksAccountTxt="STD RATE COGS ACC INV";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSalesAccountTxt));
			String actSalesAccountTxt=itemSalesAccountTxt.getAttribute("value");
			String expSalesAccountTxt="Sales - Computers";	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemWIPAccountTxt));
			String actWIPAccountTxt=Boolean.toString(itemWIPAccountTxt.getAttribute("value").isEmpty());
			String expWIPAccountTxt="true";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofShortageStockAccTxt));
			String actCostofShortageStockAccTxt=itemCostofShortageStockAccTxt.getAttribute("value");
			String expCostofShortageStockAccTxt="SHORTAGE COGS POSTING ACC";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofExcessStockAccTxt));
			String actCostofExcessStockAccTxt=itemCostofExcessStockAccTxt.getAttribute("value");
			String expCostofExcessStockAccTxt="EXCESS COGS POSTING ACC";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemCostofSalesReturnAccTxt));
			String actCostofSalesReturnAccTxt=itemCostofSalesReturnAccTxt.getAttribute("value");
			String expCostofSalesReturnAccTxt="SR COGS POSTING ACC";	
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemPurchaseVarianceAccTxt));
			String actPurchaseVarianceAccTxt=itemPurchaseVarianceAccTxt.getAttribute("value");
			String expPurchaseVarianceAccTxt="PURCHASE VARIANCE";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemVATExpenseAccTxt));
			String actVATExpenseAccTxt=Boolean.toString(itemVATExpenseAccTxt.getAttribute("value").isEmpty());
			String expVATExpenseAccTxt="true";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemProductionLeadTimeTxt));
			String actProductionLeadTimeTxt=itemProductionLeadTimeTxt.getAttribute("value");
			String expProductionLeadTimeTxt="0";	
			
			
			System.err.println("Item Text in Other Details Tab");
			System.err.println("********************************************************************************************************************************************************************************");
			System.err.println("Cost of Issue Account			"		+		"	Actual		"	+	actCostofIssueTxt						+		"	Expected		"	+	expCostofIssueTxt);
			System.err.println("Stocks Account					"		+		"	Actual		"	+	actStocksAccountTxt						+		"	Expected		"	+	expStocksAccountTxt);
			System.err.println("Sales Account					"		+		"	Actual		"	+	actSalesAccountTxt						+		"	Expected		"	+	expSalesAccountTxt);
			System.err.println("WIP Account						"		+		"	Actual		"	+	actWIPAccountTxt						+		"	Expected		"	+	expWIPAccountTxt);
			System.err.println("Cost of Shortage Stock A/C		"		+		"	Actual		"	+	actCostofShortageStockAccTxt			+		"	Expected		"	+	expCostofShortageStockAccTxt);
			System.err.println("Cost of Excess Stock A/C		"		+		"	Actual		"	+	actCostofExcessStockAccTxt				+		"	Expected		"	+	expCostofExcessStockAccTxt);
			System.err.println("Cost of Sale Return A/C			"		+		"	Actual		"	+	actCostofSalesReturnAccTxt				+		"	Expected		"	+	expCostofSalesReturnAccTxt);
			System.err.println("Purchase Varience A/C			"		+		"	Actual		"	+	actPurchaseVarianceAccTxt				+		"	Expected		"	+	expPurchaseVarianceAccTxt);
			System.err.println("VAT Expense Account				"		+		"	Actual		"	+	actVATExpenseAccTxt						+		"	Expected		"	+	expVATExpenseAccTxt);
			System.err.println("Production Lead Time			"		+		"	Actual		"	+	actProductionLeadTimeTxt				+		"	Expected		"	+	expProductionLeadTimeTxt);
		
			click(masterCloseBtn);
			Thread.sleep(2000);
			
			if(actCostofIssueTxt.equalsIgnoreCase(expCostofIssueTxt) && actStocksAccountTxt.equalsIgnoreCase(expStocksAccountTxt) &&actSalesAccountTxt.equalsIgnoreCase(expSalesAccountTxt)
					&&actWIPAccountTxt.equalsIgnoreCase(expWIPAccountTxt) && actCostofShortageStockAccTxt.equalsIgnoreCase(expCostofShortageStockAccTxt) && actCostofExcessStockAccTxt.equalsIgnoreCase(expCostofExcessStockAccTxt)
					&& actCostofSalesReturnAccTxt.equalsIgnoreCase(expCostofSalesReturnAccTxt) && actPurchaseVarianceAccTxt.equalsIgnoreCase(expPurchaseVarianceAccTxt)
					&& actVATExpenseAccTxt.equalsIgnoreCase(expVATExpenseAccTxt) &&  actProductionLeadTimeTxt.equalsIgnoreCase(expProductionLeadTimeTxt))
			{
				return true;
			}
			else
			{
				return false;
			}
			
		
		 }
		 
		 public boolean checkUpdatingGSTTabDataforNewItemThroughbulkImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		 {

			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 xlfile=getBaseDir()+"\\autoIt\\ImportFiles\\ItemBulkImportGST.xlsx";
				
			 	setCellData(xlfile, 0, 1, 0, "Coke",1);
				setCellData(xlfile, 0, 1, 1, "Coke",0);
				setCellData(xlfile, 0, 1, 2, "Raw Material",0);
				setCellData(xlfile, 0, 1, 3, "Weighted Avg",0);
				setCellData(xlfile, 0, 1, 4, "False",0);
				
				setCellData(xlfile, 0, 1, 23, "Coke",0);
				setCellData(xlfile, 0, 1, 24, "20",0);
				setCellData(xlfile, 0, 1, 25, "15.25",0);
				
				 Thread.sleep(5500);
				  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
				   	toggleRibbon.click();
				   	
				   	Thread.sleep(2000);
				   	
				   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
				   	advanceMasterImportExportBtn.click();
				 	Thread.sleep(4000);
				 	Select s1=new Select(BulkImportOptionSelect);
				 	s1.selectByVisibleText("Yes");
				 	Thread.sleep(2000);
					
				 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
					FileUploadBtn.click();
					Thread.sleep(2000);
					
					 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportWithExtraTab.exe");
					 Thread.sleep(7000);
					 
					 click(BulkBtn);
					 Thread.sleep(2000);
					 click(BulkImportBtn);
					 Thread.sleep(2000);
				 	
					 String expMessage1=null;
					 String actMessage1=null;
					 
					 try {
						 
						 
						 expMessage1="Records Imported Successfully";
						 actMessage1=checkValidationMessage(expMessage1);
						 
					 }
					 catch (Exception e) {
						 System.err.println("Server Error");
						 
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
						 ServerErrorPopupCloseBtn.click();
						 Thread.sleep(2000);
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterCloseBtn));
						 AdvMasterCloseBtn.click();
						 Thread.sleep(2000);
						 
						
					}
			
					 if(actMessage1.equalsIgnoreCase(expMessage1))
					 {
						 return true;
					 }
					 else
					 {
						 return false;
					 }

		 }
		 
		 @FindBy(xpath="//*[@id='divBtnGroup9']")
		 public static WebElement GSTTab;
		 
		 @FindBy(xpath="//*[@id='ItemName']")
		 public static WebElement gst_ItemName;
		 
		 @FindBy(xpath="//*[@id='ItemCost']")
		 public static WebElement gst_ItemCost;
		 
		 @FindBy(xpath="//*[@id='GSTRate']")
		 public static WebElement gst_GSTRate;
		 
		 public boolean checkImportedDatainGSTTabAfterUpdatingNewItem() throws InterruptedException
		 {
			 Thread.sleep(2000);

				 int accNameListCount = accNameList.size();
				 
				 for(int i=0;i<accNameListCount;i++)		
						
					{
						String data = accNameList.get(i).getText();
						Thread.sleep(2000);
						if(data.equals("Coke"))
						{
							Thread.sleep(2000);
							getAction().doubleClick(accChkBoxList.get(i)).build().perform();
							Thread.sleep(2000);
							break;
						}
					}
				 

				 Thread.sleep(6000);
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GSTTab));
				 GSTTab.click();
				 Thread.sleep(3000);
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gst_ItemName));
				String actgst_ItemName=gst_ItemName.getAttribute("value");
				String expgst_ItemName="Coke";
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gst_ItemCost));
				String actgst_ItemCost=gst_ItemCost.getAttribute("value");
				String expgst_ItemCost="20";
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gst_GSTRate));
				String actgst_GSTRate=gst_GSTRate.getAttribute("value");
				String expgst_GSTRate="15.25";	
				
			
				
				System.err.println("Item Text in GST Tab");
				System.err.println("********************************************************************************************************************************************************************************");
				System.err.println("Item Name		"		+		"	Actual		"	+	actgst_ItemName			+		"	Expected		"	+	expgst_ItemName);
				System.err.println("Item Cost"		+		"	Actual		"	+	actgst_ItemCost						+		"	Expected		"	+	expgst_ItemCost);
				System.err.println("GST Rate"		+		"	Actual		"	+	actgst_GSTRate						+		"	Expected		"	+	expgst_GSTRate);
				
				click(masterCloseBtn);
				Thread.sleep(2000);
				
				if(actgst_ItemName.equalsIgnoreCase(expgst_ItemName) && actgst_ItemCost.equalsIgnoreCase(expgst_ItemCost) && actgst_GSTRate.equalsIgnoreCase(expgst_GSTRate))
				{
					return true;
				}
				else
				{
					return false;
				}
		 }
		 
		 
		 public boolean checkUpdatingGSTTabDataforExistingItemThroughbulkImport() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		 {

			 
			 Thread.sleep(5500);
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 xlfile=getBaseDir()+"\\autoIt\\ImportFiles\\ItemBulkImportGST.xlsx";
				
			 	setCellData(xlfile, 0, 2, 0, "Table",1);
				setCellData(xlfile, 0, 2, 1, "Table",0);
				setCellData(xlfile, 0, 2, 2, "Raw Material",0);
				setCellData(xlfile, 0, 2, 3, "Weighted Avg",0);
				setCellData(xlfile, 0, 2, 4, "False",0);
				
				setCellData(xlfile, 0, 2, 23, "Table",0);
				setCellData(xlfile, 0, 2, 24, "10",0);
				setCellData(xlfile, 0, 2, 25, "11.22",0);
				
				 Thread.sleep(2000);
				  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
				   	toggleRibbon.click();
				   	
				   	Thread.sleep(2000);
				   	
				   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
				   	advanceMasterImportExportBtn.click();
				 	Thread.sleep(4000);
				 	Select s1=new Select(BulkImportOptionSelect);
				 	s1.selectByVisibleText("Yes");
				 	Thread.sleep(2000);
					
				 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
					FileUploadBtn.click();
					Thread.sleep(2000);
					
					 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\BulkItemImportWithExtraTab.exe");
					 Thread.sleep(7000);
					 
					 click(BulkBtn);
					 Thread.sleep(2000);
					 click(BulkImportBtn);
					 Thread.sleep(2000);
				 	
					 String expMessage1=null;
					 String actMessage1=null;
					 
					 try {
						 
						 
						 expMessage1="Records Imported Successfully";
						 actMessage1=checkValidationMessage(expMessage1);
						 
					 }
					 catch (Exception e) {
						 System.err.println("Server Error");
						 
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ServerErrorPopupCloseBtn));
						 ServerErrorPopupCloseBtn.click();
						 Thread.sleep(2000);
						 
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AdvMasterCloseBtn));
						 AdvMasterCloseBtn.click();
						 Thread.sleep(2000);
						 
						
					}
			
					 if(actMessage1.equalsIgnoreCase(expMessage1))
					 {
						 return true;
					 }
					 else
					 {
						 return false;
					 }

		 }
		 
		
		 
		 public boolean checkImportedDatainGSTTabAfterUpdatingExistingItem() throws InterruptedException
		 {
			Thread.sleep(5000);

				 int accNameListCount = accNameList.size();
				 
				 for(int i=0;i<accNameListCount;i++)		
						
					{
						String data = accNameList.get(i).getText();
						Thread.sleep(2000);
						if(data.equals("Table"))
						{
							Thread.sleep(2000);
							getAction().doubleClick(accChkBoxList.get(i)).build().perform();
							Thread.sleep(2000);
							break;
						}
					}
				 

				 Thread.sleep(2000);
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(GSTTab));
				 GSTTab.click();
				 Thread.sleep(3000);
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gst_ItemName));
				String actgst_ItemName=gst_ItemName.getAttribute("value");
				String expgst_ItemName="Table";
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gst_ItemCost));
				String actgst_ItemCost=gst_ItemCost.getAttribute("value");
				String expgst_ItemCost="10";
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gst_GSTRate));
				String actgst_GSTRate=gst_GSTRate.getAttribute("value");
				String expgst_GSTRate="11.22";	
				
			
				
				System.err.println("Item Text in GST Tab");
				System.err.println("********************************************************************************************************************************************************************************");
				System.err.println("Item Name		"		+		"	Actual		"	+	actgst_ItemName			+		"	Expected		"	+	expgst_ItemName);
				System.err.println("Item Cost"		+		"	Actual		"	+	actgst_ItemCost						+		"	Expected		"	+	expgst_ItemCost);
				System.err.println("GST Rate"		+		"	Actual		"	+	actgst_GSTRate						+		"	Expected		"	+	expgst_GSTRate);
				
				click(masterCloseBtn);
				Thread.sleep(2000);
				
				if(actgst_ItemName.equalsIgnoreCase(expgst_ItemName) && actgst_ItemCost.equalsIgnoreCase(expgst_ItemCost) && actgst_GSTRate.equalsIgnoreCase(expgst_GSTRate))
				{
					return true;
				}
				else
				{
					return false;
				}
		 }

		 
		  public boolean checkImportingItemAlongWithImageinItemMaster() throws InterruptedException, IOException, EncryptedDocumentException, InvalidFormatException
		  {
			  Thread.sleep(2000);
				 getDriver().navigate().refresh();
					Thread.sleep(4000);
					
				    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			      	homeMenu.click();
			      	Thread.sleep(2000);
			      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
			      	mastersMenu.click();
			      	Thread.sleep(2000);
			      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
			      	homeMasterItemMenu.click();
			      	Thread.sleep(2000);
			    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
			    	homeMasterItem_ItemMenu.click();
			    	
			  Thread.sleep(4000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toggleRibbon));
			   	toggleRibbon.click();
			   	
			   	Thread.sleep(5000);
			   	
			   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advanceMasterImportExportBtn));
			   	advanceMasterImportExportBtn.click();
			 	Thread.sleep(4000);
				/*
				 * Select s1=new Select(BulkImportOptionSelect); s1.selectByVisibleText("Yes");
				 * Thread.sleep(2000);
				 */
				
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(FileUploadBtn));
				FileUploadBtn.click();
				Thread.sleep(2000);
				
				 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\ItemImportWithImage.exe");
				 Thread.sleep(5000);
				 
				 
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterImportDataBtn));
					advMasterImportDataBtn.click();
					Thread.sleep(2000);
					
					String expMessage="1 Records Imported Successfully.";
					String actMessage=checkValidationMessage(expMessage);
					
					if(actMessage.equalsIgnoreCase(expMessage))
					{
						return true;
					}
					else
					{
						return false;
					}
		
		  }
		  
		  
		  @FindBy(xpath="//*[@id='pImage']")
		  public static WebElement imageTxt;
		  
		  @FindBy(xpath="//img[@id='viewimgpImage']")
		  public static WebElement viewImgBtn;
		  
		  public boolean checkImportedItemWithImage() throws InterruptedException
		  {
			  Thread.sleep(4000);
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		      	homeMenu.click();
		      	
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		      	mastersMenu.click();
		      	
		      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		      	homeMasterItemMenu.click();
		      	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
		    	homeMasterItem_ItemMenu.click();
		    	
		    	Thread.sleep(2000);
		    	
			  int accNameListCount = accNameList.size();
				 
				 for(int i=0;i<accNameListCount;i++)		
						
					{
						String data = accNameList.get(i).getText();
						if(data.equals("Plastic"))
						{
							Thread.sleep(2000);
							getAction().doubleClick(accChkBoxList.get(i)).build().perform();
							
							break;
						}
					}
				 Thread.sleep(2000);
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalTab1));
				 generalTab1.click();
				 Thread.sleep(2000);;
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
				String actItemName=nameTxt.getAttribute("value");
				String expItemName="Plastic";
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
				String actItemCode=codeTxt.getAttribute("value");
				String expItemCode="Plastic";
				
				getAction().moveToElement(viewImgBtn).build().perform();
				
				String actItemImage=imageTxt.getAttribute("value");
				String expItemImage="Plastic.jpg";
				
				
				System.out.println("Actual	"	+	actItemName			+	"Expected	"		+	expItemName);
				System.out.println("Actual	"	+	actItemCode			+	"Expected	"		+	expItemCode);
				System.out.println("Actual	"	+	actItemImage		+	"Expected	"		+	expItemImage);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewImgBtn));
				viewImgBtn.click();
				
				Thread.sleep(2000);
				ArrayList<String> newTabs = new ArrayList<String>(getDriver().getWindowHandles());	
				
				int actOpenWindowsCount = getDriver().getWindowHandles().size();
				int expOpenWindowsCount = 2;	
				
				System.out.println("Actual	"	+	actOpenWindowsCount			+	"Expected	"		+	expOpenWindowsCount);
				
				getDriver().switchTo().window(newTabs.get(1)).close();
			 	Thread.sleep(2000);
			 	
			 	
			 	getDriver().switchTo().window(newTabs.get(0));
			 	Thread.sleep(2000);
			 	
				/*
				 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
				 * masterCloseBtn)); masterCloseBtn.click(); Thread.sleep(2000);
				 */
				
			 	if(actOpenWindowsCount==expOpenWindowsCount && actItemImage.equalsIgnoreCase(actItemImage) && actItemName.equalsIgnoreCase(expItemName)
			 			&& actItemCode.equalsIgnoreCase(expItemCode))
			 	{
			 		return true;
			 	}
			 	else
			 	{
			 		return false;
			 	}
				
		  }
		  
		  @FindBy(xpath="(//*[@id='SaveImg'])[1]")
		  public static WebElement saveImgBtn;
		  
		  
		  public boolean checkSavingImageofItemImported() throws InterruptedException, AWTException, IOException
		  {
				
				  Thread.sleep(15000);
				  
				  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
				  saveImgBtn)); saveImgBtn.click();
				  
				  Thread.sleep(2000);
				  
				
				  
				  Robot robot = new Robot(); /*robot.keyPress(KeyEvent.VK_ENTER);
				  robot.keyRelease(KeyEvent.VK_ENTER);*/
				  
				  Thread.sleep(4000); robot.keyPress(KeyEvent.VK_CONTROL);
				  robot.keyPress(KeyEvent.VK_J); robot.keyRelease(KeyEvent.VK_J);
				  robot.keyRelease(KeyEvent.VK_CONTROL);
				  
				  
				  Thread.sleep(5000);
				  
				  robot.keyPress(KeyEvent.VK_TAB); robot.keyRelease(KeyEvent.VK_TAB);
				  robot.keyPress(KeyEvent.VK_TAB); robot.keyRelease(KeyEvent.VK_TAB);
				  
				  robot.keyPress(KeyEvent.VK_ENTER); robot.keyRelease(KeyEvent.VK_ENTER);
				  
				  Thread.sleep(3000);
				  File Efile1=new File(getBaseDir()+"\\autoIt\\ExportFiles\\Plastic.jpg");
				  
				  if(Efile1.exists()) { 
					  Efile1.delete(); 
					  System.out.println("File deleted"); 
					  }
				  
				  Thread.sleep(4000);
				  robot.keyPress(KeyEvent.VK_CONTROL); robot.keyPress(KeyEvent.VK_S);
				  robot.keyRelease(KeyEvent.VK_S); robot.keyRelease(KeyEvent.VK_CONTROL);
				  
				  Thread.sleep(2000);
				  
				  
				  Runtime.getRuntime().exec(getBaseDir() +"\\autoIt\\scripts\\ItemImageSave.exe");
				  
				  Thread.sleep(5000);
				  
				  robot.keyPress(KeyEvent.VK_ALT); robot.keyPress(KeyEvent.VK_F4);
				  
				  robot.keyRelease(KeyEvent.VK_ALT); robot.keyRelease(KeyEvent.VK_F4);
				  Thread.sleep(3000);
				  
				  ArrayList<String> newTabs = new
				  ArrayList<String>(getDriver().getWindowHandles());
				  
				  int actOpenWindowsCount = getDriver().getWindowHandles().size(); int
				  expOpenWindowsCount = 2;
				  
				  System.out.println("Actual	" + actOpenWindowsCount + "Expected	" +
				  expOpenWindowsCount);
				  
				  getDriver().switchTo().window(newTabs.get(1)).close(); Thread.sleep(2000);
				  
				  
				  getDriver().switchTo().window(newTabs.get(0)); Thread.sleep(2000);
				 
				 String actItemImage   = getBaseDir()+"\\autoIt\\ExportFiles\\Plastic.jpg";
				String  expItemImage   = getBaseDir()+"\\autoIt\\ImportFiles\\Plastic.jpg";
					
					double result=checkImageComparison(actItemImage,expItemImage);
					
					if(result==0.0)
					{
						return true;
					}
					else
					{
						return false;
					}
					
		  }			
					
					
		  @FindBy(xpath="//*[@title='Delete File']")
		  public static WebElement deleteImgBtn;

		  public boolean checkDeletingtheImageinImportedItem() throws InterruptedException
		  {
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteImgBtn));
			  deleteImgBtn.click();
			  Thread.sleep(2000);
			  
			  new WebDriverWait(getDriver(), 300).until(ExpectedConditions.alertIsPresent());
			  getAlert().accept();
			  Thread.sleep(2000);
			  
			  boolean actItemImage=imageTxt.getAttribute("value").isEmpty();
			System.out.println("Image is Empty"		+	actItemImage);
				
			if(actItemImage)
			{
				return true;
			}
			else
			{
				return false;
			}
		  }
								
		  
		  

	     public ItemsPage(WebDriver driver)
	     {
	    	PageFactory.initElements(driver, this);
	    	
	     }




		public static boolean checkReplenishmentTabDataAfterCloneItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
			Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(replenishmentTab1));
			 replenishmentTab1.click();
			 Thread.sleep(3000);
			 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectDefaultReplenidshmentTxt));
			Select s=new Select(itemSelectDefaultReplenidshmentTxt);
			String actDefaultReplenidshmentTxt=s.getFirstSelectedOption().getText();
			String expDefaultReplenidshmentTxt="Purchase";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectManufacturePolicyTxt));
			Select s1=new Select(itemSelectManufacturePolicyTxt);
			String actManufacturePolicyTxt=s1.getFirstSelectedOption().getText();
			String expManufacturePolicyTxt="Made to stock";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemPackingBOMTxt));
			String actPackingBOMTxt=Boolean.toString(itemPackingBOMTxt.getAttribute("value").isEmpty());
			String expPackingBOMTxt="true";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemBOMTxt));
			String actBOMTxt=Boolean.toString(itemBOMTxt.getAttribute("value").isEmpty());
			String expBOMtxt="true";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemSelectFlushingTxt));
			Select s2=new Select(itemSelectFlushingTxt);
			String actFlushingTxt=s2.getFirstSelectedOption().getText();
			String expFlushingTxt="Manual";
			
			System.err.println("Item Text in Replenishment Tab");
			System.err.println("********************************************************************************************************************************************************************************");
			System.err.println("Default Replenishment	"		+		"	Actual		"	+	actDefaultReplenidshmentTxt						+		"	Expected		"	+	expDefaultReplenidshmentTxt);
			System.err.println("Manufacture Policy		"		+		"	Actual		"	+	actManufacturePolicyTxt							+		"	Expected		"	+	expManufacturePolicyTxt);
			System.err.println("Packing BOM				"		+		"	Actual		"	+	actPackingBOMTxt								+		"	Expected		"	+	expPackingBOMTxt);
			System.err.println("BOM						"		+		"	Actual		"	+	actBOMTxt										+		"	Expected		"	+	expBOMtxt);
			System.err.println("Flushing				"		+		"	Actual		"	+	actFlushingTxt									+		"	Expected		"	+	expFlushingTxt);
			
			System.out.println("*********Saving Item*************");
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			 
			 String expMessage="Updated Successfully";
			 String actMessage=checkValidationMessage(expMessage);
			 
			 System.out.println("actMessage  : "+actMessage);
			 System.out.println("expMessage  : "+expMessage);
		
			if(actDefaultReplenidshmentTxt.equalsIgnoreCase(expDefaultReplenidshmentTxt) && actManufacturePolicyTxt.equalsIgnoreCase(expManufacturePolicyTxt)
					&& actPackingBOMTxt.equalsIgnoreCase(expPackingBOMTxt) && actBOMTxt.equalsIgnoreCase(expBOMtxt) && actFlushingTxt.equalsIgnoreCase(expFlushingTxt)
					&& actMessage.equalsIgnoreCase(expMessage))
			{
				
				
				return true;
			}
			else
			{
				return false;
			}
	
		}
	
}
