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
			 
			 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			 
			 excelReader.setCellData(xlfile, xlSheetName, 9, 7, actMessage);
			 
			 Thread.sleep(1000);
			 
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
			
			 System.out.println("********************************checkEditingInSavedItemINItemMaster********************8");
			 
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
			 
			 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			 
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
				
			 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			 
			 excelReader.setCellData(xlfile, xlSheetName, 15, 7, actMessage);
				
			 if(actMessage.equalsIgnoreCase(expMessage))
			 {
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
				 closeBtn.click();
				 excelReader.setCellData(xlfile, xlSheetName, 15, 8, resPass);
				 return true;
			 }
			 
			 else
			 {
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
				 closeBtn.click();
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
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			   	homeMenu.click();
			   	
			   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
			   	mastersMenu.click();
			   	
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
		      	homeMasterItemMenu.click();
		      	
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
			 		String  expItemList="[ITEMS GROUP, STOCK ITEM, BATCH WA ITEM, BATCH IGNORE EXP LIFO ITEM, BATCH FIFO ITEM, BATCH BR ITEM, BIN RAW MATERIAL ITEM, BIN FINISHED GOODS ITEM, BATCH BIN FINISHED GOODS ITEM, BATCH BIN WITH NO STOCK UPDATE ITEM, BATCH BIN WITH NO RESERVATION ITEM, RMA RITEM, RMA BATCH ITEM, WA COGS ITEM, FIFO COGS ITEM, BR COGS ITEM, STD RATE COGS ITEM]";
			 		
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
		      	
		  	if(advanceMasterLabel.getText().equalsIgnoreCase("Advance Master Import/Export----> Item"))
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
			
			 Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\ItemImportFromExcel.exe");
			 Thread.sleep(7000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterImportDataBtn));
			 advMasterImportDataBtn.click();
			 Thread.sleep(2000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(advMasterCloseBtn));
			 advMasterCloseBtn.click();
			 Thread.sleep(2000);
			 
			 
			 String expMessage1="20 Records Imported Successfully";
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

		public boolean checkItemListAfterImportingDataFromExcel() throws InterruptedException
		{
			Thread.sleep(2000);
			 	 
			
			
			 int itemCount = masterItemGridBodyName.size();
				

			 ArrayList<String> itemlist = new ArrayList<String>();

			 		for (int i = 0; i < itemCount; i++)
			 		{
			 			 String data = masterItemGridBodyName.get(i).getText();
			 			itemlist.add(data);
			 			
			 		}
			 		
			 		String actItemList=itemlist.toString();
			 		String  expItemList="[ITEMS GROUP, STOCK ITEM, BATCH WA ITEM, BATCH IGNORE EXP LIFO ITEM, BATCH FIFO ITEM, BATCH BR ITEM, BIN RAW MATERIAL ITEM, BIN FINISHED GOODS ITEM, BATCH BIN FINISHED GOODS ITEM, BATCH BIN WITH NO STOCK UPDATE ITEM, BATCH BIN WITH NO RESERVATION ITEM, RMA RITEM, RMA BATCH ITEM, WA COGS ITEM, FIFO COGS ITEM, BR COGS ITEM, STD RATE COGS ITEM, Item Import]";
			 		
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
			 int count = masterItemGridBodyName.size();
				
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
		    Thread.sleep(2000);
		    
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
		    masterAddGroupBtn.click();
			 
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




		@FindBy(xpath="//*[@id='ol_treeNavigation']/li[2]/span")
		public static WebElement GroupAccountLabel;

		public static boolean checkSavingItemSubGroupinItemMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
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
			
			    
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAddGroupBtn));
			    masterAddGroupBtn.click();
				 
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
			
	
		@FindBy(xpath="//*[@id='lblnchecked']/input")
		public static List<WebElement> accChkBoxList;
		
		
		public boolean checkSavingItemsinSubGroupsinItemMaster() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
		{
				
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
			
			    
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
			    masterNewBtn.click();
				 
				System.out.println("********************************checkSavingItemSubGroupItemMaster*****************************");
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
				nameTxt.click();
				nameTxt.clear();
				nameTxt.sendKeys("Item1");
				nameTxt.sendKeys(Keys.TAB);
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
				codeTxt.click();
				codeTxt.clear();
				codeTxt.sendKeys("Item1");
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
					if(data.equals("SubGroup2 Item"))
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
				nameTxt.sendKeys("Item2");
				nameTxt.sendKeys(Keys.TAB);
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
				codeTxt.click();
				codeTxt.clear();
				codeTxt.sendKeys("Item2");
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
			   	
			
			   	int accCount = accNameList.size();
			
			   	ArrayList<String> acclist = new ArrayList<String>();

			 		for (int i = 0; i < accCount; i++)
			 		{
			 			 String data = accNameList.get(i).getText();
			 			acclist.add(data);
			 			
			 		}
			 		
			 		String actAccList=acclist.toString();
			 		String  expAccList="[Group Item, ITEMS GROUP, STOCK ITEM, BATCH WA ITEM, BATCH IGNORE EXP LIFO ITEM, BATCH FIFO ITEM, BATCH BR ITEM, BIN RAW MATERIAL ITEM, BIN FINISHED GOODS ITEM, BATCH BIN FINISHED GOODS ITEM, BATCH BIN WITH NO STOCK UPDATE ITEM, BATCH BIN WITH NO RESERVATION ITEM, RMA RITEM, RMA BATCH ITEM, WA COGS ITEM, FIFO COGS ITEM, BR COGS ITEM, STD RATE COGS ITEM]";
			 		

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
			 	
			 	
			 	System.out.println("Actual Items List:  "+actAccList);
		 		System.out.println("Expected Items List:  "+expAccList);
			 	
		 		System.out.println("Actual Sub Group Items List:  "+actGrpAccList);
			 	System.out.println("Expected Sub Group Items List:  "+expGrpAccList);
			 	
				System.out.println("Actual  Items in Sub Group1 List:  "+actSubGrpAccList);
			 	System.out.println("Expected Items in Sub Group1 List:  "+expSubGrpAccList);
			 	
			 	System.out.println("Actual  Items in Sub Group2 List:  "+actSubGrpAcc2List);
			 	System.out.println("Expected Items in Sub Group2 List:  "+expSubGrpAcc2List);
			 	
			 	
			 	if(actAccList.equalsIgnoreCase(expAccList) && actGrpAccList.equalsIgnoreCase(expGrpAccList)
			 			&& actSubGrpAccList.equalsIgnoreCase(expSubGrpAccList) && actSubGrpAcc2List.equalsIgnoreCase(expSubGrpAcc2List))
			 	{getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
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


		
		
		

	     public ItemsPage(WebDriver driver)
	     {
	    	PageFactory.initElements(driver, this);
	    	
	     }
	
	
	
}
