package com.focus.Pages;

import java.awt.AWTException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Set;

import org.apache.bcel.classfile.Code;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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

public class MastersPage extends BaseEngine
{
	@FindBy (xpath="//*[@id='1']/div/span")
    public static WebElement homeMenu;
	
        //Masters Menu	
		@FindBy (xpath="//a[@id='1000']//span[contains(text(),'Masters')]")
	    public static WebElement mastersMenu;
		
		    //Accounts 
			@FindBy (xpath="//a[@id='1104']//span[contains(text(),'Account')]")
			public static WebElement accounts;
			
				//Accounts Title
				@FindBy (xpath="//span[@id='spnHeaderText']")
				public static WebElement accountsTitle;
		
				//Master Main Header Fields		
				@FindBy(xpath="//i[@class='icon-font6 icon-new']")
				public static WebElement masterNewBtn;
					
				@FindBy(xpath="//i[@class='icon-font6 icon-add-group']")
				public static WebElement masterAddGroupBtn; 
				 
				@FindBy(xpath="//i[@class='icon-font6 icon-edit']")
				public static WebElement masterEditBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-clone']")
				public static WebElement masterCloneBtn; 
		
				@FindBy(xpath="//i[@class='icon-properties icon-font6']")
				public static WebElement masterPropertiesBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-delete']")
				public static WebElement masterDeleteBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-group']")
				public static WebElement masterGroupBtn; 
				
				@FindBy(xpath="//*[@id='toggle_ribbon']")
				public static WebElement masterRibbonToExpandOptions; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-delete-all']")
				public static WebElement masterDeleteAllBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-closed-account']")
				public static WebElement masterCloseAccountOrItemOrUnitsOrDepartmentOrWareHouseOrStateOrCity; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-open-close-account']")
				public static WebElement masteropenCloseAccountOrItemOrDepartmentOrWareHouseOrState; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-import']")
				public static WebElement masterAdvanceMasterImportORExportBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-info']")
				public static WebElement masterAuthorInfoBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-up-arrow']")
				public static WebElement masterMoveUpBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-down-arrow']")
				public static WebElement masterMoveDownBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-sort']")
				public static WebElement masterSortBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-mass-update']")
				public static WebElement masterMassUpdateBtn; 
		
				@FindBy(xpath="//a[@class='lSNext']")
				public static WebElement masterRibbonControlNextBtn; 
				
				@FindBy(xpath="//*[@id='btnXMLImport']")
				public static WebElement masterImportFromXmlBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-xmlexport']")
				public static WebElement masterExportFormatToXmlBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-custamize']")
				public static WebElement masterCustamizemasterBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-options']")
				public static WebElement masterCustamizeViewBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-tree']")
				public static WebElement mastercustamizeTreeBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-stock-ledger']")
				public static WebElement masterLedgerBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-financial-1']")
				public static WebElement masterManageCreditBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-department-appropriation']")
				public static WebElement masterDepartmentAppropriationBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-budget']")
				public static WebElement masterBudgetBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-backtrack']")
				public static WebElement masterBackTrackBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-transfer']")
				public static WebElement masterTranferBtn; 
		
				@FindBy(xpath="//*[@id='btnNewSets']")
				public static WebElement masterNewSetBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-modifier']")
				public static WebElement masterModifierBtn; 
		
				@FindBy(xpath="//*[@id='btnSetType']")
				public static WebElement masterSetTypeBtn; 
		
				@FindBy(xpath="//i[@class='icon-convert icon-font6']")
				public static WebElement masterUnitConversionBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-alternate-product-1']")
				public static WebElement masterAlternateItemBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-related-product']")
				public static WebElement masterRelatedItemBtn; 
		
				@FindBy(xpath="//i[@class='icon-font6 icon-stock-ledger']")
				public static WebElement masterStockLedgerBtn; 
		
				
				
				@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[11]")
				private static List<WebElement> masterAccountsList;
	    
				//New General Elements
				//Name
				@FindBy(xpath="//input[@id='sName']")
				public static WebElement nameTxt;
				
				
				//Code
				@FindBy(xpath="//input[@id='sCode']")
				public static WebElement codeTxt;
				
			    //AccountType	
				@FindBy(xpath="//select[@id='iAccountType']")
				public static WebElement accountTypeDropdown;
				
				
				//CreditLimit
				@FindBy(xpath="//input[@id='fCreditLimit']")
				public static WebElement creditLimitTxt;

				//CreditDays
				@FindBy(xpath="//input[@id='iCreditDays']")
				public static WebElement creditdaysTxt;
				
				
				//ChequeDiscountLimit
				@FindBy(xpath="//input[@id='fChequeDiscountLimit']")
				public static WebElement chequeDiscountLimitTxt;
				
				
				//Rate Of Interest
				@FindBy(xpath="//input[@id='fRateofinterest']")
				public static WebElement rateofinterestTxt;
				
				
				//BankAccount
				@FindBy(xpath="//input[@id='iBankAc']")
				public static WebElement bankAccountTxt;
				
				@FindBy(xpath="//*[@id='iBankAc_input_image']/span")
				public static WebElement bankAccount_ExpansionBtn;
				
				@FindBy(xpath="//*[@id='iBankAc_input_settings']/span")
				public static WebElement bankAccount_SettingBtn;

	
				@FindBy(xpath="//*[@id='btnMasterSaveClick']")
				public static WebElement saveBtn;
				
				@FindBy(xpath="//i[@class='icon-font6 icon-close']")
				public static WebElement closeBtn;
	
				
				
				@FindBy(xpath="//*[@id='btnDelete']/i")
				public static WebElement deleteBtn;
				
				@FindBy(xpath="//div[@id='idGlobalError']")
				public static WebElement validationConfirmationMessage;

				@FindBy(xpath="//div[@class='theme_color font-6']")
				public static WebElement errorMessage;
				
				@FindBy(xpath="//span[@class='icon-reject2 theme_color']")
				public static WebElement errorMessageCloseBtn;
				
				
				@FindBy(xpath="//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[5]")
				public static WebElement masterCloseBtn;
				
				
				@FindBy(xpath="//*[@id='LandingGridBody']/tr[1]/td[11]")
				public static WebElement accountCreation;
				
				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[11]")
				public static WebElement accountNewCreationName;
				
				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[2]/td[11]")
				public static WebElement accountNewCreationNameOfLane;               
				
				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[8]/div[1]/label[1]/input[1]")
				public static WebElement accountFirstCheckBoxToSelection;

				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/ol[1]/li[5]/span[1]")
				public static WebElement accountGroupTitleDisplay;
				
				
				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/ol[1]/li[4]/span[1]")
				public static WebElement clickOnVendorGroup;
				
				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[2]/form[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/input[1]")
				public static WebElement getVendorGroupName;
				
				@FindBy(xpath="//button[@class='btn-link theme_icon-color btnGenerateMasterclose closeMasterPartial clsMasterSaveClose']//i[@class='icon-font6 icon-close']")
				public static WebElement closeAccountCreationScreen;
			
				@FindBy(xpath="//input[@id='chkRetainSelection']")
				public static WebElement accountMasterRetain;
				
				@FindBy(xpath="//input[@id='liSelectAllMasters']")
				public static WebElement accountMasterSelect;
				
				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[1]/td[8]/div[1]/label[1]/input[1]")
				public static WebElement accountFirstCheckBox;
				                
				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/table[1]/tbody[1]/tr[2]/td[8]/div[1]/label[1]/input[1]")
				public static WebElement accountSecondCheckBox;
				
				@FindBy(xpath="//button[@class='btn-link theme_icon-color btnGenerateMasterclose closeMasterPartial clsMasterSaveClose']")
				public static WebElement closeAccMasterCreationScreen;
				
				
				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/ul[1]/li[2]/button[5]/i[1]")
				public static WebElement closeMaster;
				
				
				@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[9]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]")
				public static WebElement getMsgOnDelete;
				
				@FindBy(xpath="//button[@id='btnOkForDelete']")
				public static WebElement clickOnOkInDelete;
				
				@FindBy(xpath="//button[@id='btnCancelForDelete']")
				public static WebElement clickOnCancelInDelete;
				
				
				@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[8]//input")
				private static List<WebElement> masterGridBodyChkbox;

			  @FindBy(xpath="//*[@id='LandingGridBody']/tr/td[11]")
			  private static List<WebElement> masterGridBodyName;
				
				
				
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
	
		
		
		
		// Restore
		
		@FindBy (xpath="//*[@id='25']/span")
	    private static WebElement dataMangementMenu;
		
		@FindBy (xpath="//*[@id='26']/span")
		private static WebElement backup;
			
		@FindBy (xpath="//*[@id='27']/span")
		private static WebElement restore;
		
		@FindBy (xpath="//button[@class='btn Fbutton']")
		private static WebElement folderpathExpandBtn;
		
		@FindBy (xpath="//div[@id='RestoreControls']//div[1]//div[1]")
		private static WebElement restoreCompanyBtn;
		
		@FindBy (xpath="//button[@class='Fbutton'][contains(text(),'Yes')]")
		private static WebElement overRideYesBtn;
		
		@FindBy (xpath="//button[@class='Fbutton'][contains(text(),'No')]")
		private static WebElement overRideNoBtn;
		
		@FindBy (xpath="//button[@class='Fbutton'][contains(text(),'Cancel')]")
		private static WebElement overRideCancelBtn;
		
		@FindBy(xpath="//span[@class='hidden-xs']")
		private static WebElement userNameDisplay;
		
		@FindBy(xpath="//*[@id='companyLogo']")
		private static WebElement companyLogo;
		
		@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
		private static WebElement companyName;
		
		@FindBy(xpath="//*[@id='userprofile']/li/span[2]")
		private static WebElement logoutOption;
		
		public static boolean checkRestoreCompanyForSanity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
		{
			LoginPage lp=new LoginPage(getDriver()); 
			
		    String unamelt="su";
		  
		    String pawslt="su";
		    
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
			dataMangementMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restore));
			restore.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(folderpathExpandBtn));
			folderpathExpandBtn.click();
			
			Thread.sleep(3000);
		     
		     Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\SanityRestore.exe");
		     
		     Thread.sleep(3000);
		     
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtn));
		     restoreCompanyBtn.click();
		     
		     try
		     {
		    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(overRideYesBtn));
		 		 overRideYesBtn.click();
		 		 Thread.sleep(20000);
		     }
		     catch(Exception e)
		     {
		    	 System.err.println("NO OLDER COMPANY EXISTS");
		     }
		     
		     if(getIsAlertPresent())
		     {
		    	 getWaitForAlert();
		    	 
		    	 getAlert().accept();
		     }
			
		     Thread.sleep(3000);
		     
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		     userNameDisplay.click();
			
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		     logoutOption.click();
		     
		     Thread.sleep(3000);
		     
		     lp.enterUserName(unamelt);
		     
		     lp.enterPassword(pawslt);
		     
		     Thread.sleep(2000);
		     
		     lp.clickOnSignInBtn();
		     
		     Thread.sleep(5000);
		     
		     String actUserInfo1=userNameDisplay.getText();
				
			System.out.println("User Info  : "+actUserInfo1);
			
			System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			companyLogo.click();
			
			String getCompanyTxt1=companyName.getText();
			String getLoginCompanyName1=getCompanyTxt1.substring(0, 19);
			System.out.println("company name  :  "+ getLoginCompanyName1);
			companyLogo.click();
			
			String expUserInfo1           ="SU";
			String expLoginCompanyName1   ="Automation Company ";
		     
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
		
		
	
	
	
	// Accounts
	
	public static boolean checkSavingAccountINAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	    homeMenu.click();
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	    mastersMenu.click();
	     
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accounts));
	    accounts.click();
	    	  
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountsTitle));
	    
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
	    masterNewBtn.click();
		 
		System.out.println("********************************checkSavingAccountINAccountMaster*****************************");
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		nameTxt.click();
		nameTxt.clear();
		nameTxt.sendKeys("Master Account");
		nameTxt.sendKeys(Keys.TAB);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		codeTxt.click();
		codeTxt.clear();
		codeTxt.sendKeys("MasterAccount");
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
	
	
	
	
	
	
	
	
	public static boolean checkEditingInSavedAccountINAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 System.out.println("********************************checkEditingInSavedAccountINAccountMaster********************");
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountFirstCheckBox));
			
		 int count = masterGridBodyName.size();
			
			for (int i = 0; i < count; i++) 
			{
				String data = masterGridBodyName.get(i).getText();
				
				if (data.equalsIgnoreCase("Master Account")) 
				{
					getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
					break;
				}
			}
		    
		    Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Master Account Update");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterAccountUpdate");
		 codeTxt.sendKeys(Keys.TAB);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountTypeDropdown));
		 accountTypeDropdown.sendKeys(Keys.TAB);
         
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Updated Successfully";
			
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
	
	
	
	
	
	
	
	public static boolean checkDeleteAccountInAccountMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountFirstCheckBox));
		 
		 int count = masterGridBodyName.size();
			
			for (int i = 0; i < count; i++) 
			{
				String data = masterGridBodyName.get(i).getText();
				
				if (data.equalsIgnoreCase("Master Account Update")) 
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
	
	
	
	
	@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[12]")
	private static List<WebElement> masterItemsList;
	
	@FindBy(xpath="//*[@id='221']/span")
	private static WebElement  homeMasterItemMenu;
	
	@FindBy(xpath="//*[@id='1105']/span")
	private static WebElement  homeMasterItem_ItemMenu;
	
	// Item
	
	public static boolean checkSavingItemInItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
      	homeMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
      	mastersMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
      	homeMasterItemMenu.click();
      	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_ItemMenu));
    	homeMasterItem_ItemMenu.click();
    	
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		 masterNewBtn.click();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Master Item");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterItem");
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
	
	
	 @FindBy(xpath="//*[@id='LandingGridBody']/tr/td[12]")
	  private static List<WebElement> masterGridBodyItemName;
	
	
	public static boolean checkEditingInSavedItemINItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 int count = masterGridBodyItemName.size();
			
		for (int i = 0; i < count; i++) 
		{
			String data = masterGridBodyItemName.get(i).getText();
			
			if (data.equalsIgnoreCase("Master Item")) 
			{
				getAction().doubleClick(masterGridBodyItemName.get(i)).build().perform();
				break;
			}
		}
		    
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Master Item Update");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterItemUpdate");
		 codeTxt.sendKeys(Keys.TAB);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Updated Successfully";
			
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
	
	
	
	@FindBy(xpath="//tbody[@id='LandingGridBody']/tr[1]/td[12]")
	public static WebElement itemNewCreationName;
	
	@FindBy(xpath="//tbody[@id='LandingGridBody']/tr[1]/td[10]")
	public static WebElement unitNewCreationName;
	
	

	public static boolean checkDeleteItemInItemMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 int count = masterGridBodyItemName.size();
			
			for (int i = 0; i < count; i++) 
			{
				String data = masterGridBodyItemName.get(i).getText();
				
				if (data.equalsIgnoreCase("Master Item Update")) 
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
         
         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
         clickOnOkInDelete.click();
         
         //Thread.sleep(1000);
         
         String expMessage="Record Deleted Successfully.";
 		
         String actMessage=checkValidationMessage(expMessage);
         
         System.out.println("********************************checkDeleteItemInItemMaster***************************");
			
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
	
	
	


	
	
	 @FindBy (xpath="//a[@id='1114']")
	 public static WebElement unitsMenu;
	 
	 @FindBy(xpath="//input[@id='iNoOfDecimals']")
		private static WebElement noOfDecimalsTxt;
		
		@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[10]")
		private static List<WebElement> masterItemUnitsList;
		
		@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[8]/div[1]/label/input")
		private static List<WebElement> masterItemUnitsChkboxList;
		
	
	// Units
	
	
	
	 public static boolean checkSavingUnitsInUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
		System.out.println("***********checkUnitMasterInUnitMenu Method Excutes **********");
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
     	homeMenu.click();
     	
     	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
     	mastersMenu.click();
     	
     	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
     	homeMasterItemMenu.click();
	     
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitsMenu));
	    unitsMenu.click();
	    
		 System.out.println("********************************checkSavingUnitsInUnitsMaster********************");
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		 masterNewBtn.click();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Master Units");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterUnits");
		 codeTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noOfDecimalsTxt));
		 noOfDecimalsTxt.click();
		 noOfDecimalsTxt.clear();
		 noOfDecimalsTxt.sendKeys("2");
		 noOfDecimalsTxt.sendKeys(Keys.TAB);
		 
		 Thread.sleep(1000);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Saved Successfully";
			
		 String actMessage=checkValidationMessage(expMessage);
		 
		 Thread.sleep(1000);
		 
		 System.out.println("Message     : "+actMessage+" Value Expected : "+expMessage);
			
		 if(actMessage.equalsIgnoreCase(expMessage))
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
	
	
	
	
	
	public static boolean checkEditingInSavedUnitsINUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 System.out.println("********************************checkEditingInSavedUnitsINUnitsMaster********************8");
		 
		 int count = masterItemUnitsList.size();
			
		 for (int i = 0; i < count; i++) 
		 {
			String itemUnit = masterItemUnitsList.get(i).getText();
			
			if (itemUnit.equalsIgnoreCase("Master Units")) 
			{
				getAction().doubleClick(masterItemUnitsList.get(i)).build().perform();
				break;
			}
		 }
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Master Units Update");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterUnitsUpdate");
		 codeTxt.sendKeys(Keys.TAB);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Updated Successfully";
			
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
	
	
	
	
	
	
	public static boolean checkDeleteUnitsInUnitsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		int count = masterItemUnitsList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = masterItemUnitsList.get(i).getText();
			
			if (data.equalsIgnoreCase("Master Units Update")) 
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
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
        clickOnOkInDelete.click();
        
        //Thread.sleep(1000);
        
        String expMessage="Record Deleted Successfully.";
		
        String actMessage=checkValidationMessage(expMessage);
        
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
	
	
	
	 @FindBy (xpath="//a[@id='1106']//span[contains(text(),'Department')]")
	 public static WebElement departmentMenu;
	
	
	
	// Department
	
	public static boolean checkSavingDepartmentInDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
      	homeMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
      	mastersMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentMenu));
	    departmentMenu.click();

		 System.out.println("***********checkSavingDepartmentInDepartmentMaster**********");
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		 masterNewBtn.click();
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Master Department");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterDepartment");
		 codeTxt.sendKeys(Keys.TAB);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Saved Successfully";
			
		 String actMessage=checkValidationMessage(expMessage);
		 
		 Thread.sleep(1000);
		 
		 if(actMessage.equalsIgnoreCase(expMessage))
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
	
	
	
	  @FindBy(xpath="//*[@id='LandingGridBody']/tr/td[10]")
	  private static List<WebElement> masterGridBodyDeptName;
	
	public static boolean checkEditingInSavedDepartmentINDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		int count = masterGridBodyDeptName.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = masterGridBodyDeptName.get(i).getText();
			
			if (data.equalsIgnoreCase("Master Department")) 
			{
				getAction().doubleClick(masterGridBodyDeptName.get(i)).build().perform();
				break;
			}
		}
	    
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Master Department Update");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterDepartmentUpdate");
		 codeTxt.sendKeys(Keys.TAB);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Updated Successfully";
			
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
	
	
	

	
	
	
	public static boolean checkDeleteInDepartmentMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		int count = masterGridBodyDeptName.size();
			
		for (int i = 0; i < count; i++) 
		{
			String data = masterGridBodyDeptName.get(i).getText();
			
			if (data.equalsIgnoreCase("Master Department Update")) 
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
         
         getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clickOnOkInDelete));
         clickOnOkInDelete.click();
         
         //Thread.sleep(1000);
         
         String expMessage="Record Deleted Successfully.";
 		
         String actMessage=checkValidationMessage(expMessage);
         
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
	
	
	
	
	 @FindBy (xpath="//a[@id='1115']//span[contains(text(),'Bins')]")
	 public static WebElement binsMenu;
	 
	 @FindBy(xpath="//input[@id='iColNo']")
	 public static WebElement binColNoTxt;
	
	// Bins
	
	 public static boolean checkSavingBinsInBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
      	homeMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
      	mastersMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binsMenu));
      	binsMenu.click();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		 masterNewBtn.click();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Master Bin");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterBin");
		 codeTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binColNoTxt));
		 binColNoTxt.click();
		 binColNoTxt.clear();
		 binColNoTxt.sendKeys("1");
		 binColNoTxt.sendKeys(Keys.TAB);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Saved Successfully";
			
		 String actMessage=checkValidationMessage(expMessage);
		 
		 Thread.sleep(1000);
		 
		 if(actMessage.equalsIgnoreCase(expMessage))
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
	
	
	
	
	public static boolean checkEditingInSavedBinsINBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 int count = masterItemUnitsList.size();
		
		 for (int i = 0; i < count; i++) 
		 {
			String actbins = masterItemUnitsList.get(i).getText();
			
			if (actbins.equalsIgnoreCase("Master Bin")) 
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
		 nameTxt.sendKeys("Master Bin Update");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterBinUpdate");
		 codeTxt.sendKeys(Keys.TAB);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Updated Successfully";
			
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
	
	
	

	public static boolean checkDeleteInBinsInBinsMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 int count = masterItemUnitsList.size();
			
		 for (int i = 0; i < count; i++) 
		 {
			String actbins = masterItemUnitsList.get(i).getText();
			
			if (actbins.equalsIgnoreCase("Master Bin Update")) 
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
         
         //Thread.sleep(1000);
         
         String expMessage="Record Deleted Successfully.";
 		
         String actMessage=checkValidationMessage(expMessage);
         
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
	
					 

	@FindBy (xpath="//a[@id='1107']")
	 public static WebElement warehouseMenu;
	
	
	// Warehouse
	
	
  	public  boolean checkSavingWarehouseInWarehouseMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
  	{
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
      	homeMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
      	mastersMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseMenu));
	    warehouseMenu.click();
	   
	    Thread.sleep(2000);
	    
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		 masterNewBtn.click();
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Master Warehouse");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterWarehouse");
		 codeTxt.sendKeys(Keys.TAB);
		 

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Saved Successfully";
			
		 String actMessage=checkValidationMessage(expMessage);
		 
		 Thread.sleep(1000);
			 
      	if(actMessage.equalsIgnoreCase(expMessage))
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
  	

	
	
  	
  	
  	
	public static boolean checkEditingInSavedWarehouseINWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 int count = masterItemUnitsList.size();
		
		 for (int i = 0; i < count; i++) 
		 {
			String actbins = masterItemUnitsList.get(i).getText();
			
			if (actbins.equalsIgnoreCase("Master Warehouse")) 
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
		 nameTxt.sendKeys("Master Warehouse Update");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("MasterWarehouseUpdate");
		 codeTxt.sendKeys(Keys.TAB);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 String expMessage="Updated Successfully";
			
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
	
	
	

	public static boolean checkDeleteInWarehouseInWarehouseMaster() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		 int count = masterItemUnitsList.size();
			
		 for (int i = 0; i < count; i++) 
		 {
			String actbins = masterItemUnitsList.get(i).getText();
			
			if (actbins.equalsIgnoreCase("Master Warehouse Update")) 
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
         
         //Thread.sleep(1000);
         
         String expMessage="Record Deleted Successfully.";
 		
         String actMessage=checkValidationMessage(expMessage);
         
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
	
	
	
	
	
	
	
	
	
	
	@FindBy(xpath="/html[1]/body[1]/section[1]/div[2]/div[1]/section[1]/div[1]/div[2]/fieldset[1]/div[1]/div[2]/div[2]/input[1]")
	private static WebElement coinsNameTxt;
	
	@FindBy(xpath="//span[@id='btnSave']")
	private static WebElement saveIcon;
	
	@FindBy(xpath="//a[@id='220']//span[contains(text(),'Currency')]")
	private static WebElement currencyMenu;
	
	@FindBy(xpath="//span[contains(text(),'Currency Master')]")
	private static WebElement currencyMasterMenu;
	
	@FindBy(xpath="//a[@id='71']//span[contains(text(),'Exchange Rate')]")
	private static WebElement exchangeRateMenu;
	
	@FindBy(xpath="//span[contains(text(),'Exchange Rate History')]")
	private static WebElement exchangeRateHistoryMenu;
	
	@FindBy(xpath="//span[@class='navText']//span[contains(text(),'Currency Master')]")
	private static WebElement currencyMasterlabel;
	
	@FindBy(xpath="//*[@id='ISOCurrencyCode']")
	private static WebElement ISOCurrencyCodeTxt;
	
	// Currency
	
	public static boolean checkUpdatingISOCurrencyCodeSLLCoinName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		     
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
		currencyMenu.click();
	         
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMasterMenu));
		currencyMasterMenu.click();
		    	  
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ISOCurrencyCodeTxt));
		ISOCurrencyCodeTxt.click();
		ISOCurrencyCodeTxt.sendKeys("SLL");
		
		Thread.sleep(2000);
		
		ISOCurrencyCodeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(coinsNameTxt));
		coinsNameTxt.click();
		coinsNameTxt.clear();
		coinsNameTxt.sendKeys("MasterSLL");
		
		getAction().moveToElement(coinsNameTxt).sendKeys(Keys.TAB).perform();	
		
		Thread.sleep(2000);
				    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
		saveIcon.click();
	  
		System.out.println("*********************checkUpdatingISOCurrencyCodeSLLCoinName******************************************");
		
		String expMessage="Data saved successfully";
		
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
	
	
	
	
	@FindBy(xpath="//div[@id='myNavbar']/ul/li[4]/div")
	private static WebElement currencyMasterCancelBtn;
	
	
	public static boolean checkUpdatingAgainISOCurrencyCodeSLLCoinName() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ISOCurrencyCodeTxt));
		ISOCurrencyCodeTxt.click();
		ISOCurrencyCodeTxt.sendKeys("SLL");
		
		Thread.sleep(2000);
		
		ISOCurrencyCodeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(coinsNameTxt));
		coinsNameTxt.click();
		coinsNameTxt.clear();
		coinsNameTxt.sendKeys("SLL");
		
		getAction().moveToElement(coinsNameTxt).sendKeys(Keys.TAB).perform();		
				    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
		saveIcon.click();
	  
		System.out.println("*********************checkUpdatingAgainISOCurrencyCodeSLLCoinName******************************************");
		
		String expMessage="Data saved successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMasterCancelBtn));
			currencyMasterCancelBtn.click();
			return true;
		}	 
		else
		{    
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMasterCancelBtn));
			currencyMasterCancelBtn.click();
			return false;
		}
     }
	
	
	
	
	
	
	
	
	
	@FindBy(xpath="//input[@id='id_option']")
	private static WebElement baseCurrencyTxt;
	
	@FindBy(xpath="//input[@id='EffectiveDate']")
	private static WebElement withEffectiveDateTxt;
	
	@FindBy(xpath="//span[@id='exchange_currency']")
	private static WebElement exchangeRateIcon;
	
	@FindBy(xpath="//*[@id='id_option_table_body']/tr/td[2]")
    private static List<WebElement> baseCurrencyListCount;
	
	@FindBy(xpath="//span[@id='DeleteData']")
	private static WebElement exchnageRateDeleteOption;
	
	@FindBy(xpath="//input[@id='ExchangeTable_DefinedAs']")
	private static WebElement gridDefineAsTxt;
	
	@FindBy(xpath="//input[@id='ExchangeTable_Rate']")
	private static WebElement gridRateTxt;
	
	@FindBy(xpath="//input[@id='ExchangeTable_CurrencyName']")
	private static WebElement gridcurrencyName1Txt;
	
	@FindBy(xpath="//i[@class='icon-font7 icon-save']")
	private static WebElement exchangeRateDefinitionSaveIcon;
	                    
	@FindBy(xpath="//span[@id='btnCurrencyMasterCancel']//i[@class='icon-font7 icon-close']")
	private static WebElement exchangeRateDefinitionPopCloseIcon;
	
	@FindBy(xpath="//span[@id='closeExchangeRate']")
	private static WebElement exchangeRateDefinitionPageCloseIcon;
	
	@FindBy(xpath="//i[@class='icon-font6 icon-close']")
	private static WebElement exchangeRateHistoryCloseIcon;
	
	@FindBy(xpath="//i[@class='icon-font7 icon-refresh']")
	private static WebElement loadIcon;
	
	@FindBy(xpath="//input[@id='id_option']")
	private static WebElement baseCurrencyInExchangeRateHistory;
	
	@FindBy(xpath="//tbody[@id='HistoryTableBody']/tr[2]/td[3]")
	private static WebElement valueUSD;
	
	@FindBy(xpath="//tbody[@id='HistoryTableBody']/tr[2]/td[4]")
	private static WebElement valueAED;
	
	@FindBy(xpath="//td[@id='ExchangeTable_col_1-2']")
	private static WebElement selCurrencyOne;
	
	@FindBy(xpath="//td[@id='ExchangeTable_col_2-2']")
	private static WebElement selCurrecnyTwo;
	
	@FindBy(xpath="//td[@id='ExchangeTable_col_1-3']")
	private static WebElement defCurrecnyOne;
	
	@FindBy(xpath="//td[@id='ExchangeTable_col_2-3']")
	private static WebElement defCurrecnyTwo;
	
	
	
	
	
	// Exchange Rate & Exchange History
	
	public static boolean checkDeleteOptionInExchangeRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	    homeMenu.click();
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	    mastersMenu.click();
	     
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
        currencyMenu.click();
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateMenu));
        exchangeRateMenu.click();
        
        Thread.sleep(2000);
        
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyTxt));
		baseCurrencyTxt.click();
		baseCurrencyTxt.sendKeys("I");
				
		int baseCurrencyCount=baseCurrencyListCount.size();
					
		System.err.println(baseCurrencyCount);
					
		for(int i=0 ; i < baseCurrencyCount ;i++)
		{
			String data=baseCurrencyListCount.get(i).getText();
						
			if(data.equalsIgnoreCase("INR"))
			{
				baseCurrencyListCount.get(i).click();
						
				break;
			}
		}	
						
		baseCurrencyTxt.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchnageRateDeleteOption));
		
		exchnageRateDeleteOption.click();
		
		 getWaitForAlert();
		
	     String actCompanySaveMsg=getAlert().getText();
	     
	     getAlert().accept();
	     
	     System.out.println("*******************************checkDeleteOptionInExchangeRate********************************");
	     
	     System.out.println("Company message is displaying as "+ actCompanySaveMsg);
	     
	     String expMessage="Exchange Rate deleted successfully.";
			
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
	
	
				
			
			
				
	public static boolean checkSavingCurrencyExchnageRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyTxt));
		baseCurrencyTxt.click();
		baseCurrencyTxt.sendKeys("I");
		
		int baseCurrencyCount=baseCurrencyListCount.size();
			
		System.err.println(baseCurrencyCount);
			
		for(int i=0 ; i < baseCurrencyCount ;i++)
		{
			String data=baseCurrencyListCount.get(i).getText();
				
			if(data.equalsIgnoreCase("INR"))
			{
				baseCurrencyListCount.get(i).click();
					
				break;
			}
		}	
				
		baseCurrencyTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(withEffectiveDateTxt));
		
		withEffectiveDateTxt.click();
		
		getAction().moveToElement(withEffectiveDateTxt).sendKeys(Keys.TAB).perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridcurrencyName1Txt));
		
		gridcurrencyName1Txt.sendKeys(Keys.TAB);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridDefineAsTxt));
			
		gridDefineAsTxt.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateTxt));
		gridRateTxt.sendKeys("100");
		gridRateTxt.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridcurrencyName1Txt));
		
		gridcurrencyName1Txt.sendKeys(Keys.TAB);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridDefineAsTxt));
			
		gridDefineAsTxt.sendKeys(Keys.TAB);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateTxt));
		gridRateTxt.sendKeys("50");
		gridRateTxt.sendKeys(Keys.TAB);
		
     	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionSaveIcon));
		exchangeRateDefinitionSaveIcon.click();
		
		System.out.println("***************checkSavingCurrencyExchnageRate***************");
		
		String expMessage="Exchange Rate updated successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
	     
		if(actMessage.equalsIgnoreCase(expMessage))
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionPageCloseIcon));
			exchangeRateDefinitionPageCloseIcon.click();
			
			return true;
		}	 
		else
		{    
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionPageCloseIcon));
			exchangeRateDefinitionPageCloseIcon.click();
			
			return false;
		}	     
	}
	
	
				
				
	
	
	
	
	
	
	public static boolean checkExchangeRateHistoryLoadButtonWithBaseCurrency() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	    homeMenu.click();
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	    mastersMenu.click();
	     
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
        currencyMenu.click();
         
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryMenu));
        exchangeRateHistoryMenu.click();
        
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loadIcon));
		loadIcon.click();
		
		String expMessage="Select the Base Currency";
		
		String actMessage=checkValidationMessage(expMessage);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		
		baseCurrencyInExchangeRateHistory.sendKeys("INR");
		Thread.sleep(2000);
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		System.out.println(defCurrecnyOne.getText());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueUSD));
		
		String actValueUSD                  = valueUSD.getText();
		String actValueAED                  = valueAED.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();
		
		String expValueUSD                  = "100.0000000000";
		String expValueAED                  = "50.0000000000";
		String expselCurrencyOne            = "USD";
		String expselCurrecnyTwo            = "AED";
		String expdefCurrecnyOne            = "INR";
		String expdefCurrecnyTwo            = "INR";
		
		System.out.println("********************************checkExchangeRateHistorySelectINR***********************************");
		
		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);
		
		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
			actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
			actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo))
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryCloseIcon));
			exchangeRateHistoryCloseIcon.click();
				return true;
		}	 
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryCloseIcon));
			exchangeRateHistoryCloseIcon.click();
			   return false;
		}
	}
	
	
				
				
				
				
				
	public static boolean checkUpdateExchangeRate() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	    homeMenu.click();
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	    mastersMenu.click();
	     
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
        currencyMenu.click();
         
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateMenu));
        exchangeRateMenu.click();
        
        Thread.sleep(2000);
        
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyTxt));
		baseCurrencyTxt.click();
		baseCurrencyTxt.sendKeys("I");
		
		int baseCurrencyCount=baseCurrencyListCount.size();
			
		System.err.println(baseCurrencyCount);
			
		for(int i=0 ; i < baseCurrencyCount ;i++)
		{
			String data=baseCurrencyListCount.get(i).getText();
				
			if(data.equalsIgnoreCase("INR"))
			{
				baseCurrencyListCount.get(i).click();
					
				break;
			}
		}	
				
		baseCurrencyTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(withEffectiveDateTxt));
		
		getAction().moveToElement(withEffectiveDateTxt).sendKeys(Keys.TAB).perform();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridcurrencyName1Txt));
		getAction().moveToElement(gridcurrencyName1Txt).sendKeys(Keys.TAB).perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridDefineAsTxt));
		getAction().moveToElement(gridDefineAsTxt).sendKeys(Keys.TAB).perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateTxt));
		gridRateTxt.sendKeys("70");
		
		getAction().moveToElement(gridRateTxt).sendKeys(Keys.TAB).perform();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridcurrencyName1Txt));
		getAction().moveToElement(gridcurrencyName1Txt).sendKeys(Keys.TAB).perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridDefineAsTxt));
		getAction().moveToElement(gridDefineAsTxt).sendKeys(Keys.TAB).perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridRateTxt));
		gridRateTxt.sendKeys("14");
		
		getAction().moveToElement(gridRateTxt).sendKeys(Keys.TAB).perform();
		
     	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionSaveIcon));
		exchangeRateDefinitionSaveIcon.click();
		
		System.out.println("*************************checkUpdateExchangeRate*************************");
		
		String expMessage="Exchange Rate updated successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateDefinitionPageCloseIcon));
 		exchangeRateDefinitionPageCloseIcon.click();
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{	
			 return true;
		}	 
		else
		{    
			 return false;
		}	     
	}
	
				
				
				
				
				
				
	public static boolean checkExchangeRateHistoryAfterUpdating() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	    homeMenu.click();
		
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	    mastersMenu.click();
	     
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(currencyMenu));
        currencyMenu.click();
         
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryMenu));
        exchangeRateHistoryMenu.click();
        
        Thread.sleep(2000);
        
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseCurrencyInExchangeRateHistory));
		baseCurrencyInExchangeRateHistory.sendKeys("INR");
		Thread.sleep(2000);
		baseCurrencyInExchangeRateHistory.sendKeys(Keys.TAB);
		
		System.out.println(defCurrecnyOne.getText());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueUSD));
		
		String actValueUSD                  = valueUSD.getText();
		String actValueAED                  = valueAED.getText();
		String actselCurrencyOne            = selCurrencyOne.getText();
		String actselCurrecnyTwo            = selCurrecnyTwo.getText();
		String actdefCurrecnyOne            = defCurrecnyOne.getText();
		String actdefCurrecnyTwo            = defCurrecnyTwo.getText();
		
		String expValueUSD                  = "70.0000000000";
		String expValueAED                  = "14.0000000000";
		String expselCurrencyOne            = "USD";
		String expselCurrecnyTwo            = "AED";
		String expdefCurrecnyOne            = "INR";
		String expdefCurrecnyTwo            = "INR";
		
		System.out.println("********************************checkExchangeRateHistoryAfterUpdating***********************************");
		
		System.out.println("valueUSD            "+actValueUSD       +"Value Excepted   :"+expValueUSD);
		System.out.println("ValueAED            "+actValueAED       +"Value Excepted   :"+expValueAED);
		System.out.println("selCurrencyOne      "+actselCurrencyOne +"Value Excepted   :"+expselCurrencyOne);
		System.out.println("selCurrecnyTwo      "+actselCurrecnyTwo +"Value Excepted   :"+expselCurrecnyTwo);
		System.out.println("defCurrecnyOne      "+actdefCurrecnyOne +"Value Excepted   :"+expdefCurrecnyOne);
		System.out.println("defCurrecnyTwo      "+actdefCurrecnyTwo +"Value Excepted   :"+expdefCurrecnyTwo);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exchangeRateHistoryCloseIcon));
		exchangeRateHistoryCloseIcon.click();
		
		if (actValueUSD.equalsIgnoreCase(expValueUSD)&& actValueAED.equalsIgnoreCase(expValueAED) && 
			actselCurrencyOne.equalsIgnoreCase(expselCurrencyOne) && actselCurrecnyTwo.equalsIgnoreCase(expselCurrecnyTwo) && 
			actdefCurrecnyOne.equalsIgnoreCase(expdefCurrecnyOne) && actdefCurrecnyTwo.equalsIgnoreCase(expdefCurrecnyTwo))
		 {	
				return true;
		 }	 
		 else
		 {
			 	return false;
		 }
	}
	
	
	
	
	
	@FindBy(xpath="//a[@id='72']/span")
	private static WebElement  homeMasterItem_UnitConversionMenu;
	
	@FindBy(xpath="//input[@id='optionBaseUnit']")
	public static WebElement unitConversionBaseUnitTxt;
	
	@FindBy(xpath="//input[@id='optionProduct']")
	public static WebElement unitConversionItemTxt;
	
	@FindBy(xpath="//td[@id='myTagsTable_col_1-2']")
	public static WebElement unitConversionTableRow1Col1;
	
	@FindBy(xpath="//td[@id='myTagsTable_col_1-3']")
	public static WebElement unitConversionTableRow1Col2;
	
	@FindBy(xpath="//td[@id='myTagsTable_col_1-4']")
	public static WebElement unitConversionTableRow1Col3;
	
	@FindBy(xpath="//td[@id='myTagsTable_col_1-5']")
	public static WebElement unitConversionTableRow1Col4;
	
	@FindBy(xpath="//td[@id='myTagsTable_col_1-6']")
	public static WebElement unitConversionTableRow1Col5;
	
	@FindBy(xpath="//input[@id='myTagsTable_UnitName']")
	public static WebElement unitConversion_EnterUnitNameTxt;
	
	@FindBy(xpath="//input[@id='myTagsTable_XFactor']")
	public static WebElement unitConversion_EnterXFactorTxt;
	
	@FindBy(xpath="//*[@id='ucCancel']/div")
	public static WebElement unitConversionCancelBtn;
	
	@FindBy(xpath="//*[@id='btnDeleteUnitConversion']/div")
	public static WebElement unitConversionDeleteBtn;
	
	
	// Unit Conversion
	
	
	public boolean checkDeletingUnitConversion() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	  	mastersMenu.click();
	  
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
      	homeMasterItemMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_UnitConversionMenu));
      	homeMasterItem_UnitConversionMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionBaseUnitTxt));
      	unitConversionBaseUnitTxt.sendKeys("Pcs");
      	Thread.sleep(2000);
      	unitConversionBaseUnitTxt.sendKeys(Keys.TAB);
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionItemTxt));
      	unitConversionItemTxt.sendKeys(Keys.TAB);
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterUnitNameTxt));
      	String actUnitName = unitConversion_EnterUnitNameTxt.getAttribute("value");
      	String expUnitName = "Dozs";
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col2));
      	String actXFactor = unitConversionTableRow1Col2.getText();
      	String expXFactor = "12.00000000";
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col5));
      	String actDescription = unitConversionTableRow1Col5.getText();
      	String expDescription = "1 Dozs = 12.00000000 Pcs";
      	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionDeleteBtn));
    	unitConversionDeleteBtn.click();
    	
    	getWaitForAlert();
    	getAlert().accept();
    	
    	String expMessage="Deleted Successfully.";
		
    	String actMessage=checkValidationMessage(expMessage);
    	
    	System.out.println("***************************************checkDeletingUnitConversion*************************************************");
    	
    	System.out.println("Unit Name   : "+actUnitName    +"  Value Expected  "+expUnitName);
    	System.out.println("XFactor     : "+actXFactor     +"  Value Expected  "+expXFactor);
    	System.out.println("Description : "+actDescription +"  Value Expected  "+expDescription);
    	System.out.println("Message     : "+actMessage     +"  Value Expected  "+expMessage);
  	
    	if(actUnitName.equalsIgnoreCase(expUnitName) && actXFactor.equalsIgnoreCase(expXFactor) && actDescription.equalsIgnoreCase(expDescription) 
    			&& actMessage.equalsIgnoreCase(expMessage))
    	{
    		return true;
    	}
    	else
    	{
    		return false;
    	}
	}

	
	
	
	public boolean checkSavingUnitConversion() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionBaseUnitTxt));
      	unitConversionBaseUnitTxt.sendKeys("Pcs");
      	Thread.sleep(2000);
      	unitConversionBaseUnitTxt.sendKeys(Keys.TAB);
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionItemTxt));
      	unitConversionItemTxt.sendKeys(Keys.TAB);
      	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col1));
    	unitConversionTableRow1Col1.click();
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterUnitNameTxt));
    	unitConversion_EnterUnitNameTxt.sendKeys("Dozs");
    	Thread.sleep(2000);
    	unitConversion_EnterUnitNameTxt.sendKeys(Keys.TAB);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterXFactorTxt));
    	unitConversion_EnterXFactorTxt.sendKeys("15");
    	unitConversion_EnterXFactorTxt.sendKeys(Keys.TAB);
	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
    	saveIcon.click();
    	
    	String expMessage="Successfully saved";
		
    	String actMessage=checkValidationMessage(expMessage);
    	
    	System.out.println("***************************************checkSavingUnitConversion*************************************************");
    	
    	System.out.println("Message     : "+actMessage     +"  Value Expected  "+expMessage);
      	
    	if(actMessage.equalsIgnoreCase(expMessage))
    	{
    		return true;
    	}
    	else
    	{
    		return false;
    	}
	}
	
	
	
	
	
	
	
	public boolean checkUpdatingUnitConversion() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionBaseUnitTxt));
      	unitConversionBaseUnitTxt.sendKeys("Pcs");
      	Thread.sleep(2000);
      	unitConversionBaseUnitTxt.sendKeys(Keys.TAB);
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionItemTxt));
      	unitConversionItemTxt.sendKeys(Keys.TAB);
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterUnitNameTxt));
      	String actUnitName = unitConversion_EnterUnitNameTxt.getAttribute("value");
      	String expUnitName = "Dozs";
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col2));
      	String actXFactor = unitConversionTableRow1Col2.getText();
      	String expXFactor = "15.00000000";
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col5));
      	String actDescription = unitConversionTableRow1Col5.getText();
      	String expDescription = "1 Dozs = 15.00000000 Pcs";
      	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col2));
    	unitConversionTableRow1Col2.click();
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterXFactorTxt));
    	unitConversion_EnterXFactorTxt.click();
    	unitConversion_EnterXFactorTxt.clear();
    	unitConversion_EnterXFactorTxt.sendKeys("12");
    	unitConversion_EnterXFactorTxt.sendKeys(Keys.TAB);
    	
    	Thread.sleep(1000);
	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveIcon));
    	saveIcon.click();
    	
    	String expMessage="Successfully saved";
		
    	String actMessage=checkValidationMessage(expMessage);
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionBaseUnitTxt));
      	unitConversionBaseUnitTxt.sendKeys("Pcs");
      	Thread.sleep(2000);
      	unitConversionBaseUnitTxt.sendKeys(Keys.TAB);
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionItemTxt));
      	unitConversionItemTxt.sendKeys(Keys.TAB);
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversion_EnterUnitNameTxt));
      	String actAfterUpdateUnitName = unitConversion_EnterUnitNameTxt.getAttribute("value");
      	String expAfterUpdateUnitName = "Dozs";
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col2));
      	String actAfterUpdateXFactor = unitConversionTableRow1Col2.getText();
      	String expAfterUpdateXFactor = "12.00000000";
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionTableRow1Col5));
      	String actAfterUpdateDescription = unitConversionTableRow1Col5.getText();
      	String expAfterUpdateDescription = "1 Dozs = 12.00000000 Pcs";
		
    	System.out.println("***************************************checkUpdatingUnitConversion*************************************************");
    	
    	System.out.println("Unit Name   : "+actUnitName    +"  Value Expected  "+expUnitName);
    	System.out.println("XFactor     : "+actXFactor     +"  Value Expected  "+expXFactor);
    	System.out.println("Description : "+actDescription +"  Value Expected  "+expDescription);
    	System.out.println("Message     : "+actMessage     +"  Value Expected  "+expMessage);
    	
    	System.out.println("Unit Name AfterUpdate  : "+actAfterUpdateUnitName    +"  Value Expected  "+expAfterUpdateUnitName);
    	System.out.println("XFactor AfterUpdate    : "+actAfterUpdateXFactor     +"  Value Expected  "+expAfterUpdateXFactor);
    	System.out.println("Description AfterUpdate: "+actAfterUpdateDescription +"  Value Expected  "+expAfterUpdateDescription);
  	
    	if(actUnitName.equalsIgnoreCase(expUnitName) && actXFactor.equalsIgnoreCase(expXFactor) && actDescription.equalsIgnoreCase(expDescription) 
    			&& actMessage.equalsIgnoreCase(expMessage)
    			&& actAfterUpdateUnitName.equalsIgnoreCase(expAfterUpdateUnitName) && actAfterUpdateXFactor.equalsIgnoreCase(expAfterUpdateXFactor) 
    			&& actAfterUpdateDescription.equalsIgnoreCase(expAfterUpdateDescription))
    	{
    		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionCancelBtn));
    		unitConversionCancelBtn.click();
    		return true;
    	}
    	else
    	{
    		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(unitConversionCancelBtn));
    		unitConversionCancelBtn.click();
    		return false;
    	}
	}
	
	
	
	
	
	
	
	
	@FindBy(xpath="//*[@id='80']/span")
	private static WebElement  homeMasterItem_buyerPriceBookMenu;
	
	@FindBy(xpath="//*[@id='76']/span")
	private static WebElement  homeMasterItem_sellerPriceBookMenu;
	
	@FindBy(xpath="//*[@id='ctrlOptionProPriceBookH']")
	private static WebElement buyerPriceBook_PriceBook;
	
	
	//First Row


	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[2]")
	private static WebElement itemFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[4]")
	private static WebElement wareHouseFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[5]")
	private static WebElement satrtingDateFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[6]")
	private static WebElement endingDateFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[7]")
	private static WebElement vendorFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[8]")
	private static WebElement currencyFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[9]")
	private static WebElement minQtyFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[10]")
	private static WebElement maxQtyFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[11]")
	private static WebElement unitFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[12]")
	private static WebElement rateFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[13]")
	private static WebElement PoFirstCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[1]/td[14]")
	private static WebElement PtFirstCellSelect; 	
			
	
	//Second Row
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[2]")
	private static WebElement itemSecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[4]")
	private static WebElement wareHouseSecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[5]")
	private static WebElement satrtingDateSecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[6]")
	private static WebElement endingDateSecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[7]")
	private static WebElement vendorSecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[8]")
	private static WebElement currencySecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[9]")
	private static WebElement minQtySecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[10]")
	private static WebElement maxQtySecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[11]")
	private static WebElement unitSecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[12]")
	private static WebElement rateSecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[13]")
	private static WebElement PoSecondCellSelect; 	
			
	@FindBy(xpath="//*[@id='PriceBookWebGrid_body']/tr[2]/td[14]")
	private static WebElement PtSecondCellSelect; 
	
	
	
	
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_2']")
	private static WebElement itemValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_4']")
	private static WebElement wareHouseValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_StartingDate']")
	private static WebElement startingDateValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_EndingDate']")
	private static WebElement endingDateValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_1']")
	private static WebElement vendorValue; 	
			
	@FindBy(xpath="//input[@id='Currency_option']")
	private static WebElement currencyOptionValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_minqty']")
	private static WebElement minQtyValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_maxqty']")
	private static WebElement maxQtyValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_option_11']")
	private static WebElement unitValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_val0']")
	private static WebElement rateValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_val1']")
	private static WebElement PoValue; 	
			
	@FindBy(xpath="//input[@id='PriceBookWebGrid_val2']")
	private static WebElement PtValue; 	
	
	@FindBy(xpath="//*[@id='chkDateRange']")
	private static WebElement buyerPriceBook_CustamizeDateRangeCheckBox;
	
	@FindBy(xpath="//*[@id='chkCustomer']")
	private static WebElement buyerPriceBook_CustomizeVendorCheckBox;
	
	@FindBy(xpath="//*[@id='chkDepartment']")
	private static WebElement buyerPriceBook_WarehouseCheckBox;

	@FindBy(xpath="//*[@id='chkQtyRange']")
	private static WebElement buyerPriceBook_QualityRangeCheckBox;
	
	@FindBy(xpath="//*[@id='chkCurrency']")
	private static WebElement buyerPriceBook_CurrencyCheckBox;
	
	@FindBy(xpath="//*[@id='chkUnit']")
	private static WebElement buyerPriceBook_UnitCheckBox;
	
	@FindBy(xpath="//*[@id='txtAbbreviation']")
	private static WebElement buyerPriceBook_AbbreviationBtn;
	
	@FindBy(xpath="//*[@id='PriceBookContainer']/div[2]/nav/div/div[1]/a/div/span[2]")
	private static WebElement buyerPriceBookLabel;
	
	@FindBy(xpath="//*[@id='btnClear']")
	private static WebElement buyerPriceBookClearBtn;
	
	@FindBy(xpath="//*[@id='btnPaste']")
	private static WebElement buyerPriceBookPasteBtn;
	
	@FindBy(xpath="//*[@id='btnCopy']")
	private static WebElement buyerPriceBookCopyBtn;
	
	@FindBy(xpath="//*[@id='btnSelectAll']")
	private static WebElement buyerPriceBookSelectAllBtn;
	
	@FindBy(xpath="//*[@id='btnClose']")
	private static WebElement buyerPriceBookCloseBtn;
	
	@FindBy(xpath="//*[@id='btnSave']")
	private static WebElement buyerPriceBookSaveBtn;
	
	@FindBy(xpath="//*[@id='ctrlOptionProPriceBookH_table_body']/tr/td[2]")
	private static List<WebElement> pricebookComboBoxList;
	
	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_2_table_body']/tr/td[2]")
	private static List<WebElement> itemComboBoxList;
	
	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_4_table_body']/tr/td[2]")
	private static List<WebElement> warehouseComboBoxList;
	
	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_1_table_body']/tr/td[2]")
	private static List<WebElement> vendorComboBoxList;
	
	@FindBy(xpath="//*[@id='Currency_option_table_body']/tr/td[2]")
	private static List<WebElement> currencyComboBoxList;
	
	@FindBy(xpath="//*[@id='PriceBookWebGrid_option_11_table_body']/tr/td[2]")
	private static List<WebElement> unitComboBoxList;
	
	@FindBy(xpath="//input[@id='btnLoad']")
	private static WebElement priceBookFilterAndLoadBtn;
	
	// Seller Price Book
	
	public boolean checkSavingSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	  	mastersMenu.click();
	  
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
      	homeMasterItemMenu.click();
      	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_sellerPriceBookMenu));
		homeMasterItem_sellerPriceBookMenu.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys("Master Seller PriceBook");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.SPACE);
		
		int count = itemComboBoxList.size();
		   
		for(int i=0;i<count;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("WA COGS ITEM"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("123");
		rateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingSellerPriceBook********************************");
		
		String expMessage="Price Book created successfully";
		
		String actMessage=checkValidationMessage(expMessage);
	
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
		   return true;
	   } 
	   else 
	   {
		   return false;
	   }
	}
	
	

	
	
	
	
	public boolean checkUpdatingSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys("Master Seller PriceBook");
		Thread.sleep(2000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateFirstCellSelect));
		rateFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.clear();
		rateValue.sendKeys("100");
		rateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkUpdatingSellerPriceBook********************************");
		
		String expMessage = "Price Book updated successfully";
		String actMessage=checkValidationMessage(expMessage);
		
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
		   return true;
	   } 
	   else 
	   {
		   return false;
	   }
	}
	
	
	
	
	
	
	
	public boolean checkDeletingSellerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys("Master Seller PriceBook");
		Thread.sleep(2000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
		deleteBtn.click();
		
		getWaitForAlert();
		getAlert().accept();
		
		System.out.println("*********************************checkDeletingSellerPriceBook********************************");
		
		String expMessage = "Pricebook details deleted successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
	
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
		   return true;
	   } 
	   else 
	   {
		   return false;
	   }
	}
					 
			 
	
	
	
	
	
	
	// Buyer Price Book
	
	public boolean checkSavingBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	  	homeMenu.click();
	  	
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
	  	mastersMenu.click();
	  
	  	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItemMenu));
      	homeMasterItemMenu.click();
      	
      	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMasterItem_buyerPriceBookMenu));
	  	homeMasterItem_buyerPriceBookMenu.click();
	  	
	  	Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys("Master Buyer PriceBook");
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemFirstCellSelect));
		itemFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemValue));
		itemValue.click();
		itemValue.sendKeys(Keys.SPACE);
		
		int count = itemComboBoxList.size();
		   
		for(int i=0;i<count;i++)
		{
			String data = itemComboBoxList.get(i).getText();
			
			if(data.equalsIgnoreCase("WA COGS ITEM"))
			{
				itemComboBoxList.get(i).click();
				break;
			}
		}
		   
		itemValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.sendKeys("123");
		rateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkSavingBuyerPriceBook********************************");
		
		String expMessage="Price Book created successfully";
		
		String actMessage=checkValidationMessage(expMessage);
	
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
		   return true;
	   } 
	   else 
	   {
		   return false;
	   }
	}
	
	

	
	
	
	
	public boolean checkUpdatingBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys("Master Buyer PriceBook");
		Thread.sleep(2000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateFirstCellSelect));
		rateFirstCellSelect.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rateValue));
		rateValue.clear();
		rateValue.sendKeys("100");
		rateValue.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBookSaveBtn));
		buyerPriceBookSaveBtn.click();
		
		System.out.println("*********************************checkUpdatingBuyerPriceBook********************************");
		
		String expMessage = "Price Book updated successfully";
		String actMessage=checkValidationMessage(expMessage);
		
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
		   return true;
	   } 
	   else 
	   {
		   return false;
	   }
	}
	
	
	
	
	
	
	
	public boolean checkDeletingBuyerPriceBook() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(buyerPriceBook_PriceBook));
		buyerPriceBook_PriceBook.sendKeys("Master Buyer PriceBook");
		Thread.sleep(2000);
		buyerPriceBook_PriceBook.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(priceBookFilterAndLoadBtn));
		priceBookFilterAndLoadBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
		deleteBtn.click();
		
		getWaitForAlert();
		getAlert().accept();
		
		System.out.println("*********************************checkDeletingBuyerPriceBook********************************");
		
		String expMessage = "Pricebook details deleted successfully.";
		
		String actMessage=checkValidationMessage(expMessage);
	
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
		   return true;
	   } 
	   else 
	   {
		   return false;
	   }
	}
	
				
	
	
	@FindBy(xpath="//div[contains(text(),'Settings')]")
	private static WebElement  settingsmenuBtn;
	
	@FindBy(xpath="//*[@id='68']/span")
	 private static WebElement  transactionAuthorizationBtn;
	 
	  @FindBy(xpath="//*[@id='id_menu_tree_Purchases Returns']/a/span")
	  private static WebElement purchasesReturns; 
	
	  @FindBy(xpath="//*[@id='id_menu_tree_Purchases Returns (1)']/a/span")
	  private static WebElement purchasesReturns1; 
	  
	  @FindBy(xpath="//*[@id='txtDoc']")
	  private static WebElement documentNameDropdown; 
	  
	  @FindBy(xpath="//*[@id='spnAddLevel']/i")
	  private static WebElement addLevelBtn; 
	  
	  
	  @FindBy(xpath="//*[@id='txtDescription']")
	  private static WebElement descriptionTxt; 
		
	  @FindBy(xpath="//*[@id='chkConditionNotRequired']")
	  private static WebElement conditionNotRequiredChkBox; 
	  
	  @FindBy(xpath="//*[@id='id_report_popup_body']/div[2]/div/div[1]/ul/li[2]/a/span")
	  private static WebElement userSelectionTab;
	  
	  
	  
	  @FindBy(xpath="//*[@id='id_report_popup_body']/div[2]/div/div[1]/ul/li[3]/a/span")
	  private static WebElement alertsTab;
	  
	  @FindBy(xpath="//u[contains(text(),'Email Header Template')]")
	  private static WebElement emailHeaderTemplateLink;
	  
	  @FindBy(xpath="//input[@id='smsTemplateOptCtrl']")
	  private static WebElement templateTxt;
	  
	  @FindBy(xpath="//tbody[@id='smsTemplateOptCtrl_table_body']/tr/td[2]")
	  private static List<WebElement> custTemplatesList;
	  
	  @FindBy(xpath="//span[@id='btnCloseSMSModal']")
	  private static WebElement custTemplateCloseBtn;
	  
	  
	  @FindBy(xpath="//*[@id='ddlUserSelection']")
	  private static WebElement userselectionDropdown;
	
	  @FindBy(xpath="//*[@id='ddlCriteria']")
	  private static WebElement criteriaDropdown;
	  
	  @FindBy(xpath="//*[@id='id_tblUserSelection_User']")
	  private static WebElement userDropdown;
		
	  @FindBy(xpath="//td[@id='id_tblUserSelection_col_1-1']")
	  private static WebElement userrow1;
		
	  @FindBy(xpath="//td[@id='id_tblUserSelection_col_2-1']")
	  private static WebElement userrow2;
	  
	  @FindBy(xpath="//*[@id='id_report_popup_ok']")
	  private static WebElement definitionOkBtn;					
				
	  @FindBy(xpath="//span[@id='id_report_popup_cancel']")
	  private static WebElement definitionCloseBtn;	
	  
	  
	  @FindBy(xpath="//*[@id='ulAuthBtns']/li[5]/span/i")
	  private static WebElement tranAuthCancelBtn; 
			
	  @FindBy(xpath="//*[@id='btnSaveAuthorization']")
	  private static WebElement tranAuthSaveBtn; 
	  
	  @FindBy(xpath="//*[@id='spndeleteProfile']")
	  private static WebElement tranAuthDeleteBtn; 
	
	  
	  @FindBy(xpath="//*[@id='ctrlAuthorizationH']")
	  private static WebElement transAuthNameTxt; 
	  
	  @FindBy(xpath="//div[@id='divLevel1']//i[@class='icon-edit icon-font7']")
	  private static WebElement Level1editBtn;
	  
	  
	  
	  @FindBy(xpath="//*[@id='68_0_AdvanceFilter_']/table/tbody/tr[1]/td[1]/select")
	  private static WebElement Row1ConjunctionDropdown; 
		
	  @FindBy(xpath="//*[@id='68_0_AdvanceFilter_']/table/tbody/tr[1]/td[2]/input")
	  private static WebElement Row1SelectFieldTxt; 
	
	  @FindBy(xpath="//*[@id='68_0_AdvanceFilter_']/table/tbody/tr[1]/td[3]/select")
	  private static WebElement Row1SelectOperatorDropdown; 
		
	  @FindBy(xpath="//*[@id='68_0_AdvanceFilter_']/table/tbody/tr[1]/td[4]/select")
	  private static WebElement Row1CompareWithDropdown; 
		
	  @FindBy(xpath="//*[@id='68_0_AdvanceFilter_']/table/tbody/tr[1]/td[5]/input")
	  private static WebElement Row1ValueTxt; 
	
	  @FindBy(xpath="//input[@id='advancefilter_master_68_0_']")
	  private static WebElement Row1VendorValueTxt; 
	
	@FindBy(xpath="//*[@id='advancefilter_master_68_0__table_body']/tr/td[2]")
	private static List<WebElement> Row1ValueList; 
	
	@FindBy(xpath="//input[@id='advancefilter_master_68_0_']")
	private static WebElement InputValueTxt; 
 
	@FindBy(xpath="//div[@id='filterTree_68_0_AdvanceFilter_']//a[@id='4']")
	private static WebElement selectFieldVendorAC;
	
	@FindBy(xpath="//*[@id='filterTree_68_0_AdvanceFilter_']/ul/li[1]/ul/li[1]/a")
	private static WebElement selectFieldVendorACName;
	
	
	@FindBy(xpath="//div[@id='filterTree_68_0_AdvanceFilter_']//a[@id='23']")
	private static WebElement selectFieldItem;
	
	@FindBy(xpath="//div[@id='filterTree_68_0_AdvanceFilter_']//a[@id='5021']")
	private static WebElement selectFieldItemName;
	
	@FindBy(xpath="//div[@id='filterTree_68_0_AdvanceFilter_']//a[@id='26']")
	private static WebElement selectFieldQuantity;
	
	
	
	
	// Transaction Authorization
	
	public boolean checkSavingTransactionAuthorizationInPurchaseReturns() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
    {
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
   		settingsmenuBtn.click();
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionAuthorizationBtn));
   		transactionAuthorizationBtn.click();
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNameDropdown));
   		documentNameDropdown.click();
   			
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesReturns));
   		purchasesReturns.click();
   	
   		Thread.sleep(3000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
   		transAuthNameTxt.click();
   		transAuthNameTxt.sendKeys("Returns");
   		transAuthNameTxt.sendKeys(Keys.TAB);
   			
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addLevelBtn));
   		addLevelBtn.click();
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
   		descriptionTxt.click();
   		descriptionTxt.sendKeys("PR");
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
   		conditionNotRequiredChkBox.click();
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
   		userSelectionTab.click();
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
   		userrow1.click();
   		Select user = new Select(userDropdown);
			user.selectByVisibleText("UserAllOptionsST");
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
   		alertsTab.click();
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateLink));
   		emailHeaderTemplateLink.click();
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(templateTxt));
   		templateTxt.click();
   		templateTxt.sendKeys(Keys.SPACE);
   		Thread.sleep(2000);
   		
   		ArrayList<String>  TemplatesList= new ArrayList<String>(); 
			
			int actCount = custTemplatesList.size();
			int expCount = 3; 

			System.err.println("Templates List : "+actCount+"  Value Expected  "+expCount);
			
			for(int i=0; i < actCount; i++)
			{
				String data = custTemplatesList.get(i).getText();
				TemplatesList.add(data);
			}
			
			templateTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
   		
			String actTemplateList = TemplatesList.toString();
			String expTemplateList = "[createBudgetEmailTemplate, createEmailTemplate, createEmailTemplate]";
			
			System.out.println("Actual   : "+actTemplateList);
			System.out.println("Expected : "+expTemplateList);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custTemplateCloseBtn));
			custTemplateCloseBtn.click();
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
   		definitionOkBtn.click();
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthSaveBtn));
   		tranAuthSaveBtn.click();
   		
   		System.out.println("*****************************************checkSavingTransactionAuthorizationInPurchaseReturns******************************************");
   		
   		String expValidationMsg="Authorization Flow saved successfully.";
   		String actValidationMsg=checkValidationMessage(expValidationMsg);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
			tranAuthCancelBtn.click();
   		
   		if(actCount==expCount && actTemplateList.equalsIgnoreCase(expTemplateList) && actValidationMsg.equalsIgnoreCase(expValidationMsg))
   		{
   			return true;
   		}
   		else
   		{
   			return false;
   		}
   	}
		
	    	
	    	
	    	
   	
	 
	 
   	
   	public boolean 	checkEditingTransactionAuthorizationInPurchaseReturns() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
   	{
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
   		settingsmenuBtn.click();
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionAuthorizationBtn));
   		transactionAuthorizationBtn.click();
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNameDropdown));
   		documentNameDropdown.click();
   			
           getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesReturns1));
           purchasesReturns1.click();
   		
   		Thread.sleep(2000);
   				
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
   		transAuthNameTxt.click();
   		transAuthNameTxt.sendKeys("Returns");
           Thread.sleep(3000);
   		transAuthNameTxt.sendKeys(Keys.TAB);
   			
   		Thread.sleep(3000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1editBtn));
   		Level1editBtn.click();
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
   		conditionNotRequiredChkBox.click();
   		
   		Thread.sleep(2000);
   		
   		Select s1=new Select(Row1ConjunctionDropdown);
   		
   		s1.selectByIndex(1);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Row1SelectFieldTxt));
   		Row1SelectFieldTxt.click();
   		
   		Thread.sleep(1000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldVendorAC));
   		selectFieldVendorAC.click();
   		
   		Thread.sleep(1000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldVendorACName));
   		selectFieldVendorACName.click();
   		
   		Thread.sleep(2000);
   		
   		Select s2=new Select(Row1SelectOperatorDropdown);
   		
   		s2.selectByIndex(1);
   		
   		Select s3=new Select(Row1CompareWithDropdown);
   		
   		s3.selectByIndex(1);
   		
   		///Row1CompareWithDropdown.sendKeys(Keys.TAB);
   		
   		Thread.sleep(2000);

   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Row1VendorValueTxt));
   		Row1VendorValueTxt.click();
   		Row1VendorValueTxt.sendKeys("V");
   		
   		Thread.sleep(2000);
   		
   		int count=Row1ValueList.size();
			
			for(int i=0 ; i < count ;i++)
			{
				String data=Row1ValueList.get(i).getText();
				
				System.err.println(data);
				
				if(data.equalsIgnoreCase("Vendor A"))
				{
					Row1ValueList.get(i).click();
						
					break;
				}
			}	
					
			Row1VendorValueTxt.sendKeys(Keys.TAB);
   		
   		
   		
   	/*	Thread.sleep(3000);
   		
   		Row1ValueTxt.sendKeys(Keys.TAB);*/
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
   		definitionOkBtn.click();
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthSaveBtn));
   		tranAuthSaveBtn.click();
   		
   		System.out.println("***************************************checkEditingTransactionAuthorizationInPurchaseReturns**************************************");
   		
   		String expValidationMsg="Authorization Flow saved successfully.";
   		
   		String actValidationMsg=checkValidationMessage(expValidationMsg);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
		tranAuthCancelBtn.click();
   		
   		if(actValidationMsg.equalsIgnoreCase(expValidationMsg))
   		{
   			return true;
   		}
   		else
   		{
   			return false;
   		}
   	}
	    	
	    	

   	
   	
   	
   	public boolean 	checkDeletingTransactionAuthorizationInPurchaseReturns() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
   	{
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
   		settingsmenuBtn.click();
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionAuthorizationBtn));
   		transactionAuthorizationBtn.click();
   		
   		Thread.sleep(2000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNameDropdown));
   		documentNameDropdown.click();
   			
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesReturns1));
   		purchasesReturns1.click();
   		
   		Thread.sleep(2000);
   				
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
   		transAuthNameTxt.click();
   		transAuthNameTxt.sendKeys("Returns");
           Thread.sleep(3000);
   		transAuthNameTxt.sendKeys(Keys.TAB);
   			
   		Thread.sleep(3000);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthDeleteBtn));
   		tranAuthDeleteBtn.click();
   		
   		getWaitForAlert();
   		getAlert().accept();
   		
   		System.out.println("***************************************checkEditingTransactionAuthorizationInPurchaseReturns**************************************");
   		
   		String expValidationMsg="Selected Authorization flow deleted successfully.";
   		
   		String actValidationMsg=checkValidationMessage(expValidationMsg);
   		
   		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
			tranAuthCancelBtn.click();
   		
   		if(actValidationMsg.equalsIgnoreCase(expValidationMsg))
   		{
   			return true;
   		}
   		else
   		{
   			return false;
   		}
   	}
	    
	
	
   	
   	
   	@FindBy(xpath="//*[@id='67']/span")
	private static WebElement  masterAuthorizationBtn;
	
	@FindBy(xpath="//input[@id='txtDoc']")
	private static WebElement  masterDropdown;
    	
	@FindBy(xpath="//li[contains(text(),'Warehouse')]")
	private static WebElement  masterAuthWarehouse;
	
	@FindBy(xpath="//ul[@id='MasterMenu']/li[4]")
	private static WebElement  masterAuthWarehouse1;
	
	@FindBy(xpath="//*[@id='67_0_AdvanceFilter_']/table/tbody/tr/td[1]/select")
	private static WebElement conjunctionDropdown;
    	
	@FindBy(xpath="//tbody//input[@placeholder='Select Field']")
	private static WebElement selectFieldTxt; 

    //Selection Field  Inner Fields
	@FindBy(xpath="//*[@id='filterTree_67_0_AdvanceFilter_']/ul/a[1]")
	private static WebElement sName;
	
	
	@FindBy(xpath="//*[@id='67_0_AdvanceFilter_']/table/tbody/tr/td[3]/select")
	private static WebElement selectOperatorDropdown; 
		
	@FindBy(xpath="//*[@id='67_0_AdvanceFilter_']/table/tbody/tr/td[4]/select")
	private static WebElement compareWithDropdown; 
		
	@FindBy(xpath="//*[@id='67_0_AdvanceFilter_']/table/tbody/tr/td[5]/input")
	private static WebElement valueTxt;
	
   	
   	// Master Authorization
   	
   	public boolean 	checkSavingMasterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
		masterAuthorizationBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropdown));
		masterDropdown.click();
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthWarehouse));
		masterAuthWarehouse.click();
	
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
		transAuthNameTxt.click();
		transAuthNameTxt.sendKeys("warehouse");
		transAuthNameTxt.sendKeys(Keys.TAB);
			
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addLevelBtn));
		addLevelBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
		descriptionTxt.click();
		descriptionTxt.sendKeys("WH");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
		userSelectionTab.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
		userrow1.click();
		Select user = new Select(userDropdown);
		user.selectByVisibleText("UserAllOptionsST");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
		alertsTab.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateLink));
		emailHeaderTemplateLink.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(templateTxt));
		templateTxt.click();
		templateTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		
		ArrayList<String>  TemplatesList= new ArrayList<String>(); 
		
		int actCount = custTemplatesList.size();
		int expCount = 3; 

		System.err.println("Templates List : "+actCount+"  Value Expected  "+expCount);
		
		for(int i=0; i < actCount; i++)
		{
			String data = custTemplatesList.get(i).getText();
			TemplatesList.add(data);
		}
		
		templateTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		String actTemplateList = TemplatesList.toString();
		String expTemplateList = "[createBudgetEmailTemplate, createEmailTemplate, createEmailTemplate]";
		
		System.out.println(actTemplateList);
		System.out.println(expTemplateList);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custTemplateCloseBtn));
		custTemplateCloseBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthSaveBtn));
		tranAuthSaveBtn.click();
		
		System.out.println("*****************************************checkSavingMasterAuthorization******************************************");
		
		String expValidationMsg="Authorization Flow saved successfully.";
		String actValidationMsg=checkValidationMessage(expValidationMsg);
		
		Thread.sleep(2000);
		
		if(actCount==expCount && actTemplateList.equalsIgnoreCase(expTemplateList) && actValidationMsg.equalsIgnoreCase(expValidationMsg))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
   	
   	
   	
   	
   	
   	
   	
   	
   	public boolean 	checkEditingMasterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropdown));
		masterDropdown.click();
			
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthWarehouse1));
        masterAuthWarehouse1.click();
		
		Thread.sleep(2000);
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
		transAuthNameTxt.click();
		transAuthNameTxt.sendKeys("warehouse");
        Thread.sleep(3000);
		transAuthNameTxt.sendKeys(Keys.TAB);
			
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1editBtn));
		Level1editBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
		conditionNotRequiredChkBox.click();
		
		Thread.sleep(2000);
		
		Select s1=new Select(conjunctionDropdown);
		
		s1.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
		selectFieldTxt.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sName));
		sName.click();
		
		Thread.sleep(2000);
		
		Select s2=new Select(selectOperatorDropdown);
		
		s2.selectByIndex(1);
		
		Select s3=new Select(compareWithDropdown);
		
		s3.selectByIndex(1);
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueTxt));
		valueTxt.click();
		valueTxt.sendKeys("TrialWarehouse");
		valueTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
		definitionOkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthSaveBtn));
		tranAuthSaveBtn.click();
		
		System.out.println("***************************************checkEditingMasterAuthorization**************************************");
		
		String expValidationMsg="Authorization Flow saved successfully.";
		
		String actValidationMsg=checkValidationMessage(expValidationMsg);
		
		Thread.sleep(2000);
		
		if(actValidationMsg.equalsIgnoreCase(expValidationMsg))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
   	
   	
   	
   	
   	
   	
   	
   	public boolean 	checkDeletingMasterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropdown));
		masterDropdown.click();
			
        getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthWarehouse1));
        masterAuthWarehouse1.click();
		
		Thread.sleep(2000);
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transAuthNameTxt));
		transAuthNameTxt.click();
		transAuthNameTxt.sendKeys("warehouse");
        Thread.sleep(3000);
		transAuthNameTxt.sendKeys(Keys.TAB);
			
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthDeleteBtn));
		tranAuthDeleteBtn.click();
		
		getWaitForAlert();
		getAlert().accept();
		
		System.out.println("***************************************checkDeletingMasterAuthorization**************************************");
		
		String expValidationMsg="Selected Authorization flow deleted successfully.";
		
		String actValidationMsg=checkValidationMessage(expValidationMsg);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranAuthCancelBtn));
		tranAuthCancelBtn.click();
		
		if(actValidationMsg.equalsIgnoreCase(expValidationMsg))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
   	
   	
   	
   	
	
	
   	
   	
   	
   	// Design Workflow
   	
   	
	@FindBy(xpath="//a[@id='74']//span[contains(text(),'Design Workflow')]")
	private static WebElement designWorkflowbtn; 
	
	@FindBy(xpath="//div[@class='navText']//span[contains(text(),'Design Workflow')]")
	private static WebElement designWorkflowLabel; 
	
	@FindBy(xpath="//span[@id='btnClear']//i[@class='icon-save icon-font6']")
	private static WebElement dwSaveBtn; 
	
	@FindBy(xpath="//i[@class='icon icon-clear icon-font6']")
	private static WebElement clearBtn; 
	
	@FindBy(xpath="//i[@class='icon-delete icon-font6']")
	private static WebElement dwDeleteBtn; 
	
	@FindBy(xpath="//i[@class='icon-close icon-font6']")
	private static WebElement dwCloseBtn;
	
	@FindBy(xpath="//input[@id='optWorkflow']")
	private static WebElement workflowNameTxt; 
	
	
	
   	@FindBy(xpath="//*[@id='optWorkflow_input_settings']/span")
	private static WebElement workFlowNameSettingsBtn; 
	
	@FindBy(xpath="//span[contains(text(),'Request For Quote')]")
	private static WebElement requestForQuoteDragAndDrop;
	
	@FindBy(xpath="//span[@class='DragLi ui-draggable ui-draggable-handle'][contains(text(),'Purchases Quotations')]")
	private static WebElement purchasesQuotationsDragAndDrop; 
	
	@FindBy(xpath="//div[@class='flowchart-operator-title ui-draggable-handle']")  //As it Is First One Selected Drag On FlowChart Is Same Id Property For Every Element
	private static WebElement FirstDraggedElement; 
	
	@FindBy(xpath="//div[contains(text(),'Purchases Quotations')]")
	private static WebElement draggedPurchasesQuotations; 
	
	@FindBy(xpath="/html[1]/body[1]/ul[1]/li[1]")
	private static WebElement draggedlinkBtn; 
	
	@FindBy(xpath="/html[1]/body[1]/ul[1]/li[2]")
	private static WebElement draggeddeleteBtn; 
	
	
	@FindBy(xpath="//*[@id='line0']")
	private static WebElement link0TO1; 
	
	@FindBy(xpath="//*[@id='line1']")
	private static WebElement link1TO2; 
	
	@FindBy(xpath="//*[@id='line2']")
	private static WebElement link2TO3; 
	
	@FindBy(xpath="//*[@id='line3']")
	private static WebElement link3TO4; 
	
	@FindBy(xpath="//*[@id='line4']")
	private static WebElement link4TO5; 
	
	@FindBy(xpath="//*[@id='line5']")
	private static WebElement link5TO6; 
	
	@FindBy(xpath="//*[@id='line6']")
	private static WebElement link6TO7; 
	
	@FindBy(xpath="//*[@id='line7']")
	private static WebElement link7TO8; 
	
	@FindBy(xpath="//*[@id='line8']")
	private static WebElement link8TO9; 
	
	
	
	
	@FindBy(xpath="//*[@id='idNode_0']/div[2]/div[1]/div/div/div[2]")
	private static WebElement linkNode0Left; 
	
	@FindBy(xpath="//*[@id='idNode_1']/div[2]/div[1]/div/div/div[2]")
	private static WebElement linkNode1Left; 
	
	@FindBy(xpath="//*[@id='idNode_2']/div[2]/div[1]/div/div/div[2]")
	private static WebElement linkNode2Left;
	
	@FindBy(xpath="//*[@id='idNode_3']/div[2]/div[1]/div/div/div[2]")
	private static WebElement linkNode3Left; 
	
	@FindBy(xpath="//*[@id='idNode_4']/div[2]/div[1]/div/div/div[2]")
	private static WebElement linkNode4Left; 
	
	@FindBy(xpath="//*[@id='idNode_5']/div[2]/div[1]/div/div/div[2]")
	private static WebElement linkNode5Left;
	
	@FindBy(xpath="//*[@id='idNode_6']/div[2]/div[1]/div/div/div[2]")
	private static WebElement linkNode6Left; 
	
	@FindBy(xpath="//*[@id='idNode_7']/div[2]/div[1]/div/div/div[2]")
	private static WebElement linkNode7Left; 
	
	@FindBy(xpath="//*[@id='idNode_8']/div[2]/div[1]/div/div/div[2]")
	private static WebElement linkNode8Left;
	
	
	
	
	
	@FindBy(xpath="//*[@id='idNode_0']/div[2]/div[2]/div/div/div[2]")
	private static WebElement linkNode0Right; 
	
	@FindBy(xpath="//*[@id='idNode_1']/div[2]/div[2]/div/div/div[2]")
	private static WebElement linkNode1Right; 
	
	@FindBy(xpath="//*[@id='idNode_2']/div[2]/div[2]/div/div/div[2]")
	private static WebElement linkNode2Right; 
	
	@FindBy(xpath="//*[@id='idNode_3']/div[2]/div[2]/div/div/div[2]")
	private static WebElement linkNode3Right; 
	
	@FindBy(xpath="//*[@id='idNode_4']/div[2]/div[2]/div/div/div[2]")
	private static WebElement linkNode4Right; 
	
	@FindBy(xpath="//*[@id='idNode_5']/div[2]/div[2]/div/div/div[2]")
	private static WebElement linkNode5Right; 
	
	@FindBy(xpath="//*[@id='idNode_6']/div[2]/div[2]/div/div/div[2]")
	private static WebElement linkNode6Right; 
	
	@FindBy(xpath="//*[@id='idNode_7']/div[2]/div[2]/div/div/div[2]")
	private static WebElement linkNode7Right; 
	
	@FindBy(xpath="//*[@id='idNode_8']/div[2]/div[2]/div/div/div[2]")
	private static WebElement linkNode8Right; 
	
	@FindBy(xpath="//div[@id='settingIcon']//span[@class='iconbar']")
	private static WebElement rightSideExpandBtn; 
	
	@FindBy(xpath="//a[contains(text(),'Definition')]")
	private static WebElement DefinitionTab; 
	
	@FindBy(xpath="//select[@id='DDLLinkValue']")
	private static WebElement DefinitionLinkValueDropdown; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[1]//li[1]//input[1]")
	private static WebElement dateChkBox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[2]//li[1]//input[1]")
	private static WebElement customerAccountChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[4]//li[1]//input[1]")
	private static WebElement salesOrderandInvoicewareHouseChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[3]//li[1]//input[1]")
	private static WebElement autosalesOrderandInvoicewareHouseChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[3]//li[1]//input[1]")
	private static WebElement narrationChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[4]//li[1]//input[1]")
	private static WebElement salesOrderandInvoicenarrationChkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[4]//li[1]//input[1]")
	private static WebElement itemchkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[3]//li[1]//input[1]")
	private static WebElement materialRequiitionAndPquotationsitemchkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[5]//li[1]//input[1]")
	private static WebElement salesOrderandInvoiceitemchkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[5]//li[1]//input[1]")
	private static WebElement purchaseOrderandPvoucheritemchkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[6]//li[1]//input[1]")
	private static WebElement purchaseVoucherAndPReturnsitemchkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[7]//li[1]//input[1]")
	private static WebElement purchaseVoucherNAndMRNitemchkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[6]//li[1]//input[1]")
	private static WebElement mrnAndPurchaseReturnsitemchkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[10]//li[1]//input[1]")
	private static WebElement salesInvoicenAndDeliveryNotesitemchkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[11]//li[1]//input[1]")
	private static WebElement salesInvoiceAndSalesRetutrnsitemchkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[6]//li[1]//input[1]")
	private static WebElement stockTransferAndSalesInvoiceNitemchkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[5]//li[1]//input[1]")
	private static WebElement unitsChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[6]//li[1]//input[1]")
	private static WebElement salesOrderandInvoiceunitsChkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[6]//li[1]//input[1]")
	private static WebElement quantityChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[7]//li[1]//input[1]")
	private static WebElement salesOrderandInvoicequantityChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[7]//li[1]//input[1]")
	private static WebElement voucherQtyhkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[8]//li[1]//input[1]")
	private static WebElement salesOrderandInvoicevoucherQtyhkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[8]//li[1]//input[1]")
	private static WebElement reserveChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[9]//li[1]//input[1]")
	private static WebElement salesOrderandInvoicereserveChkbox;
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[9]//li[1]//input[1]")
	private static WebElement rateChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[10]//li[1]//input[1]")
	private static WebElement salesOrderandInvoicerateChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[10]//li[1]//input[1]")
	private static WebElement grossChkbox; 
	
	@FindBy(xpath="//div[@id='LeftSideChkbox']//ul[11]//li[1]//input[1]")
	private static WebElement salesOrderandInvoicegrossChkbox; 
	
	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[1]//li[1]//input[1]")
	private static WebElement voucherGrossChkbox; 
	
	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[2]//li[1]//input[1]")
	private static WebElement netChkbox; 
	
	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[3]//li[1]//input[1]")
	private static WebElement voucherNetChkbox; 
	
	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[4]//li[1]//input[1]")
	private static WebElement userChkbox; 
	
	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[5]//li[1]//input[1]")
	private static WebElement roleChkbox; 
	
	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[6]//li[1]//input[1]")
	private static WebElement user1Chkbox; 
	
	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[7]//li[1]//input[1]")
	private static WebElement user2Chkbox; 
	
	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[8]//li[1]//input[1]")
	private static WebElement user3Chkbox; 
	
	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[9]//li[1]//input[1]")
	private static WebElement user4Chkbox; 
	
	@FindBy(xpath="//div[@id='rightsideCheckbox']//ul[10]//li[1]//input[1]")
	private static WebElement user5Chkbox; 
	
	@FindBy(xpath="//div[@class='col-lg-12']//button[@class='Fbutton'][contains(text(),'Ok')]")
	private static WebElement OkBtn; 
	
	@FindBy(xpath="//button[@class='Fbutton'][contains(text(),'Cancel')]")
	private static WebElement CancelBtn; 
	
	
	
	
   	
   	public boolean checkSavingTestWorkFlow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
	    settingsmenuBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(designWorkflowbtn));
		designWorkflowbtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workflowNameTxt));
		workflowNameTxt.click();
		workflowNameTxt.sendKeys("TestWorkFlow");
		workflowNameTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(requestForQuoteDragAndDrop));
		requestForQuoteDragAndDrop.click();
		
		getAction().dragAndDropBy(requestForQuoteDragAndDrop, 350, -250).build().perform();
	
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchasesQuotationsDragAndDrop));
		purchasesQuotationsDragAndDrop.click();
		
		getAction().dragAndDrop(purchasesQuotationsDragAndDrop, FirstDraggedElement).build().perform();
		
		getAction().dragAndDropBy(draggedPurchasesQuotations, 350, 0).build().perform();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(draggedPurchasesQuotations));
		
		getAction().contextClick(FirstDraggedElement).build().perform();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(draggedlinkBtn));
		draggedlinkBtn.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(linkNode1Left));
		linkNode1Left.click();
		
		Thread.sleep(2000);
		
		FirstDraggedElement.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(link0TO1));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(link0TO1));
		link0TO1.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DefinitionLinkValueDropdown));
		Select s=new Select(DefinitionLinkValueDropdown);
		s.selectByVisibleText("Quantity");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(materialRequiitionAndPquotationsitemchkbox));
		materialRequiitionAndPquotationsitemchkbox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBtn));
		OkBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwSaveBtn));
		dwSaveBtn.click();
		
		System.out.println("********************************checkSavingTestWorkFlow***********************************");
		
		String expMessage="Workflow saved successfully";
		
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
	
	
	
	
		
		
	public boolean checkEditingSavedTestWorkFlow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workflowNameTxt));
		workflowNameTxt.click();
		workflowNameTxt.sendKeys("TestWorkFlow");
		Thread.sleep(2000);
		workflowNameTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(link0TO1));
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(link0TO1));
		link0TO1.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DefinitionLinkValueDropdown));
		Select s=new Select(DefinitionLinkValueDropdown);
		s.selectByVisibleText("Quantity");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountChkbox));
		customerAccountChkbox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(OkBtn));
		OkBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwSaveBtn));
		dwSaveBtn.click();
		
		System.out.println("*******************************checkEditingSavedTestWorkFlow**************************");
		
		String expMessage="Workflow saved successfully";
		
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


		
		
		
	
	public boolean checkDeletingTestWorkFlow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(workflowNameTxt));
		workflowNameTxt.click();
		workflowNameTxt.sendKeys("TestWorkFlow");
		Thread.sleep(2000);
		workflowNameTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dwDeleteBtn));
		dwDeleteBtn.click();

		Thread.sleep(1000);
		
		getWaitForAlert();
		getAlert().accept();
		
		System.out.println("****************************checkDeletingTestWorkFlow***************************");
		
		String expMessage="Workflow Deleted Successfully !!";
		
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
   	
   	
   	
   	
   	
   	
	
	
	
	
	// Security
	
	
	@FindBy (xpath="//*[@id='6']/span")
    private static WebElement securityMenu;
	
	//Home Menu Network Policy Menu
	@FindBy (xpath="//*[@id='7']/span")
	private static WebElement networkPolicy;
		
			@FindBy (xpath="//*[@id='securityNavBar']/div/div[1]/a/div/span[2]")
			private static WebElement networkPolicyTitle;
			
			@FindBy (xpath="//*[@id='allowDiv']/div[1]/label")
			private static WebElement allowRadioBtnInNetworkPolicy;
		
			@FindBy (xpath="//*[@id='allowDiv']/div[2]/label")
			private static WebElement blockRadioBtnInNetworkPolicy;
			
			@FindBy (xpath="//*[@id='IPtable']/thead/tr/td[1]")
			private static WebElement startIp;
			
			@FindBy (xpath="//*[@id='IPtable']/thead/tr/td[2]")
			private static WebElement endIp;
				
			@FindBy (xpath="//*[@id='0']/td[1]")
			private static WebElement startIpValue;
			
			@FindBy (xpath="//*[@id='0']/td[2]")
			private static WebElement endIpValue;
			
			@FindBy (xpath="//*[@id='0']/td[2]/span[1]")
			private static WebElement closeImage;
			
			@FindBy (xpath="//*[@id='From']")
			private static WebElement fromIp;
			
			@FindBy (xpath="//*[@id='To']")
			private static WebElement toIp;
			
			@FindBy (xpath="//i[@class='icon-font6 icon-add']")
			private static WebElement addIp;
			
			@FindBy (xpath="//i[@class='icon-font6 icon-clear']")
			private static WebElement clearIp;
			
			@FindBy (xpath="//*[@id='buttons']/span[1]/i")
			private static WebElement saveBtnInNetworkPolicy;
			
			@FindBy (xpath="//*[@id='addbtn']/i")
			private static WebElement updateBtnInNetworkPolicy;
			
			@FindBy (xpath="//*[@id='buttons']/span[2]/i")
			private static WebElement closeBtnInNetworkPolicy;

		@FindBy (xpath="//*[@id='8']/span")
		private static WebElement passwordPolicy;
		
		@FindBy (xpath="//li[@class='treeview theme_menu_color']//span[contains(text(),'Create Profile')]")
		private static WebElement createProfile;
		
		@FindBy (xpath="//*[@id='10']/span")
		private static WebElement createRole;
		
		@FindBy (xpath="//*[@id='11']/span")
		private static WebElement createUser;
		
		@FindBy (xpath="//*[@id='12']/span")
		private static WebElement changePassword;
		
		@FindBy (xpath="//*[@id='13']/span")
		private static WebElement updatePersonalInfo;
		
		@FindBy (xpath="//*[@id='14']/span")
		private static WebElement usageLog;
		
		@FindBy (xpath="//*[@id='644']/span")
		private static WebElement userRightsReport;
		
		@FindBy (xpath="//*[@id='652']/span")
		private static WebElement roleRightsReport;
		
		@FindBy (xpath="//*[@id='653']/span")
		private static WebElement profileRightsReport;
		
		@FindBy(xpath="//*[@id='allow']")
		private static WebElement allowRadioBtn;
		
		@FindBy(xpath="//div[@id='idGlobalError']")
		private static WebElement validationMessage;
		
		
		//Password Policy Menu
		@FindBy (xpath="//*[@id='8']/span")
	    private static WebElement passwordPolicySubMenu;
		
		//Password Policy Screen Title
		@FindBy (xpath="//*[@id='securityNavBar']/div/div[1]/a/div/span[2]")
	    private static WebElement passwordPolicyTitle;
		
		//Password Policy List
		@FindBy (xpath="//*[@id='PolicyListDiv']")
		private static WebElement passwordPolicyList;

		@FindBy (xpath="//*[@id='addPasswordPolicybtn']/i")
		private static WebElement passwordPolicyAddBtn;
		
		//Password Policy DeleteBtn
		@FindBy (xpath="//*[@id='controlBtns']/span[2]/i")
		private static WebElement passwordPolicyDeletebtn;
		
		//Password Policy CancelBtn
		@FindBy (xpath="//*[@id='controlBtns']/span[3]/i")
		private static WebElement passwordPolicyCancelBtn;
		
		//Password Policy SaveBtn
		@FindBy (xpath="//*[@id='btnSaveId']/i")
		private static WebElement passwordPolicySaveBtn;
		
		//Password Policy SaveBtnone
		@FindBy (xpath="//span[@id='btnSaveId']")
		private static WebElement pPSaveBtn;
				
	    //Password Policy CloseBtn
	  	@FindBy (xpath="//*[@id='myNavbar']/ul/li/span[2]/i")
	  	private static WebElement passwordPolicyCloseBtn;
		
	    //Password Policy policyName
	  	@FindBy (xpath="//*[@id='policyName']")
	  	private static WebElement policyName;

	    //Password Policy passwordLength
	  	@FindBy (xpath="//*[@id='MinPassLength']")
	  	private static WebElement passwordLength;
	   	
	    //Password Policy passwordComplexity
	  	@FindBy (xpath="//*[@id='complexity']")
	  	private static WebElement passwordComplexity;
	  	
	    //Password Policy doNotAllowPreviousPassword
	  	@FindBy (xpath="//*[@id='Donotallprevious']")
	  	private static WebElement doNotAllowPrevious;
	  	
	    //Password Policy passwordExpiryDays
	  	@FindBy (xpath="//*[@id='passwordexpirydays']")
	  	private static WebElement passwordExpiryDays;
	  	
	    //Password Policy passwordOTPExpiryInMins
	  	@FindBy (xpath="//*[@id='OTPexpiryinminute']")
	  	private static WebElement passwordExpiryInMins;
	  	
	    //Password Policy passwordInvalidAttempts
	  	@FindBy (xpath="//*[@id='NoOfInvalidAttemps']")
	  	private static WebElement passwordInvalidAttempts;
	  	
	    //Password Policy passwordLockOut
	  	@FindBy (xpath="//*[@id='Lockoutperiod']")
	  	private static WebElement passwordLockOut;
	  	
	    //Password Policy passwordLockOutUnits
	  	@FindBy (xpath="//*[@id='Units']")
	  	private static WebElement passwordLockOutUnits;
	  	
	    //Password Policy sendMailCheck
	  	@FindBy (xpath="//*[@id='Sendemailchk']")
	  	private static WebElement sendMailCheck;
	  	
	    //Password Policy sendMailOnSuccess
	  	@FindBy (xpath="//*[@id='sendemailonloginsuccesschk']")
	  	private static WebElement sendMailOnSuccess;
	  	
	  	
	    //Password Policy sendMailOnFailure
	  	@FindBy (xpath="//*[@id='Sendemailonloginfailurechk']")
	  	private static WebElement sendMailOnFailure;
	  	
	  	//Password Policy sendMailOnFailure
	  	@FindBy (xpath="//*[@id='PolicyDiv']/div[14]/div/label")
	  	private static WebElement changePasswordAfterFirstLogin;
	  	
	  	//Password Policy sendMailOnFailure
	  	@FindBy (xpath="//*[@id='PolicyDiv']/div[15]/div/label")
	  	private static WebElement cannotChangePassword;
	  	
	  	//Password Policy sendMailOnFailure
	  	@FindBy (xpath="//*[@id='PolicyDiv']/div[16]/div/label")
	  	private static WebElement otpBasedLogin;
	  	
	  	//Password Policy sendMailOnFailure
	  	@FindBy (xpath="//*[@id='sendEmailId']")
	  	private static WebElement sendEmailId;
	  	
	  	//Password Policy sendMailOnFailure
	  	@FindBy (xpath="//*[@id='sendemailonloginsuccessId']")
	  	private static WebElement sendEmailOnLoginSuccess;
	  	
	  	//Password Policy sendMailOnFailure
	  	@FindBy (xpath="//*[@id='SendemailonloginfailureId']")
	  	private static WebElement sendEmailOnLoginFailure;
	  	
		

    //Web Elements of CraeteUser Screen
	@FindBy(xpath="//span[contains(text(),'Create user')]")
	private static WebElement createUserLabel;

	@FindBy(xpath="//span[contains(text(),'Additional Info')]")
	private static WebElement createUserAdditionalInfoTab;

	@FindBy(xpath="//ul[@class='nav nav-tabs theme_background-color theme_button_color']//span[contains(text(),'Restrictions')]")
	private static WebElement createUserRestrictionsTab;

	@FindBy(xpath="//span[contains(text(),'Device')]")
	private static WebElement createUserDeviceTab;

	@FindBy(xpath="//span[contains(text(),'Restriction for entry')]")
	private static WebElement createUserRestrictionforentryTab;

	@FindBy(xpath="//span[contains(text(),'User Information')]")
	private static WebElement createUserUserInformationTab;

	@FindBy(xpath="//i[@class='icon-font6 icon-create-user']")
	private static WebElement createUserIcon;

	@FindBy(xpath="//i[@class='icon-font6 icon-group']")
	private static WebElement createUserCreateGroupIcon;

	@FindBy(xpath="//i[@class='icon-font6 icon-cancel']")
	private static WebElement createUserCancelIcon;

	@FindBy(xpath="//input[@id='loginName']")
	private static WebElement createUserLoginNameTxt;

	@FindBy(xpath="//select[@id='ERPRoles']")
	private static WebElement createUserERPRoleDropDown;

	@FindBy(xpath="//select[@id='passwordPolicy']")
	private static WebElement createUserPasswordPolicyDropdown;

	@FindBy(xpath="//input[@id='password']")
	private static WebElement createUserPasswordTxt;

	@FindBy(xpath="//input[@id='confirmpassword']")
	private static WebElement createUserConfirmPasswordTxt;

	@FindBy(xpath="//input[@id='Username']")
	private static WebElement createUserNameTxt;

	@FindBy(xpath="//input[@id='lginAbbr']")
	private static WebElement createUserLoginAbbrivationTxt;

	@FindBy(xpath="//select[@id='SecurityQues']")
	private static WebElement createUserSecurityQuestionDropDown;

	@FindBy(xpath="//input[@id='SecurityAns']")
	private static WebElement createUserSecurityAnswerTxt;

	@FindBy(xpath="//select[@id='language']")
	private static WebElement createUserLanguageTxt;

	@FindBy(xpath="//select[@id='altLanguage']")
	private static WebElement createUserAlternateLanguageTxt;

	@FindBy(xpath="//i[@class='icon-unlock-user icon-font6']")
	private static WebElement createUserUnlockUserIcon;

	@FindBy(xpath="//i[@class='icon-move-user icon-font6']")
	private static WebElement createUserMoveUserIcon;

	@FindBy(xpath="//i[@class='icon-reset icon-font6']")
	private static WebElement createUserResetIcon;

	@FindBy(xpath="//span[@id='btnLoadFrom']")
	private static WebElement createUserLoadFromIcon;	
	
	@FindBy(xpath="//span[@id='btnSave']//i[@class='icon-save icon-font6']")
	private static WebElement createUserSaveIcon;

	@FindBy(xpath="//i[@class='icon-close icon-font6']")
	private static WebElement createUserCloseIcon;

	@FindBy(xpath="//input[@id='emailId']")
	private static WebElement createUserEmailTxt;
	
	@FindBy(xpath="//input[@id='Emailpassword']")
	private static WebElement createUserEmailPasswordTxt;
	
	@FindBy(xpath="//input[@id='phone']")
	private static WebElement createUserPhoneTxt;
	
	@FindBy(xpath="//input[@id='mobile']")
	private static WebElement createUserMobileTxt;
	
	@FindBy(xpath="//input[@id='IsEmailAuthPermission']")
	private static WebElement createUserSetPermissionForEmailAuthorizationChkBox;
	
	@FindBy(xpath="//select[@id='DominSecurityMping']")
	private static WebElement createUserDomainUserMappingDrpDwn;
	
	@FindBy(xpath="//input[@id='domainUserName']")
	private static WebElement createUserDomainUserMappingTxt;
	
	@FindBy(xpath="//input[@id='ipAddr1']")
	private static WebElement createUserUserValuesTxt1;
	
	@FindBy(xpath="//input[@id='ipAddr2']")
	private static WebElement createUserUserValuesTxt2;
	
	@FindBy(xpath="//input[@id='ipAddr3']")
	private static WebElement createUserUserValuesTxt3;
	
	@FindBy(xpath="//input[@id='ipAddr4']")
	private static WebElement createUserUserValuesTxt4;
	
	@FindBy(xpath="//input[@id='ipAddr5']")
	private static WebElement createUserUserValuesTxt5;
	
	@FindBy(xpath="//select[@id='typeOfUser']")
	private static WebElement createUserUserTypeDrpDwn;
	
	@FindBy(xpath="//input[@id='cmbUserTypeMaster']")
	private static WebElement createUserUserTypeCombo;
	
	@FindBy(xpath="//select[@id='crmRoles']")
	private static WebElement createUserCRMRolesDrpDwn;
	
	@FindBy(xpath="//input[@id='UserTypeMobile']")
	private static WebElement createUserMobileChkBox;
	
	@FindBy(xpath="//input[@id='UserTypeOffline']")
	private static WebElement createUserOfflineChkBox;
	
	@FindBy(xpath="//input[@id='UserTypeWeb']")
	private static WebElement createUserWebChkBox;
	
	
	

	@FindBy(xpath="//span[contains(text(),'Create User')]")
	private static WebElement createUserMenu;

    
	@FindBy(xpath="//*[@id='dashName']")
	private static WebElement dashboard;
	
	
	 @FindBy(xpath="//input[@id='donotshow']")
		private static WebElement doNotShowCheckbox;
	    
	    @FindBy(xpath="//span[@class='pull-right']")
		private static WebElement closeBtnInDemoPopupScreen;
	
	  //Home Menu
		@FindBy(xpath="//*[@id='2']/span")
		private static WebElement  homeCompanyMenu; 
		
		//Home Menu Security Menu
		@FindBy(xpath="//*[@id='6']/span")
		private static WebElement  homeSecurityMenu; 
		
			//Home Menu Password Policy Menu
			@FindBy(xpath="//span[contains(text(),'Password policy')]")
			private static WebElement  homeSecurityPasswordPolicyMenu;  
			
				@FindBy(xpath="//div[@class='navText']//span[contains(text(),'Password policy')]")
				private static WebElement  passwordPolicyScreenLabel; 
				
				@FindBy(xpath="//input[@id='policyName']")
				private static WebElement  passwordPolicyPolicyNameTxt; 
				
				@FindBy(xpath="//input[@id='MinPassLength']")
				private static WebElement  passwordPolicyMinPasswordLengthTxt; 
				
				@FindBy(xpath="//select[@id='complexity']")
				private static WebElement  passwordPolicyComplexityDropdown; 
				
				@FindBy(xpath="//input[@id='Donotallprevious']")
				private static WebElement  passwordPolicyDonotallpreviousTxt; 
				
				@FindBy(xpath="//i[@class='icon-font6 icon-new']")
				private static WebElement  passwordPolicyAddIcon; 
				
				@FindBy(xpath="//i[@class='icon-font6 icon-delete']")
				private static WebElement  passwordPolicyDeleteIcon; 
				
				@FindBy(xpath="//i[@class='icon-font6 icon-cancel']")
				private static WebElement  passwordPolicyCancelIcon; 
				
				@FindBy(xpath="//span[@id='btnSaveId']//i[@class='icon-save icon-font6']")
				private static WebElement  passwordPolicySaveIcon; 
				
				@FindBy(xpath="//i[@class='icon-close icon-font6']")
				private static WebElement  passwordPolicyCloseIcon;
			
			@FindBy(xpath="//span[contains(text(),'Create Profile')]")
			private static WebElement  homeSecurityCreateProfileMenu;   
			
			@FindBy(xpath="//span[contains(text(),'Create Role')]")
			private static WebElement  homeSecurityCreateRoleMenu; 
			
			@FindBy(xpath="//span[contains(text(),'Create User')]")
			private static WebElement  homeSecurityCreateUserMenu; 
		
		@FindBy(xpath="//div[@class='navText']//span[contains(text(),'Create Profile')]")
		private static WebElement  createProfileLabel; 
		

		@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/ul/li[1]/a/span")
		private static WebElement financialsTransactionsExpandBtn;

		@FindBy(xpath="//li[@class='treeview']//a[@id='2007']/i")
		private static WebElement financialsTransactionsPurchasesExpandBtn;
		
		
		@FindBy(xpath="//*[@id='Profilemenu']/li[2]/ul/li[1]/ul/li/a/span")
		private static List<WebElement> financialsTransactionsList;
		
		@FindBy(xpath="//li[@class='treeview']//a[@id='2001']//span[contains(text(),'Cash and Bank')]")
		private static WebElement financialsTransactionsCashAndBank;
		
		@FindBy(xpath="//li[@class='active']//a[@id='2007']//span[contains(text(),'Purchases')]")
		private static WebElement financialsTransactionsPurchases;
		
		@FindBy(xpath="//li[@class='treeview']//a[@id='2014']//span[contains(text(),'Sales')]")
		private static WebElement financialsTransactionsSales;
		
		@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Journals')]")
		private static WebElement financialsTransactionsJournals;
		
		@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Auto Postings')]")
		private static WebElement financialsTransactionsAutoPostings;
		
				@FindBy(xpath="//input[@id='profileName']")
				private static WebElement  createProfileProfileNameCombo; 
				
				@FindBy(xpath="//ul[@id='Profilemenu']/li[1]/a/span")
				private static WebElement  createProfileHomeExpandBtn; 
				
				@FindBy(xpath="//ul[@id='Profilemenu']/li[2]/a/span")
				private static WebElement  createProfileFinancalsExpandBtn; 
				
				@FindBy(xpath="//ul[@id='Profilemenu']/li[3]/a/span")
				private static WebElement  createProfileInventoryExpandBtn; 
				
				@FindBy(xpath="//ul[@id='Profilemenu']/li[4]/a/span")
				private static WebElement  createProfileFixedAssestsExpandBtn; 
				
						@FindBy(xpath="//li[@class='treeview']//a[@id='50002']//span[contains(text(),'Asset Usage')]")
						private static WebElement transAssetUsageOption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Add Asset Value')]")
						private static WebElement transAddAssetValueOption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Reduce Asset Value')]")
						private static WebElement transReduceAssetValueoption;
			
			
						@FindBy(xpath="//li[@class='treeview']//a[@id='50005']//span[contains(text(),'Disposal Of Asset')]")
						private static WebElement transDisposalOfAssetOption;
			
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Post Depreciation Entries')]")
						private static WebElement transPostDepreciationEntriesoption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Capitalization of Asset')]")
						private static WebElement transCapitalizationofAssetoption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Multiple Asset Usage')]")
						private static WebElement transMultipleAssetUsageoption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Add Depreciation Years')]")
						private static WebElement transAddDepreciationYearsoption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Transfer Of Asset Request')]")
						private static WebElement transTransferOfAssetRequestoption;
			
						@FindBy(xpath="//li[@class='treeview']//a[@id='50004']//span[contains(text(),'Transfer Of Asset')]")
						private static WebElement transTransferOfAssetoption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Transfer of asset')]")
						private static WebElement reportsTransferofassetoption;
			
			
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Assets Usage')]")
						private static WebElement reportsAssetsUsageOption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Add Assets value')]")
						private static WebElement reportsAddAssetsvalueOption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Component Added')]")
						private static WebElement reportsComponentAddedOption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Reduce asset value')]")
						private static WebElement reportsReduceassetvalueOption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Component Reduced')]")
						private static WebElement reportsComponentReducedOption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Disposal of Asset')]")
						private static WebElement reportsDisposalofAssetOption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Fixed Asset Report')]")
						private static WebElement reportsFixedAssetReportOption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Depreciation Schedule')]")
						private static WebElement reportsDepreciationScheduleOption;
			
						@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Disposal Of Asset Request')]")
						private static WebElement reportsDisposalOfAssetRequestOption;
				
				@FindBy(xpath="//ul[@id='Profilemenu']/li[5]/a/span")
				private static WebElement  createProfileProductionExpandBtn; 
				
				@FindBy(xpath="//ul[@id='Profilemenu']/li[6]/a/span")
				private static WebElement  createProfilePointOfSaleExpandBtn; 
				
				@FindBy(xpath="//ul[@id='Profilemenu']/li[7]/a/span")
				private static WebElement  createProfileQualityControlExpandBtn; 
				
				@FindBy(xpath="//ul[@id='Profilemenu']/li[8]/a/span")
				private static WebElement  createProfileSettingExpandBtn; 
				
				@FindBy(xpath="//span[contains(text(),'Save')]/i")
				private static WebElement  createProfileSaveIcon; 
				
				@FindBy(xpath="//i[@class='icon-delete icon-font6']")
				private static WebElement  createProfileDeleteIcon; 
				
				@FindBy(xpath="//i[@class='icon-close icon-font6']")
				private static WebElement  createProfileCloseIcon; 
				
				@FindBy(xpath="//i[@class='icon-font6 icon-selectall']")
				private static WebElement  createProfileSelectAllIcon; 
				
				@FindBy(xpath="//i[@class='icon-font6 icon-unselectall1']")
				private static WebElement  createProfileUnSelectAllIcon; 
		
				
		@FindBy(xpath="//div[@class='navText']//span[contains(text(),'Create Role')]")
		private static WebElement  createRoleLabel; 
		
		
		
		
		

@FindBy(xpath="//*[contains(text(),'ProfileLoadFrom')]")
private static WebElement avaliableProfileProfileLoadFrom ;

@FindBy(xpath="//span[2]//span[1]")
private static WebElement  createRoleMOveFiledsFromLeftSideToRightSide; 


@FindBy(xpath="//span[3]//span[1]")
private static WebElement createRoleMOveFiledsFromRigheSideToLeftSide;
		
				@FindBy(xpath="//input[@id='roleName']")
				private static WebElement createRoleRoleNameCombo; 
				
				@FindBy(xpath="//tbody[@id='roleName_table_body']/tr/td[2]")
				private static List<WebElement> createRoleRoleNameComboList;
				
				@FindBy(xpath="//select[@id='passwordPolicy']")
				private static WebElement  createRolePasswordPolicyDropdown; 
				
				@FindBy(xpath="//span[contains(text(),'Save')]")
				private static WebElement  createRoleSaveicon; 
				
				@FindBy(xpath="//i[@class='icon-delete icon-font6']")
				private static WebElement  createRoleDeleteIcon; 
				
				@FindBy(xpath="//i[@class='icon-close icon-font6']")
				private static WebElement  createRoleCloseIcon; 
				
				@FindBy(xpath="//span[2]//span[1]")
				private static WebElement  createRoleMOveFiledsFronLeftSideToRightSide; 
				
				@FindBy(xpath="//*[@id='myNavbar']/ul/li/span[2]/i")
				private static WebElement CreateRoleDeleteButtonImage; 
		
		/*@FindBy(xpath="//*[@id='idGlobalError']/div/table/tbody/tr/td[2]/div[2]")
		public static WebElement newErrorMessage;*/
		 
		@FindBy(xpath="//span[@class='icon-reject2 theme_color']")
		public static WebElement newErrorMessageCloseBtn;
		
		@FindBy(xpath="//div[@id='idGlobalError']/div/table/tbody/tr/td[3]/span")
		public static List<WebElement> errorMessageCloseBtnList;
		
	
	public static boolean checkPasswordPolicyTestSave() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));	
		securityMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicySubMenu));
		passwordPolicySubMenu.click();
		
		Thread.sleep(2000);
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyAddBtn));
		passwordPolicyAddBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(policyName));
		policyName.sendKeys("MasterPassword");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLength));
		passwordLength.sendKeys("2");	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicySaveBtn));
		passwordPolicySaveBtn.click();
		
		String expMessage="Password policy created successfully";
		
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
	

	
	public static boolean checkUpdateTestPasswordPolicy() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		int passwordListCount = passwordList.size();
		
		for (int i = 0; i < passwordListCount; i++) 
		{
			String data = passwordList.get(i).getText();
			
			if (data.equalsIgnoreCase("MasterPassword")) 
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
		passwordLength.sendKeys("5");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicySaveBtn));		
        passwordPolicySaveBtn.click();
        
        String expMessage="Password policy updated successfully";
		
		String actMessage=checkValidationMessage(expMessage);
		
		Thread.sleep(2000);
	 
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	@FindBy(xpath="//*[@id='PolicyListDiv']/li")
	private static List<WebElement> passwordList;
	
	public static boolean checkAlertOnDeletePasswordPolicy() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		int passwordListCount = passwordList.size();
		
		for (int i = 0; i < passwordListCount; i++) 
		{
			String data = passwordList.get(i).getText();
			
			if (data.equalsIgnoreCase("MasterPassword")) 
			{
				passwordList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyDeletebtn));
		 passwordPolicyDeletebtn.click();
		 
		 getWaitForAlert();
		  		 
	     String getExpectedCompMsg="Are you sure that you want to delete this policy?";
	     
	     String actCompanySaveMsg=getDriver().switchTo().alert().getText();
	     
	     System.out.println("actCompanySaveMsg : "+actCompanySaveMsg);
	     
	     getDriver().switchTo().alert().accept();
		   
		 String expMessage="Policy removed Successfully";
			
	     String actMessage=checkValidationMessage(expMessage);
	     
	     Thread.sleep(2000);
	     
	     if(actCompanySaveMsg.equalsIgnoreCase(getExpectedCompMsg) && actMessage.equalsIgnoreCase(expMessage))
	     {
	    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
	    	 createUserCloseIcon.click();
	    	return true;
	     }
	     else
	     {
	    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
	    	 createUserCloseIcon.click();
	    	 return false;
	     }
	     
    	}
	
	
	
	@FindBy(xpath="//*[@id='Profilemenu']/li/a/span")
	private static List<WebElement> profileMenuPanelList;
		
	public boolean CheckSavingAllProfileSTwithAlloptionsEnabled() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();
	        
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		createProfileProfileNameCombo.click();
		createProfileProfileNameCombo.sendKeys("MasterProfile");
		
		int profileMenuPanelListCount = profileMenuPanelList.size();
		
		for(int i=0;i<profileMenuPanelListCount;i++)
		{
			profileMenuPanelList.get(i).click();
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileSelectAllIcon));
			createProfileSelectAllIcon.click();
		}
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileSaveIcon));
	    createProfileSaveIcon.click();
	    
	    System.out.println("***********************************************CheckSavingAllProfileSTwithAlloptionsEnabled************************************************");
	    
	    String expMessage="Profile saved Successfully";
		
	    String actMessage=checkValidationMessage(expMessage);
	    
	    System.out.println("Saving Message of Profile : " +actMessage+" Value Expected : "+expMessage);
	    
	    Thread.sleep(2000);
	    
	    /*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileCloseIcon));
	    createProfileCloseIcon.click();*/
	    
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
			return true;
		} 
		else 
		{
			return false;
		}
	   
	}
	
	
	
	

	 @FindBy(xpath="//*[@id='restrictionsDiv']/li/div/label/input")
	 private static List<WebElement> restrictionTabChkBoxList;
	 
	 @FindBy(xpath="//*[@id='restrictionsDiv']/li/div/label/span")
	 private static List<WebElement> restrictionTabChkBoxNameList;
	 
	 
	 @FindBy(xpath="//*[@id='Profilemenu']/li[2]/ul/li[1]/ul/li[3]/a/span")
		private static WebElement financialsTransactionsSalesExpandBtn;
		
		@FindBy(xpath="//*[@id='Profilemenu']/li[2]/ul/li[1]/ul/li[3]/ul/li/a/span")
		private static List<WebElement> financialsTransactionsSalesOptionsList;
		
		@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Sales Invoices')]")
		private static WebElement financialsTransactionsSalesInvoiceOption;
		
		@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Sales Returns')]")
		private static WebElement financialsTransactionsSalesReturnsOption;
		
		@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Auto Sales Invoice')]")
		private static WebElement financialsTransactionsAutoSalesInvoiceOption;
		
		@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Sales Invoice N')]")
		private static WebElement financialsTransactionsSalesInvoiceNOption;
		
		@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Cash Sales')]")
		private static WebElement financialsTransactionsCashSalesOption;
		
		@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Hire-Purchase Sales')]")
		private static WebElement financialsTransactionsHirePurchaseSalesOption;
		
		
		
		@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Purchases Vouchers')]")
		private static WebElement financialsTransactionsPurchasesVoucherOption;
		
	 @FindBy(xpath="//tbody[@id='profileName_table_body']/tr/td[2]")
	 private static List<WebElement> createProfileProfileNameList;
	 
	 
	 
	 
	public static boolean checkUpdatingThreSavedProfile() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();
		
		 Thread.sleep(2000);*/
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		 createProfileProfileNameCombo.click();
		 createProfileProfileNameCombo.sendKeys(Keys.SPACE);
		 
		 int actSize=createProfileProfileNameList.size();
		 
		 for (int i = 0; i < actSize; i++) 
		 {
			 String actProfileName=createProfileProfileNameList.get(i).getText();
		 	 if (actProfileName.equalsIgnoreCase("MasterProfile"))
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
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileSaveIcon));
		 createProfileSaveIcon.click();
		 
		String expMessage = "Profile Updated Successfully";
		String actMessage = checkValidationMessage(expMessage);
		 
		 String actRestrictionTabCheckboxSelected = Boolean.toString(selected);
		 
		 String expRestrictionTabCheckboxSelected = "false";
		 
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
	
	
	
	
	
	
	
	
	public static boolean checkDeleteUpdatedProfile() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();
	        
		Thread.sleep(3000);*/
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		 createProfileProfileNameCombo.click();
		 createProfileProfileNameCombo.sendKeys(Keys.SPACE);
		 
		 int actSize=createProfileProfileNameList.size();
		 
		 for (int i = 0; i < actSize; i++) 
		 {
			 String actProfileName=createProfileProfileNameList.get(i).getText();
		 	 if (actProfileName.equalsIgnoreCase("MasterProfile"))
		 	 {
		 		createProfileProfileNameList.get(i).click();
		 		
		 		createProfileProfileNameCombo.sendKeys(Keys.TAB);
		 		
		 		break;
			 }
		 }
		 
		 Thread.sleep(2000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileDeleteIcon));
		 createProfileDeleteIcon.click();
		
		 getWaitForAlert();
		 
		 String actAlert=getAlert().getText();
		 
		 String expAlert="are you sure that you want to delete Profile";
		 	 
		 getAlert().accept();
		 
		 String expMessage1="Profile deleted Successfully";
		 
		 String actMessage1=checkValidationMessage(expMessage1);
		 
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileCloseIcon));
		 createProfileCloseIcon.click();
		 
		 System.out.println("Validation Message                 : *"+actMessage1    +"*  value expected  "+expMessage1);
		 
		 Thread.sleep(2000);
		 
		 if (actMessage1.equalsIgnoreCase(expMessage1))
		 {
			 return true;
		 } 
		else 
		{
			return false;
		}
	}

	
	
	
	
	
	@FindBy(xpath="//span[contains(text(),'Security')]")
	private static WebElement securitymenu ;
	
	@FindBy(xpath="//span[contains(text(),'Create Role')]")
	private static WebElement createRoleMenu;
	
	
	@FindBy(xpath="//*[contains(text(),'All Profile')]")
	private static WebElement avaliableAllProfile ;
	
	public static boolean checkSaveingRoleWithAllfieldsInAllTabs() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	     homeMenu.click();

		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securitymenu));
		  securitymenu.click();
			
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleMenu));
		  createRoleMenu.click();

		  Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
			createRoleRoleNameCombo.click();
			
			createRoleRoleNameCombo.sendKeys("MasterRole");
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
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
			createRoleSaveicon.click();

			String expMessage="Role saved Successfully";
		  
			String actMessage=checkValidationMessage(expMessage);
		  
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
	
	
	
	
	
	@FindBy(xpath="//a[contains(text(),'Assigned Profiles')]")
	private static WebElement  assignedProfilesTab;

	@FindBy(xpath="//a[contains(text(),'Additions')]")
	private static WebElement  additionTab;

	@FindBy(xpath="//a[contains(text(),'Exclusions')]")
	private static WebElement  exclusionsTab;
	
	
	
	@FindBy(xpath="//div[@id='Additions']//label[@class='Flabel font-4'][contains(text(),'Menu')]")
	private static WebElement menuLabel;

	@FindBy(xpath="//ul[@id='Addmenu']//a[@id='1']//i[@class='icon-expand']")
	private static WebElement  addHomeExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']//a[@id='60']//i[@class='icon-expand']")
	private static WebElement  addFininicalExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']//a[@id='135']//i[@class='icon-expand']")
	private static WebElement  addInventoryExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']//a[@id='136']//i[@class='icon-expand']")
	private static WebElement  addFixedAssestsExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']//a[@id='3000']//i[@class='icon-expand']")
	private static WebElement  addProductionExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']//a[@id='8000']//i[@class='icon-expand']")
	private static WebElement  addPointOfSaleExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']//a[@id='150']//i[@class='icon-expand']")
	private static WebElement  addQualityExpandBtn; 

	@FindBy(xpath="//ul[@id='Addmenu']//a[@id='16']//i[@class='icon-expand']")
	private static WebElement  addSettingExpandBtn; 

	@FindBy(xpath="//li[@class='active']//a[@id='61']//i[@class='icon-expand']")
	private static WebElement  addFinTransationExpandBtn;

	@FindBy(xpath="//li[@class='active']//a[@id='2007']//i[@class='icon-expand']")
	private static WebElement addFinTransationpurchaseExpandBtn;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Purchases Vouchers')]")
	private static WebElement addFinTransPurPurchasesVouchers;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Purchases Returns')]")
	private static WebElement addFinTransPurPurchasesReturns;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Purchase Vouchers N')]")
	private static WebElement addFinTransPurPurchaseVouchersN;

	@FindBy(xpath="//label[contains(text(),'Always suspend on saving')]//input[@id='2008']")
	private static WebElement purAlwaysSuspendChkBox;

	@FindBy(xpath="//label[contains(text(),'Hide Menu')]//input[@id='2008']")
	private static WebElement purHideMenuChkbox;

	@FindBy(xpath="//li[1]//div[1]//label[1]//input[1]")
	private static WebElement finPurAccessChkbox;

	@FindBy(xpath="//li[2]//div[1]//label[1]//input[1]")
	private static WebElement finPurAddChkbox;

	@FindBy(xpath="//li[3]//div[1]//label[1]//input[1]")
	private static WebElement finPurEditChkbox;

	@FindBy(xpath="//label[contains(text(),'Delete')]//input[@id='2008']")
	private static WebElement finPurDeleteChkbox;

	@FindBy(xpath="//li[5]//div[1]//label[1]//input[1]")
	private static WebElement finPurPrintChkbox;

	@FindBy(xpath="//li[6]//div[1]//label[1]//input[1]")
	private static WebElement finPurRePrintChkbox;

	@FindBy(xpath="//label[contains(text(),'Print Barcode')]//input[@id='2008']")
	private static WebElement finPurPrintBarcodeChkbox;

	@FindBy(xpath="//label[contains(text(),'Print Un-Authorized documents')]//input[@id='2008']")
	private static WebElement finPurPrintUnAuthorizedDocumentsChkbox;

	@FindBy(xpath="//label[contains(text(),'Export')]//input[@id='2008']")
	private static WebElement finPurExportChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit documents entered by other')]//input[@id='2008']")
	private static WebElement finPurEditDocumentsEnteredbyOtherChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit documents that are checked')]//input[@id='2008']")
	private static WebElement finPurEditDocumentsThatAreCheckedChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit Reconciled documents')]//input[@id='2008']")
	private static WebElement finPurEditReconciledDocumentsChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit documents authorization by higher ups')]//input[@id='2008']")
	private static WebElement finPurEditDocumentsAuthorizationByHigherUpsChkbox;

	@FindBy(xpath="//label[contains(text(),'Enter documents that exceed limit')]//input[@id='2008']")
	private static WebElement finPurEnterDocumentsThatExceedLimitChkbox;

	@FindBy(xpath="//label[contains(text(),'Enter documents that make Cash or Bank Balance Neg')]//input[@id='2008']")
	private static WebElement finPurChkboxEnterDocumentsThatMakeCashorBankBalanceNeg;

	@FindBy(xpath="//label[contains(text(),'Edit documents that are Re-Printed')]//input[@id='2008']")
	private static WebElement finPurEditDocumentsThatAreRePrintedChkbox;

	@FindBy(xpath="//label[contains(text(),'Access through API')]//input[@id='2008']")
	private static WebElement finPurAccessThroughAPIChkbox;

	@FindBy(xpath="//label[contains(text(),'Always suspend on saving')]//input[@id='2008']")
	private static WebElement finPurAlwaysSuspendOnSavingChkbox;

	@FindBy(xpath="//label[contains(text(),'Change print layout')]//input[@id='2008']")
	private static WebElement finPurChangePrintLayoutChkbox;

	@FindBy(xpath="//li[20]//div[1]//label[1]//input[1]")
	private static WebElement finPurAuthorizeChkbox;

	@FindBy(xpath="//label[contains(text(),'Reject')]//input[@id='2008']")
	private static WebElement finPurRejectChkbox;

	@FindBy(xpath="//label[contains(text(),'Show home page')]//input[@id='2008']")
	private static WebElement finPurShowHomepageChkbox;

	@FindBy(xpath="//label[contains(text(),'Stop')]//input[@id='2008']")
	private static WebElement finPurStopChkbox;

	@FindBy(xpath="//label[contains(text(),'Modify Settings')]//input[@id='2008']")
	private static WebElement finPurModifySettingsChkbox;

	@FindBy(xpath="//label[contains(text(),'View document entered by others')]//input[@id='2008']")
	private static WebElement finPurViewDocumentEnteredByOthersChkbox;

	@FindBy(xpath="//label[contains(text(),'Close Links')]//input[@id='2008']")
	private static WebElement finPurCloseLinksChkbox;

	@FindBy(xpath="//label[contains(text(),'Save Revision')]//input[@id='2008']")
	private static WebElement finPurSaveRevisionChkbox;

	@FindBy(xpath="//label[contains(text(),'Show Posting Details')]//input[@id='2008']")
	private static WebElement finPurShowPostingDetailsChkbox;

	@FindBy(xpath="//label[contains(text(),'Suspend')]//input[@id='2008']")
	private static WebElement finPurSuspendChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit exported (Sync) documents')]//input[@id='2008']")
	private static WebElement finPurEditExportedChkbox;

	@FindBy(xpath="//label[contains(text(),'Edit imported (Sync) documents')]//input[@id='2008']")
	private static WebElement finPurEditImportedChkbox;

	@FindBy(xpath="//label[contains(text(),'Hide Menu')]//input[@id='2008']")
	private static WebElement finPurHideMenuChkbox;

	@FindBy(xpath="//label[contains(text(),'Amend')]//input[@id='2008']")
	private static WebElement finPurAmendChkbox;

	@FindBy(xpath="//label[contains(text(),'Customize link')]//input[@id='2008']")
	private static WebElement finPurCustomizeLinkChkbox;

	@FindBy(xpath="//label[contains(text(),'Filter home page')]//input[@id='2008']")
	private static WebElement finPurFilterHomePageChkbox;

	@FindBy(xpath="//label[contains(text(),'Add to stock')]//input[@id='2008']")
	private static WebElement finPurAddToStockChkbox;

	@FindBy(xpath="//label[contains(text(),'Reverse Entry')]//input[@id='2008']")
	private static WebElement finPurReverseEntryChkbox;

	@FindBy(xpath="//label[contains(text(),'Customize info panel')]//input[@id='2008']")
	private static WebElement finPurCustomizeInfoPanelChkbox;

	@FindBy(xpath="//label[contains(text(),'Raise Cheque Return')]//input[@id='2008']")
	private static WebElement finPurRaiseChequeReturnChkbox;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Request for Quote')]")
	private static WebElement addRequestForQuoteOption;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Purchases Quotations')]")
	private static WebElement addPurchaseQuotationOption;

	@FindBy(xpath="//li[@class='active']//a[@id='2012']//span[contains(text(),'Purchases Orders')]")
	private static WebElement addPurchaseorderOption;

	@FindBy(xpath="//li[@class='active']//span[contains(text(),'Material Receipt Notes')]")
	private static WebElement addMaterialRecepitNotes;


	 @FindBy(xpath="//ul[@id='Addmenu']/li[8]/a/span")
	 private static WebElement createRoleAddTabSettingsMenu;
	 
	 @FindBy(xpath="//div[@id='Additions']//div[@id='controlBtns']//span[1]")
	 private static WebElement createRoleAddTabSelectAllChkBox;

	 @FindBy(xpath="//div[@id='Additions']//div[@id='controlBtns']//span[2]")
	 private static WebElement createRoleAddTabUnSelectAllChkBox;
	 
@FindBy(xpath="//*[@id='Addmenu']/li[2]/ul/li/a/span")
private static List<WebElement> additionTabFinancialsMenuList;

@FindBy(xpath="//*[@id='Addmenu']/li[2]/ul/li[1]/a/span")
private static WebElement additionTabFinancialsTransaction;

@FindBy(xpath="//*[@id='Addmenu']/li[2]/ul/li[1]/ul/li/a/span")
private static List<WebElement> additionTabFinancialsTransactionMenusList;

@FindBy(xpath="//*[@id='Addmenu']/li[2]/ul/li[1]/ul/li[2]/ul/li/a/span")
private static List<WebElement> additionTabFinancialsTransactionsPurchasesMenusList;




@FindBy(xpath="//div[@class='Fcheckbox']/label/input[@name='Alwayssuspendonsaving']")
private static WebElement CRAddAlwaysSuspendOnSavingChkBox;

@FindBy(xpath="//div[@class='Fcheckbox']/label/input[@name='HideMenu']")
private static WebElement CRAddHideMenuChkBox;

	//Request For Quote
	
	public static boolean checkEditAndUpdateSavedRole() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		createRoleRoleNameCombo.sendKeys("MasterRole");
		Thread.sleep(2000);
		createRoleRoleNameCombo.sendKeys(Keys.TAB);
		
		//Additions Tab
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(additionTab));
		additionTab.click();
		  
		Thread.sleep(2000);
		  
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addFininicalExpandBtn));
		addFininicalExpandBtn.click();
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addFinTransationExpandBtn));
		addFinTransationExpandBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addFinTransationpurchaseExpandBtn));
		addFinTransationpurchaseExpandBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addFinTransPurPurchasesVouchers));
		addFinTransPurPurchasesVouchers.click();
		*/
		
		  /*int actPurchasesMenusCount=additionTabFinancialsTransactionsPurchasesMenusList.size();
		  int expPurchasesMenusCount=2;
		  
		  System.out.println("additionTabFinancialsTransactionsPurchasesMenusList : "+actPurchasesMenusCount);
		  System.out.println("additionTabFinancialsTransactionsPurchasesMenusList : "+expPurchasesMenusCount);
		  
		
	
		  boolean actCRAddAlwaysSuspendOnSavingChkBox				=CRAddAlwaysSuspendOnSavingChkBox.isDisplayed();
		  boolean actCRAddHideMenuChkBox							=CRAddHideMenuChkBox.isDisplayed();
		 
		  
		  
		  boolean expCRAddAlwaysSuspendOnSavingChkBox				=true;
		  boolean expCRAddHideMenuChkBox							=true;
		  
		  System.out.println("CRAddAlwaysSuspendOnSavingChkBox           : "+actCRAddAlwaysSuspendOnSavingChkBox          +"  value expected  "+expCRAddAlwaysSuspendOnSavingChkBox);
		  System.out.println("CRAddHideMenuChkBox                        : "+actCRAddHideMenuChkBox                       +"  value expected  "+expCRAddHideMenuChkBox);
	 */
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleAddTabSelectAllChkBox));
		  createRoleAddTabSelectAllChkBox.click();
		  
		 Thread.sleep(2000);
		  
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
	
	public static boolean checkCompareAndDeleteUpdatedRole() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		createRoleRoleNameCombo.sendKeys(Keys.END);
		createRoleRoleNameCombo.sendKeys(Keys.SHIFT,Keys.END);
		createRoleRoleNameCombo.sendKeys("MasterRole");
		
		createRoleRoleNameCombo.sendKeys(Keys.TAB);
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
		createRoleSaveicon.click();
	
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CreateRoleDeleteButtonImage));
		CreateRoleDeleteButtonImage.click();
		
		  getWaitForAlert();
		  
		  String actAlertText = getAlert().getText();
		  String expAlertText = "are you sure that you want to delete Role";

		  getAlert().accept();
		  
		  String expMessage = "Role Successfully deleted..";
		  String actMessage = checkValidationMessage(expMessage);
		  
		
		  System.out.println("AlertText  : "+actAlertText+"  value expected  "+expAlertText);
		  System.out.println("Error Message  : "+actMessage+"  value expected  "+expMessage);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleCloseIcon));
		  createRoleCloseIcon.click();
			
		  Thread.sleep(2000);
		  
	   if ( actAlertText.equalsIgnoreCase(expAlertText) && actMessage.equalsIgnoreCase(expMessage)) 
	    {
			return true;
		} 
		else 
		{
			return false;
		}

	}
	
	
	
	// Restriction for Entry Tab
	 @FindBy(xpath="//div[@id='RestrictionEntryMasters']/ul/li")
	 private static List<WebElement> createRoleRestrictionForEntryTabMastersList;
	 
	 @FindBy(xpath="//input[@id='chkExclusion']")
	 private static WebElement createRoleRestrictionForEntryTabExclChkBox;
	 
	 @FindBy(xpath="//table[@id='restrictionEntryTable']")
	 private static WebElement createRoleRestrictionForEntryTabTable;
	 
	 @FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[1]/td[2]")
	 private static WebElement restrictionForEntryTabTableSelectRow1Col1;
	 
	 @FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[1]/td[3]//input")
	 private static WebElement restrictionForEntryTabRow1EntryChkBox;
	 
	 @FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[1]/td[4]//input")
	 private static WebElement restrictionForEntryTabRow1ReportChkBox;
	 
	 @FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[1]/td[5]//input")
	 private static WebElement restrictionForEntryTabRow1ViewChkBox;
	 
	 
	 @FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[2]/td[2]")
	 private static WebElement restrictionForEntryTabTableSelectRow2Col1;
	 
	 @FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[2]/td[3]//input")
	 private static WebElement restrictionForEntryTabRow2EntryChkBox;
	 
	 @FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[2]/td[4]//input")
	 private static WebElement restrictionForEntryTabRow2ReportChkBox;
	 
	 @FindBy(xpath="//tbody[@id='restrictionEntryTable_body']/tr[2]/td[5]//input")
	 private static WebElement restrictionForEntryTabRow2ViewChkBox;
	 
	 

	 @FindBy(xpath="//select[@id='passwordPolicy']//option[6]")
	 private static WebElement getNumericValue;
	 
	 
	 @FindBy(xpath="//select[@id='language']")
	 private static WebElement getLanguage;

	 
	 @FindBy(xpath="//select[@id='altLanguage']")
	 private static WebElement getAltLanguage;
	 

	 @FindBy(xpath="//select[@id='language']/option")
	 private static List<WebElement> createUserLanguageDropdownList;
	 
	 

	 
	 @FindBy(xpath="//input[@id='MasteroptionControl']")
	 private static WebElement restrictionForEntryTabTableEnterMasterTxt;
	 
	 @FindBy(xpath="//tbody[@id='MasteroptionControl_table_body']/tr/td[2]")
	 private static List<WebElement> restrictionForEntryTabTableMasterComboList;

@FindBy(xpath="//select[@id='altLanguage']/option")
private static List<WebElement> createUserAltLanguageDropdownList;

@FindBy(xpath="//select[@id='ERPRoles']/option")
private static List<WebElement> createUserERPRolesDropdownOptions;
	
	public static boolean checkCreateUserWithRestrictionsForAccount() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
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
		 createUserLoginNameTxt.sendKeys("MasterUser");

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		 createUserERPRoleDropDown.click();
		 createUserERPRoleDropDown.sendKeys("Role ST");
		 
	     int count = createUserERPRolesDropdownOptions.size()-1;
		 System.out.println("Count "+ count);
	    
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		 createUserPasswordTxt.sendKeys("123450");

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		 createUserConfirmPasswordTxt.sendKeys("123450");
		 	
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		 createUserNameTxt.sendKeys("AccountRestriction");
		 	
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
		 createUserLoginAbbrivationTxt.sendKeys("MSU");
		 
		 createUserSecurityQuestionDropDown.click();
		 Select s = new Select(createUserSecurityQuestionDropDown);
		 s.selectByVisibleText("In which county were you born?");

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSecurityAnswerTxt));	
		 createUserSecurityAnswerTxt.sendKeys("India");
		 	 	
		 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getLanguage));
			 getLanguage.click();
			 int counnt = createUserLanguageDropdownList.size()-1;
			 String actLanguageDropdownCount = Integer.toString(counnt);
			 String expLanguageDropdownCount = "2";
			 
			 Set LanguageDropdownList = new HashSet<String>();
			 
			 for (int i = 1; i < counnt; i++) 
			 {
				 String data = createUserLanguageDropdownList.get(i).getText();
				 LanguageDropdownList.add(data);
			 }
			 
			 String actLanguageDropdownList = LanguageDropdownList.toString();
			 String expLanguageDropdownList = "[English, Arabic]";
			 
			 System.out.println("LanguageDropdownCount  : "+actLanguageDropdownCount  +"  value expected  "+expLanguageDropdownCount);
			 System.out.println("LanguageDropdownList   : "+actLanguageDropdownList   +"  value expected  "+expLanguageDropdownList);
			 
			 getLanguage.sendKeys(Keys.TAB);
			 
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getAltLanguage));
				 getAltLanguage.click();
				 int cont = createUserAltLanguageDropdownList.size()-1;
				 String actAltLanguageDropdownCount = Integer.toString(count);
				 String expAltLanguageDropdownCount ="2";
				 
				 Set AltLanguageDropdownList = new HashSet<String>();
				 
				 for (int i = 1; i < cont; i++) 
				 {
					 String data = createUserAltLanguageDropdownList.get(i).getText();
					 AltLanguageDropdownList.add(data);
				 }
				 
				 String actAltLanguageDropdownList = AltLanguageDropdownList.toString();
				 String expAltLanguageDropdownList = "[English, Arabic]";
				 
				 getAltLanguage.sendKeys(Keys.TAB);
				 
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
			createUserEmailTxt.sendKeys("testingg@focussoftnet.com");
			
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserRestrictionforentryTab));
			 createUserRestrictionforentryTab.click();
			 
			 
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
			  
			  
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableSelectRow2Col1));
			  restrictionForEntryTabTableSelectRow2Col1.click();
			  
			  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableEnterMasterTxt));
			  String c  = "C";
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
			  
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSaveIcon));
			 createUserSaveIcon.click();
			 
			 String expMessage = "new user created successfully";
			 String actMessage = checkValidationMessage(expMessage);
			 
			 System.out.println(actMessage);
			 System.out.println(expMessage);
			 
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
	


@FindBy(xpath="//div[@id='userAndGroups']/ul/li//span/span[1]")
private static List<WebElement> createUserSavedUsersList;
	

@FindBy(xpath="//input[@id='searchUsertxtBox']")
private static WebElement searchUserTxt;

	public static boolean checkUpdatingTheCreatedUser() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchUserTxt));
		searchUserTxt.click();
		searchUserTxt.sendKeys("MasterUser");
		searchUserTxt.sendKeys(Keys.END);
		
		Thread.sleep(2000);
		
		boolean actValidationMessage = validationConfirmationMessage.getText().isEmpty();
		boolean expValidationMessage = true;
		
		 int actUsersCount=createUserSavedUsersList.size();
		 int expUsersCount=1;
		 
		 System.out.println("Users Count after SearchUser : "+actUsersCount+"  Value Expected  "+expUsersCount);
		 
		 for (int i = 0; i < actUsersCount; i++) 
		 {
			 String actUser=createUserSavedUsersList.get(i).getText();
			 
			 if (actUser.equalsIgnoreCase("MasterUser")) 
			 {
				 createUserSavedUsersList.get(i).click();
				 
				 Thread.sleep(2000);
				 
				 break;
			 } 
		 }
				 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
			 createUserPasswordTxt.clear();
			 createUserPasswordTxt.sendKeys("543210");
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
			 createUserConfirmPasswordTxt.clear();
			 createUserConfirmPasswordTxt.sendKeys("543210");
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
			 createUserNameTxt.clear();
			 createUserNameTxt.sendKeys("UCreationTest");
			 
			 for(int j=0;j<5;j++)
			 {
				 if(validationConfirmationMessage.getText().isEmpty()==false)
				 {
					 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
					 errorMessageCloseBtn.click();
				 }	
			 }
	 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSaveIcon));
			 createUserSaveIcon.click();
			 
			 String expMessage = "User updated successfully.";
			 String actMessage = checkValidationMessage(expMessage);
			 
			 Thread.sleep(2000);
			 
			 if (actValidationMessage==expValidationMessage && actUsersCount==expUsersCount && actMessage.equalsIgnoreCase(expMessage)) 
			 {
				return true;
			 }
			 else 
			 {
				return false;
			 }
		 }
	
	


	 @FindBy(xpath="//i[@class='icon-delete icon-font6']")
	 private static WebElement createUserDeleteIcon;

	public static boolean checkDeletingUpdatedUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchUserTxt));
		searchUserTxt.click();
		searchUserTxt.sendKeys("MasterUser");
		searchUserTxt.sendKeys(Keys.END);
		
		Thread.sleep(2000);
		
		boolean actValidationMessage = validationConfirmationMessage.getText().isEmpty();
		boolean expValidationMessage = true;
		
		 int actUsersCount=createUserSavedUsersList.size();
		 int expUsersCount=1;
		 
		 System.out.println("Users Count after SearchUser : "+actUsersCount+"  Value Expected  "+expUsersCount);
		 
		 for (int i = 0; i < actUsersCount; i++) 
		 {
			 String actUser=createUserSavedUsersList.get(i).getText();
			 System.out.println(actUser);
			 
			 if (actUser.equalsIgnoreCase("MasterUser"))
			 {
				 createUserSavedUsersList.get(i).click();
				 
				 Thread.sleep(2000);
				 
				 break;
			 }
		 }
		 
		 Thread.sleep(2000);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserDeleteIcon));
		createUserDeleteIcon.click();
		
		Thread.sleep(2000);
		
		getWaitForAlert();
		
		getAlert().accept();
		
		String expMessage = "User is deleted successfully";
		String actMessage = checkValidationMessage(expMessage);
		
		if (actValidationMessage==expValidationMessage && actUsersCount==expUsersCount && actMessage.equalsIgnoreCase(expMessage)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			return true;
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserCloseIcon));
			createUserCloseIcon.click();
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	
	     public MastersPage(WebDriver driver)
	     {
	    	PageFactory.initElements(driver, this);
	    	
	     }
	
	
	
}
