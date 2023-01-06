package com.focus.Pages;

import java.io.IOException;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.Alert;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class CompanyOptionsPage extends BaseEngine
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
				
			  private static Alert alert;
				
			  public static Alert getAlert()
			  {
				alert=getDriver().switchTo().alert();
				return alert;
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
		
		
		
		
		
		@FindBy(xpath="//i[@class='icon-restore-1 icon-font4']")
		private static WebElement  loginPageRestoreBtn;
		
		//User Name and password in Credentials
		@FindBy(xpath="//*[@id='txtAdminUsername']")
		private static WebElement adminName;
		
		@FindBy(xpath="//*[@id='txtAdminPassword']")
		private static WebElement adminPassword;
		
		@FindBy(xpath="/html/body/section/div[8]/div/div/div[4]/button[1]")
		private static WebElement loginButton;
		
		@FindBy(xpath="/html/body/section/div[8]/div/div/div[4]/button[2]")
		private static WebElement cancelButton;

		//Restore Screen Options
		@FindBy(xpath="//div[contains(text(),'Restore Company')]")
		private static WebElement restoreCompanyBtnOfLoginPage;
		
		@FindBy(xpath="//input[@id='txtRestoreFile']")
		private static WebElement restoreInputTxt;
		
		@FindBy(xpath="//button[@class='btn Fbutton']")
		private static WebElement restoreBrowseBtn;
		
		
		
		
		public static boolean checkDeleteCompanyOptionInSearch() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		{
			String actAlertTxt = null;
			String expAlertTxt = "Are you sure you want to Delete the Company";
			
			String actCompanyDeletedAlertTxt = null;
			String expCompanyDeletedAlertTxt = "Company deleted successfully";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
			companyLogo.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
			String actGetLoginCompanyNameInformation = companyName.getText();
			
			if (actGetLoginCompanyNameInformation.startsWith("Automation Company")) 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(getSearch));
				getSearch.click();
							
				getSearch.sendKeys("Delete Company");
				Thread.sleep(2000);
				getSearch.sendKeys(Keys.ENTER);
				
				getWaitForAlert();
				
				actAlertTxt = getAlert().getText();
				
				getAlert().accept();
				
				Thread.sleep(2000);
				
				getWaitForAlert();
				
				actCompanyDeletedAlertTxt = getAlert().getText();
				
				getAlert().accept();
				
				Thread.sleep(2000);
				
				System.out.println(actAlertTxt+"  Value Expected  "+expAlertTxt);
				System.out.println(actCompanyDeletedAlertTxt+"  Value Expected  "+expCompanyDeletedAlertTxt);
				
				getDriver().navigate().refresh();
				
				Thread.sleep(3000);
			}
			
			if (actAlertTxt.equalsIgnoreCase(expAlertTxt) && actCompanyDeletedAlertTxt.equalsIgnoreCase(expCompanyDeletedAlertTxt))
			{
				return true;
			}
			else
			{
				return false;
			}
		}
		
		
		
		
	
	public static boolean checkRestoreCompanyForSanity() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
		dataMangementMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restore));
		restore.click();
		*/
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginPageRestoreBtn));
		loginPageRestoreBtn.click();
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(adminName));
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(adminPassword));
		adminPassword.sendKeys("focus");
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginButton));
		loginButton.click();
		
		Thread.sleep(2000);
		
        boolean actRestoreCompanyBtnOfLoginPage  =restoreCompanyBtnOfLoginPage.isDisplayed();
        boolean actRestoreInputTxt               =restoreInputTxt.isDisplayed();
        boolean actRestoreBrowseBtn              =restoreBrowseBtn.isDisplayed();
        
        boolean expRestoreCompanyBtnOfLoginPage  =true;
        boolean expRestoreInputTxt               =true;
        boolean expRestoreBrowseBtn              =true;
		
        getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreBrowseBtn));
		restoreBrowseBtn.click();
		
		Thread.sleep(3000);
	     
	     Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\SanityRestore.exe");
	     
	     Thread.sleep(3000);
	     
	     getWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtnOfLoginPage));
		 restoreCompanyBtnOfLoginPage.click();
	     
    	 getWaitForAlert();
    	 String actAlertTxt = getAlert().getText();
    	 getAlert().accept();
 		 
 		 System.err.println(actAlertTxt);
 		 
 		String expAlertTxt = "Restore company code : 010";
	    	 
	    Thread.sleep(3000);
	    
		if(actAlertTxt.equalsIgnoreCase(expAlertTxt))
		{
			return true;
		}
		else
		{
			return false;
		}
     }
	
	
	
	
	
	
	public static boolean checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		LoginPage lp=new LoginPage(getDriver()); 
		
	    String unamelt="su";
	    
	    String pawslt="su";
	    
	     lp.enterUserName(unamelt);
	     
	     lp.enterPassword(pawslt);
	     
	     Thread.sleep(2000);
	     
	     lp.clickOnSignInBtn();
	     
	     Thread.sleep(10000);
	     
	     if (reindexingPopup.isDisplayed()==false) 
	     {
			System.out.println("Reindexing Log is Not Displaying");
		 } 
	     else 
	     {
	    	 System.out.println("Reindexing Log is Displaying");
	    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reindexingPopupCancelBtn));
	    	 reindexingPopupCancelBtn.click();
	    	 Thread.sleep(2000);
	    	 lp.clickOnSignInBtn();
		 }
	     
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
	
	
	
	
	@FindBy (xpath="//div[@id='ReIndexingError_Modal']/div/div")
	private static WebElement reindexingPopup;
	
	@FindBy (xpath="(//div[@id='ReIndexingError_Modal']//div/button)[1]")
	private static WebElement reindexingPopupCloseBtn;
	
	@FindBy (xpath="(//div[@id='ReIndexingError_Modal']//div/button)[2]")
	private static WebElement reindexingPopupCancelBtn;
	
	@FindBy (xpath="//div[@id='ReIndexingError_Modal']//div/input")
	private static WebElement reindexingPopupSaveBtn;
	
	
	public static boolean checkLogOutAndLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		LoginPage lp=new LoginPage(getDriver()); 
		
	    String unamelt="su";
	    
	    String pawslt="su";
	    
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
	     
	     /*if (reindexingPopup.isDisplayed()==false) 
	     {
			System.out.println("Reindexing Log is Not Displaying");
		 } 
	     else 
	     {
	    	 System.out.println("Reindexing Log is Displaying");
	    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reindexingPopupCancelBtn));
	    	 reindexingPopupCancelBtn.click();
	    	 Thread.sleep(2000);
	    	 lp.clickOnSignInBtn();
		 }*/
	     
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
			
		    
		     
		    
	 		 
	
	
	
	
	
	
	
	// Edit Company
	
	private static String xlfile;
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;
	private static WebElement loginTitle;
	private static String url;
	private static int cSize;
	private static String userN;	
	private static String getReleaseDateFromAbout;
	private static String modifyGetReleaseDateFromAbout;
	
	@FindBy(xpath="//*[@id='quickAccessDiv']/ul/li[5]/a")
	private static WebElement aboutOption;
	
	
		@FindBy(xpath="//*[@id='about']/div[1]/img")
		private static WebElement aboutImage;
		
		@FindBy(xpath="//*[@id='idAboutCompany']/div[1]/label[2]")
		private static WebElement aboutVersioin;

		@FindBy(xpath="//*[@id='idAboutCompany']/div[2]/label[2]")
		private static WebElement aboutRelease;

		@FindBy(xpath="//*[@id='idAboutCompany']/div[3]/label[2]")
		private static WebElement aboutLicensed;
		
	@FindBy(xpath="//input[@id='txtSearchMenu_MainLayout']")
	private static WebElement searchOption;
	
	@FindBy(xpath="//*[@id='mainHeader_MainLayout']/nav/div/ul/li[2]/a/i")
	private static WebElement productTourOption;
	
	@FindBy(xpath="//*[@id='mainHeader_MainLayout']/nav/div/ul/li[3]/a/i")
	private static WebElement alertsOption;

	@FindBy(xpath="//*[@id='idquickAccessMenus_Main']/span[1]")
	private static WebElement aboutIdQuickAccessMenu;

	@FindBy (xpath="//*[@id='CompanyName']")
	 private static WebElement getCompanyNameTextInEditScreen;
	 
	 @FindBy (xpath="//input[@id='txtSearchMenu_MainLayout']")
	 private static WebElement getSearch;
	 
	 /*@FindBy (xpath="//*[@id='txtSearch']")
	 private static WebElement getSearch;*/
	 
	 @FindBy (xpath="//*[@id='107']")
	 private static WebElement getTextFromSearch;
	
	 @FindBy (xpath="//li[@class='dropdown searchBox hidden-xs open']//li[1]")
	 private static WebElement getSearchOne;
	 
	 @FindBy(xpath="//div[contains(text(),'Ok')]")
	 private static WebElement okButtonInEditCompanyScreen;
	 
	 @FindBy (xpath="//*[@id='CreatedDate']")
	 private static WebElement companyDateInEditCompany;
	
	 @FindBy(xpath="//div[contains(text(),'Cancel')]")
	 private static WebElement cancelButtonInEditCompanyScreen;
	
	 @FindBy(xpath="//*[@id='txtUsername']")
		private static WebElement username;

		@FindBy(id="txtPassword")
		private static WebElement password;
		
	    @FindBy(id="ddlCompany")
	    private static WebElement companyDropDownList;

		@FindBy(id="btnSignin")
		private static WebElement signIn;
		
		@FindBy(xpath="//*[@id='dashName']")
		private static WebElement dashboard;
	    
		@FindBy(id="chkRememberMe")
		private static WebElement rememberMeChk;
		
		@FindBy(id="chkRememberPwd")
		private static WebElement rememberPwdChk;
		
	    @FindBy(xpath="//div[@id='Layout_Popup']//input[2]")
		private static WebElement noIn;
	    
	    @FindBy(xpath="//input[@id='donotshow']")
	   	private static WebElement doNotShowCheckbox;
	       
	    @FindBy(xpath="//span[@class='pull-right']")
	   	private static WebElement closeBtnInDemoPopupScreen;
	       
	    @FindBy(xpath="//*[@id='idRecentMenus_Main']/i")
	   	private static WebElement recentMenus;
	   	
	   	@FindBy(xpath="//img[@class='user-image theme_button_color']")
	   	private static WebElement imgOfTheUser;
	   	
	   	@FindBy(xpath="//span[@class='hidden-xs']")
		private static WebElement textNameOfTheUser;
		
		@FindBy(xpath="//i[@class='fa fa-caret-down']")
		private static WebElement dropdownMenuUserChangePasswordLanguageAndLogout;
	   	
		@FindBy(xpath="//*[@id='quickAccessDiv']/ul/li[1]/a")
		private static WebElement favouriteOption;
		
		@FindBy(xpath="//*[@id='quickAccessDiv']/ul/li[2]/a")
		private static WebElement recentItemsOpion;
		
		@FindBy(xpath="//*[@id='quickAccessDiv']/ul/li[3]/a")
		private static WebElement runOption;
		
		@FindBy(xpath="//*[@id='quickAccessDiv']/ul/li[4]/a/span")
		private static WebElement quickAccessOption;
	   	
	   	
	 public static boolean checkAvailibityOfOptionsInWrapperLayout() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	 {
			/*excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

			 String unamelt=excelReader.getCellData("Sheet1", 47, 6);
			 String passwordlt=excelReader.getCellData("Sheet1", 48, 6);

		     getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
		     username.click();     
	         username.sendKeys(unamelt);
	         
	         getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
			 
	         getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(password));
			 password.sendKeys(passwordlt);
			 getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
			  
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rememberMeChk));
		     rememberMeChk.click();
			 
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rememberPwdChk));
			 rememberPwdChk.click();

			 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(signIn));
			 signIn.click();
			
			 Thread.sleep(5000);*/

			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(aboutIdQuickAccessMenu));
		
			 boolean actaboutIdQuickAccessMenu=aboutIdQuickAccessMenu.isDisplayed();
			 boolean actsearchOption=searchOption.isDisplayed();
			 //boolean actproductTourOption=productTourOption.isDisplayed();
			 boolean actalertsOption=alertsOption.isDisplayed();
			 boolean actrecentMenus=recentMenus.isDisplayed();
			 boolean actimgOfTheUser=imgOfTheUser.isDisplayed();
			 boolean acttextNameOfTheUser=textNameOfTheUser.isDisplayed();
			 boolean actdropdownMenuUserChangePasswordLanguageAndLogout=dropdownMenuUserChangePasswordLanguageAndLogout.isDisplayed();
			 boolean actcompanyLogo=companyLogo.isDisplayed();
			 
			 boolean expaboutIdQuickAccessMenu=true;
			 boolean expsearchOption=true;
			 boolean expproductTourOption=true;
			 boolean expalertsOption=true;
			 boolean exprecentMenus=true;
			 boolean expimgOfTheUser=true;
			 boolean exptextNameOfTheUser=true;
			 boolean expdropdownMenuUserChangePasswordLanguageAndLogout=true;
			 boolean expcompanyLogo=true;
			 
			 System.out.println("aboutIdQuickAccessMenu  : " + actaboutIdQuickAccessMenu   + "  value expected  " + expaboutIdQuickAccessMenu);
			 System.out.println("searchOption            : " + actsearchOption             + "  value expected  " + expsearchOption);
			 System.out.println("alertsOption            : " + actalertsOption             + "  value expected  " + expalertsOption);
			 System.out.println("recentMenus             : " + actrecentMenus              + "  value expected  " + exprecentMenus);
			 System.out.println("imgOfTheUser            : " + actimgOfTheUser             + "  value expected  " + expimgOfTheUser);
			 System.out.println("textNameOfTheUser       : " + acttextNameOfTheUser        + "  value expected  " + exptextNameOfTheUser);
			 System.out.println("dropdownMenuUserChangePasswordLanguageAndLogout: " + actdropdownMenuUserChangePasswordLanguageAndLogout + "  value expected  " + expdropdownMenuUserChangePasswordLanguageAndLogout);
			 System.out.println("companyLogo             : " + actcompanyLogo              + "  value expected  " + expcompanyLogo);
			 
			 boolean actWrapperLayout=actaboutIdQuickAccessMenu==expaboutIdQuickAccessMenu && actsearchOption==expsearchOption && 
					 actalertsOption==expalertsOption && actrecentMenus==exprecentMenus && actimgOfTheUser==expimgOfTheUser && acttextNameOfTheUser==exptextNameOfTheUser && 
				 	 actdropdownMenuUserChangePasswordLanguageAndLogout==expdropdownMenuUserChangePasswordLanguageAndLogout && actcompanyLogo==expcompanyLogo;
			 
			 boolean expWrapperLayout=true;
			 
			 if(actWrapperLayout==expWrapperLayout)
			 {		
				 return true;
			 }							
			else
			{
				return false;
			}
		 }

		
		
		
		public static boolean checkAboutIdQuickAccessMenu() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(aboutIdQuickAccessMenu));
			
			System.out.println("checkAboutIdQuickAccessMenu");
			
			aboutIdQuickAccessMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(aboutOption));
			
			boolean actfavouriteOption          =favouriteOption.isDisplayed();
			boolean actrecentItemsOpion         =recentItemsOpion.isDisplayed();
			boolean actrunOption                =runOption.isDisplayed(); 
			boolean actquickAccessOption        =quickAccessOption.isDisplayed();
			boolean actaboutOption              =aboutOption.isDisplayed();
			
			boolean expfavouriteOption          =true;
			boolean exprecentItemsOpion         =true;
			boolean exprunOption                =true; 
			boolean expquickAccessOption        =true;
			boolean expaboutOption              =true;
			
			 System.out.println("favouriteOption   : "+actfavouriteOption    +"  value expected  "+expfavouriteOption);
			 System.out.println("recentItemsOpion  : "+actrecentItemsOpion   +"  value expected  "+exprecentItemsOpion);
			 System.out.println("runOption         : "+actrunOption          +"  value expected  "+exprunOption);
			 System.out.println("quickAccessOption : "+actquickAccessOption  +"  value expected  "+expquickAccessOption);
			 System.out.println("aboutOption       : "+actaboutOption        +"  value expected  "+expaboutOption);
			
			if(actfavouriteOption==expfavouriteOption && actrecentItemsOpion==exprecentItemsOpion && actrunOption==exprunOption && 
				actquickAccessOption==expquickAccessOption && actaboutOption==expaboutOption)
			{
				excelReader.setCellData(xlfile, "Sheet1", 49, 9, resPass);
			    return true;
			}				
			else
			{
				excelReader.setCellData(xlfile, "Sheet1", 49, 9, resFail);
				return false;
		
			}		
			
		}
		
		
		
		
	
	public static boolean clickOnAbout() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(aboutOption));
		
		System.out.println("clickonabout");
		
		aboutOption.click();
	
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(aboutLicensed));
		
		boolean actaboutImage         =aboutImage.isDisplayed();
		boolean actaboutVersioin      =aboutVersioin.isDisplayed();
		boolean actaboutRelease       =aboutRelease.isDisplayed(); 
		boolean actaboutLicensed      =aboutLicensed.isDisplayed();
		
		boolean expaboutImage         =true;
		boolean expaboutVersioin      =true;
		boolean expaboutRelease       =true; 
		boolean expaboutLicensed      =true;
		
		 System.out.println("aboutImage    : "+actaboutImage    +"  value expected  "+expaboutImage);
		 System.out.println("aboutVersioin : "+actaboutVersioin +"  value expected  "+expaboutVersioin);
		 System.out.println("aboutRelease  : "+actaboutRelease  +"  value expected  "+expaboutRelease);
		 System.out.println("aboutLicensed : "+actaboutLicensed +"  value expected  "+expaboutLicensed);
		 
		if(actaboutImage==expaboutImage && actaboutVersioin==expaboutVersioin && actaboutRelease==expaboutRelease && actaboutLicensed==expaboutLicensed)
		{
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(aboutRelease));
			getReleaseDateFromAbout=aboutRelease.getText();
			modifyGetReleaseDateFromAbout=" "+getReleaseDateFromAbout;

	        getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(aboutIdQuickAccessMenu));
			System.out.println("ClickOnCloseAbout");
			aboutIdQuickAccessMenu.click();	 
		 
		  excelReader.setCellData(xlfile, "Sheet1", 50, 9, resPass);
          return true;
					
		}
		else
		{
			
			 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(aboutIdQuickAccessMenu));
				System.out.println("ClickOnCloseAbout");
				aboutIdQuickAccessMenu.click();	 
				
			excelReader.setCellData(xlfile, "Sheet1", 50, 9, resFail);
			return false;
		}
		
	}
	
	
	
	public static String getReleasedate() throws InterruptedException
	{
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(aboutRelease));
		String getReleaseDateFromAbout=aboutRelease.getText();
		String modifyGetReleaseDateFromAbout=" "+getReleaseDateFromAbout;
		
		return modifyGetReleaseDateFromAbout;
	}
	
	

	
	public static void clickOnAboutToClose()
	{
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(aboutIdQuickAccessMenu));
		System.out.println("ClickOnCloseAbout");
		aboutIdQuickAccessMenu.click();
	}
	
	
	public static boolean checkEditCompanyOptionInSearch() throws EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
		
		getSearch.click();
		getSearch.sendKeys("Edit Company");
		
		
        String serachDropdown= getTextFromSearch.getText();
        

		
		
		System.out.println("recentDropdown  :  "+serachDropdown);
		
		
		boolean getUpdatedLang=false;
		String[] str= serachDropdown.split("\n");   
		for(String sto :str)
		{
		
			sto.equals("Edit Company");
			getUpdatedLang=true;
			break;
		}
		
		if (getUpdatedLang==true)
		{
			excelReader.setCellData(xlfile, "Sheet1", 51, 9, resPass);
			return true;
			
		}
		else
		{
			excelReader.setCellData(xlfile, "Sheet1", 51, 9, resFail);
			return false;
			
		}
		
	
		
	}
	
	
	
	
	@FindBy (xpath="//*[@id='page_Content']/form/div/div[1]/nav/div/div[1]/a/div/span[2]")
    private static WebElement CreateCompanyTitle;
	
	
	public static boolean checkSelectingEditCompanyInSearch() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getTextFromSearch.click();
		
		Thread.sleep(5000);
		
		//Old_LoginPageCompanyCreationPage llccp=new Old_LoginPageCompanyCreationPage(getDriver());
		
		//boolean getCompanyFields=llccp.verifyAvailablityOfCreateCompanyFieldsInEditCompany();
	
		System.out.println("CreateCompanyTitle  :  " + CreateCompanyTitle.getText());
		
		//System.out.println("getCompanyFields  :  " + getCompanyFields);
		
		Thread.sleep(3000);
		
		String actCreateCompanyTitle=CreateCompanyTitle.getText();
		
		boolean expCompanyFields=true;
		String expCreateCompanyTitle="Edit Company";
		
		if(actCreateCompanyTitle.equalsIgnoreCase(expCreateCompanyTitle) /*&& getCompanyFields==expCompanyFields*/)
		{
			excelReader.setCellData(xlfile, "Sheet1", 52, 9, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, "Sheet1", 52, 9, resFail);
			return false;
		}
	}
	
	
	public static boolean checkUpdateTheCompanyAlertThroughSearch() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
    {
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		 
		 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(getCompanyNameTextInEditScreen));
		 getCompanyNameTextInEditScreen.click();
		 getCompanyNameTextInEditScreen.sendKeys(Keys.END);
		 getCompanyNameTextInEditScreen.sendKeys(Keys.SHIFT, Keys.HOME);
		 getCompanyNameTextInEditScreen.sendKeys("Automation Company "+modifyGetReleaseDateFromAbout);
	     
	     Thread.sleep(2000);
	     
	     getAction().moveToElement(getCompanyNameTextInEditScreen).sendKeys(Keys.TAB).perform();
	     companyDateInEditCompany.sendKeys("0201");
	    	
	     Thread.sleep(2000);
	     
	     getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(okButtonInEditCompanyScreen));
	     okButtonInEditCompanyScreen.click();
	     
	     
	     getWaitForAlert();
	     
			
		 String getExpectedCompMsg="Company updated successfully";
					     
		 //Old_LoginPageCompanyCreationPage cco=new Old_LoginPageCompanyCreationPage(getDriver());
		     
		 System.out.println("q");
	     
         String actCompanySaveMsg=getAlert().getText();
	     
	     System.out.println("qq");
	     
	     System.out.println("getExpectedCompMsg  :  "+getExpectedCompMsg);
	     
	     System.out.println("qqq");
	     
	     System.out.println("actCompanySaveMsg  :  "+actCompanySaveMsg);
	     
	     System.out.println("qqqq");
	     
	     
	     if(getExpectedCompMsg.equalsIgnoreCase(actCompanySaveMsg))
	     {
	    	 getAlert().accept();
	    	 excelReader.setCellData(xlfile, "Sheet1", 53, 9, resPass);
	    	 return true;
	     }
	     else
	     {
	    	 excelReader.setCellData(xlfile, "Sheet1", 53, 9, resFail);
	    	 return false;
	     }
	   
    }
	
    
	
    // commented Code Because after Updating Sign In Page Displays """cancelButtonInEditCompanyScreen"" and ""DashBoard"" 
    public static boolean clickOnCloseTheEditCompanyThroughSearch() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
    {
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
    	
		/*Thread.sleep(5000);
		
     	getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(cancelButtonInEditCompanyScreen));
    	cancelButtonInEditCompanyScreen.click();*/
    	
   	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
    	
    	
    	getFluentWebDriverWait().until(ExpectedConditions.textToBePresentInElement(userNameDisplay, "SU"));
    	String userInfo=userNameDisplay.getText();
    	
    	System.out.println("User Info  : "+userInfo);
    	System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());
    	
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
    	companyLogo.click();
    	
    	String getCompanyTxt=companyName.getText();
    	String getLoginCompanyName=getCompanyTxt.substring(0, 19);
    	System.out.println("company name  :  "+ getLoginCompanyName);
    	companyLogo.click();
    	
    	/*Thread.sleep(5000);
    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
    	
    	String getDashboard=dashboard.getText();
    	
    	System.out.println(getDashboard);*/
    	
    	String expuserInfo="SU";
    	String expLoginCompanyName="Automation Company ";
    	
    	System.out.println("userInfo             "+userInfo             +"  value expected  "+expuserInfo);
    	System.out.println("getLoginCompanyName  "+getLoginCompanyName  +"  value expected  "+expLoginCompanyName);
    	
    		if(userInfo.equalsIgnoreCase(expuserInfo)&&getLoginCompanyName.equalsIgnoreCase(expLoginCompanyName))
    		{
    			getLogger().info("Login User the Company Name and Information is displaying correct");
    			System.out.println("Pass : Login User and Company Name is displaying Correct");
    		
    			
    			return true;
    		/*
    			if(getDashboard.equalsIgnoreCase("Dashboard"))
    			{
    				excelReader.setCellData(xlfile, "Sheet1", 54, 9, resPass);
    				return true;
    				
    			}
    		    else
    		    {
    		    	excelReader.setCellData(xlfile, "Sheet1", 54, 9, resFail);
    		    	return false;
    		    	
    		    }*/
    		}
    		else
    		{
    			excelReader.setCellData(xlfile, "Sheet1", 54, 9, resFail);
    			getLogger().info("Fail : Login User the Company Name and Information is displaying Wrong");
    			System.out.println("Fail : Login Name the Company Name and Information is displaying Wrong");
    			return false;
    		}

    	
    }
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	public CompanyOptionsPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
}
