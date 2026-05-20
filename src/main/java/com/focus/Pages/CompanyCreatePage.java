package com.focus.Pages;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

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

import com.focus.Pages.LoginPage;
import com.focus.base.BaseEngine;
import com.focus.utilities.DriverUtility;

public class CompanyCreatePage extends BaseEngine{
	
	@FindBy(xpath="//i[@title='Create Company']")
    private static WebElement companyCreateBtn;
	
	@FindBy(xpath="//*[@id='CompanyName']")
	private static WebElement companyName;
	
	@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
	private static WebElement companyNameinCompanyLogo;
	
	@FindBy(xpath="//*[@id='id_footer']/strong[1]")
	public static WebElement Company_Name;
	
	@FindBy(xpath="//div[@id='errorCompanyName']")
	private static WebElement getCompanyNameError;
	
	@FindBy(xpath="//*[@id='CalendarType']")
	private static WebElement calendarType;
	
	@FindBy(xpath="//*[@class='icon-calendar datectrl-img datectrl-gap1']")
	private static WebElement datePicker;	
	
	@FindBy(xpath="//input[@id='CreatedDate']")
	private static WebElement accountingDate;
	
	@FindBy(xpath="//*[@id='CreatedDate_day_today']/td/span[2]")
	private static WebElement getDate;
	
	@FindBy(id="securityNormal")
	private static WebElement securityNormal;
	
	@FindBy(xpath="//*[@id='securityNormal']")
	private static WebElement securityNormal1;
	
	@FindBy(id="securityLow")
	private static WebElement securityLow;
	
    @FindBy(id="securityHigh")
    private static WebElement securityHigh;
	
	@FindBy(xpath="//input[@id='SUserPassword']")
    private static WebElement passwordSU;
	
	@FindBy(xpath=" //div[@id='errorUserPassword']")
	private static WebElement getPasswordErrorMsg;
	
	@FindBy(id="chkRemeberLoginCredentials")
	private static WebElement rememberLoginCredentials;
	
	@FindBy(id="chkUnauthorizedUsers")
	private static WebElement unAuthorizeUsers;
	
	@FindBy(id="Currency1")
	private static WebElement globalCurrency;
	
	@FindBy(id="Country")
	private static WebElement country;
	
	@FindBy(id="DefLanguages")
	private static WebElement defaultLanguage;
	
	@FindBy(xpath="//*[@id='companyLogo']")
	public static WebElement companyLogo;
	@FindBy(xpath="//*[@id='CompanyLogo']")
	private static WebElement CompanyLogo;
	
	@FindBy(xpath="//img[contains(@src,'data.image/png;base64,iVBOR')]")
	private static WebElement companylgo;
	
	@FindBy(xpath="/html/body/section/div[2]/header/nav/div/ul/li[6]/a/span[1]")
	private static WebElement userImage;

	@FindBy(xpath="//span[@class='icon icon-logout']")
	private static WebElement logoutOption;
	
	@FindBy(xpath="//*[@id='errmsgDiv']")
	private static WebElement mandatoryMsgs;
	@FindBy(xpath="//*[@id='company_Languages']//input[@value=0]")
	private static WebElement languagesEnglish;
	
	@FindBy(xpath="//*[@id='company_Languages']//input[@value=0]//following-sibling::span")
	private static WebElement languagesEnglishSelected;
	
	@FindBy(xpath="//*[@id='company_Languages']//input[@value=1]")
	private static WebElement languagesArabic;
	
	@FindBy(xpath="//*[@id='company_Languages']//input[@value=1]//following-sibling::span")
	private static WebElement languagesArabicSelected;
		
	@FindBy(xpath="//*[@id='company_Languages']//input[@value=2]")
	private static WebElement languagesPersian;
	
	@FindBy(xpath="//*[@id='company_Languages']//input[@value=2]//following-sibling::span")
	private static WebElement languagesPersianSelected;
	
	
	@FindBy(xpath="//label[@class='chk-margin Flabel']//input[@value='3']")
	private static WebElement languagesChinese;
	
	@FindBy(xpath="//label[@class='chk-margin Flabel']//input[@value='4']")
	private static WebElement languagesSpanish;
	
	@FindBy(xpath="//label[@class='chk-margin Flabel']//input[@value='5']")
	private static WebElement languagesPortuguese;
	
	@FindBy(xpath="//label[@class='chk-margin Flabel']//input[@value='6']")
	private static WebElement languagesIndonesian;
	
	@FindBy(xpath="//*[contains(@class,'icon-ok hiconright2')]")
	private static WebElement okButtonInCreateCompany;
	
	@FindBy(xpath="//span[contains(@class,'icon-close icon-font6')]")
	private static WebElement cancelButtonInCreateCompany;
	
	@FindBy(xpath="//i[@class='indicator pull-right icon-collepse icon-font6']")
	private static WebElement companyMinimizeAndMaximize;
	
	@FindBy(xpath="//*[@id='1']")
	private static WebElement homeMenu;
	
	@FindBy(xpath="//*[@id='2']/div/span")
	private static WebElement companyMenu;
	
	@FindBy(xpath="//*[@id='2']/span")
	private static WebElement companyMenuExpand;
	
	
	
	@FindBy(xpath="//*[@id='txtAdminUsername']")
	private static WebElement adminName;
	
	@FindBy(xpath="//*[@id='txtAdminPassword']")
	private static WebElement adminPassword;
	
//	@FindBy(xpath="/html/body/section/div[8]/div/div/div[4]/button[1]")
	@FindBy(xpath="//button[contains(text(),'Login')]")
	private static WebElement okButton;
	
	@FindBy(xpath="//*[@id='ShowAuthForCreateRestore']//div[4]/button[2]")
	private static WebElement cancelButton;
	
	@FindBy(xpath="//*[@id='CreatedDate_day_today']/td/span[1]")
	private static WebElement dateTodayBtn;
	
	private static final DateFormat sdf = new SimpleDateFormat("dd-MM-yyyy");
	
	 
	 @FindBy (xpath="//input[@id='txtSearchMenu_MainLayout']")
	 private static WebElement getSearch;
	 
	 
	 @FindBy (xpath="//*[@id='107']")
	 private static WebElement getTextFromSearch;
	 
	 @FindBy(xpath="//*[@id='dropdown03']/span")
		private static WebElement aboutIdQuickAccessMenu;
		
			@FindBy(xpath="//*[@id='quickAccessDiv']/ul/li[1]/a")
			private static WebElement favouriteOption;
			
			@FindBy(xpath="//*[@id='quickAccessDiv']/ul/li[2]/a")
			private static WebElement recentItemsOpion;
			
			@FindBy(xpath="//*[@id='quickAccessDiv']/ul/li[3]/a")
			private static WebElement runOption;
			
			@FindBy(xpath="//*[@id='quickAccessDiv']/ul/li[4]/a/span")
			private static WebElement quickAccessOption;
			
			@FindBy(xpath="//*[@id='nav-About-tab']")
			private static WebElement aboutOption;
			
				@FindBy(xpath="//*[@id='about']/div[1]/img")
				private static WebElement aboutImage;
				
				@FindBy(xpath="//*[@id='idAboutCompany']/div[1]/label[2]")
				private static WebElement aboutVersioin;

				@FindBy(xpath="//*[@id='idAboutCompany']/div[2]/label[2]")
				private static WebElement aboutRelease;

				@FindBy(xpath="//*[@id='idAboutCompany']/div[3]/label[2]")
				private static WebElement aboutLicensed;
				
				private static String getReleaseDateFromAbout;
				private static String modifyGetReleaseDateFromAbout;
				
				 @FindBy (xpath="//*[@id='CompanyName']")
				 private static WebElement getCompanyNameTextInEditScreen;
				 
				 @FindBy(xpath="//*[@class='icon-ok hiconright2']")
				 private static WebElement okButtonInEditCompanyScreen;
				 
				 @FindBy (xpath="//*[@id='CreatedDate']")
				 private static WebElement companyDateInEditCompany;
				 
				 @FindBy(xpath="//*[@id='btnMasterSaveClick']")
					public static WebElement saveBtn;
					
				 
					@FindBy(xpath="(//i[@class='icon-close hiconright2'])[1]")
					public static WebElement masterCloseBtn;
					
					@FindBy(xpath="//*[@id='iProductType']")
					public static WebElement productTypeSelect;
					
					@FindBy(xpath="//*[@class='icon-pick']")
					private static WebElement pickBtn;

					@FindBy(xpath="//input[@id='txtNewReference']")
					private static WebElement newReferenceTxt;

					@FindBy(xpath="//*[@class='icon-ok']")
					private static WebElement Bill_OkBtn;
					
	
	 
public static boolean checkLoginScreen() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
{
	
	  Thread.sleep(3000);

      getDriver().navigate().refresh();

      LoginPage lp=new LoginPage(getDriver()); 

    //  lp.checkLoginPageTitleByURLInputInBrowser(DriverUtility.SANITYUrl);
      
      getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(companyCreateBtn));
      boolean actcompanyCreateBtnDisplayed=companyCreateBtn.isDisplayed();
      
      if(actcompanyCreateBtnDisplayed)
      {
    	  return true;
      }
      else
      {
    	  return false;
      }
      

}



	public static boolean checkCreateCompanyInLoginPage() throws InterruptedException, IOException 
	 {
		 	click(companyCreateBtn);
		 	Thread.sleep(2000);
	    	click(adminName);
	    	adminName.sendKeys("admin");
	    	
	    	click(adminPassword);
	  	   	adminPassword.sendKeys("focus");
	  		
	  		click(okButton);
		
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(companyName));
			boolean actCompanyFields = companyName.isDisplayed() ;
			
			 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyName));
			 companyName.sendKeys("New Company");
			 Thread.sleep(2000);
			 
			 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordSU));
			 passwordSU.click();
			 passwordSU.sendKeys("su");
			 
			 //Finding the Count of the Calender Type
			 Select oSelect = new Select(calendarType);
			 List <WebElement> elementCount = oSelect.getOptions();
			
			 int iSize = elementCount.size();
			 System.out.println("Calender Count :"+iSize);
			 
			 //Finding the Georgian in the Calender Type
			 String Strq= calendarType.getText();
				
			 boolean getCalenderValue=false;
			 String[] str= Strq.split("\n");     
		     
			 for(String st1 : str)
		     {
		    	  st1.equals("Georgian");
		    	  getCalenderValue=true;
		    	  break;
		     }		
			 
			// System.out.println("Login Crednentials on Create Company Calender Value Actual   : " + iSize +  " Value Expected : " + 4);
			 System.out.println("Login Crednentials on Create Company Calender Value Actual   : " + getCalenderValue +  " Value Expected : " + true);
	
			 datePicker.click();
			 Thread.sleep(2000);
				
				String getCalenderPopupDate=getDate.getText();

				String ConPopupDate=getCalenderPopupDate.replace("/", "-");
				
				Date date=new Date();
				
				String sysDate=sdf.format(date);
				
			    System.out.println("Login Crednentials on Create Company Date Value Actual   : " + ConPopupDate +  " Value Expected : " + sysDate);
			    
			    dateTodayBtn.click();
			 
			    String getActualCountryValues=country.getText();
				
				//System.out.println("Country Combo box Values are : "+ getActualCountryValues);
				
				 boolean getCountryValue=false;
				 
				 String[] str1= getActualCountryValues.split("\n");     
			      
				 for(String st1 : str1)
			     {
			    	
			    	  st1.equals("India");
			    	  getCountryValue=true;
			    	  break;
			     }
				 
				 System.out.println("Login Crednentials on Create Company Country Value Actual   : " + getCountryValue +  " Value Expected : " + true);
				 
				 String getActualGlobalCurrencyValues=globalCurrency.getText();
					
					//System.out.println("Global Currency Combo box Values are : "+ getActualGlobalCurrencyValues);
					
					 boolean getGlobalCurrencyValue=false;
					 String[] str2= getActualGlobalCurrencyValues.split("\n");    
				      
					 for(String st1 : str2)
				      {
				    	
				    	  st1.equals("Indian Rupees, Rupees");
				    	  getGlobalCurrencyValue=true;
				    	  break;
				      }
				      
					 System.out.println("Login Crednentials on Create Company Currency Value Actual   : " + getGlobalCurrencyValue +  " Value Expected : " + true);
				      
					 
					 Select oSelect1 = new Select(defaultLanguage);
					 List <WebElement> elementCount1 = oSelect1.getOptions();
					 
					 boolean actvalue=languagesEnglish.isSelected();
					 
					 languagesArabicSelected.click();
					 
					 languagesPersianSelected.click();
					 Thread.sleep(2000);
					 
					 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CompanyLogo));
			           
				     CompanyLogo.click();
				     
				     Thread.sleep(3000);
				     
				     Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\FocusCompanyLogo.exe");
				     
				     Thread.sleep(3000);
				     
				     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(okButtonInCreateCompany));
				     
				     okButtonInCreateCompany.click();
				         
				    // Thread.sleep(180000);
				     
				     new WebDriverWait(getDriver(), 500).until(ExpectedConditions.alertIsPresent());
				     String actAlertTxt=getAlert().getText();
				     String expAlertTxt="Company Created Successfully";
				     System.out.println("Actual Alert Text after click OK------------"		+	actAlertTxt);
				     System.out.println("Expected Alert Text after click OK----------"		+	expAlertTxt);
				     getAlert().accept();
				     Thread.sleep(2000);
				     
				     getDriver().navigate().refresh();
				     Thread.sleep(2000);
				     
				     
						/*
						 * getWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption
						 * )); logoutOption.click(); Thread.sleep(2000);
						 */

					 getWebDriverWait().until(ExpectedConditions.visibilityOf(username));
					 boolean actUsernameDislayed=username.isDisplayed();
					 
					 System.out.println("User name displayed"	+	actUsernameDislayed);
					 
				     if(actCompanyFields&&actAlertTxt.equalsIgnoreCase(expAlertTxt) && actUsernameDislayed)
				     {
				    	 return true;
				     }
				     else
				     {
				    	 return false;
				     }
				   
			      
			
	 }
	
	public static boolean checkLogintoCreatedCompany() throws InterruptedException
	{
		
		  LoginPage lp=new LoginPage(getDriver()); 
	       
	        String unamelt="su";

	        String pawslt="su";

	        lp.enterUserName(unamelt);
	        Thread.sleep(1000);
	        lp.enterPassword(pawslt);

	        Thread.sleep(2000);

	        String compname = "New Company";

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

	        Thread.sleep(10000);

	        String actUserInfo1=userNameDisplay.getText();

	        System.out.println("User Info  : "+actUserInfo1);

	        System.out.println("User Info Capture Text  :  "+userNameDisplay.getText());

	        Thread.sleep(2000);
	       /* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	        companyLogo.click();*/

	        String getCompanyTxt1=Company_Name.getText();
	        String getLoginCompanyName1=getCompanyTxt1.substring(0, 11);
	        System.out.println("company name  :  "+ getLoginCompanyName1);
	      //  companyLogo.click();

	        String expUserInfo1           ="SU";
	        String expLoginCompanyName1   ="New Company";

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
	 
	
	public static boolean checkEditCompanyAfterCompanyCreatedWithSecurityasNormal() throws InterruptedException
	{
		Thread.sleep(2000);
		aboutIdQuickAccessMenu.click();
		Thread.sleep(2000);
		aboutOption.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(aboutRelease));
		getReleaseDateFromAbout=aboutRelease.getText();
		modifyGetReleaseDateFromAbout=""+getReleaseDateFromAbout;
		aboutIdQuickAccessMenu.click();
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Edit Company");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(5000);
		
	/*	getSearch.click();
		getSearch.sendKeys("Edit Company");
		
		getTextFromSearch.click();*/
		
		//Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(getCompanyNameTextInEditScreen));
		getCompanyNameTextInEditScreen.sendKeys(modifyGetReleaseDateFromAbout);
		
		 getAction().moveToElement(getCompanyNameTextInEditScreen).sendKeys(Keys.TAB).perform();
	     companyDateInEditCompany.click();;
	     
	     Thread.sleep(2000);
	     
		
		/*getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(securityNormal1));
		boolean actSecurityNormal=securityNormal1.isSelected();
		System.out.println("Security normal	"	+	actSecurityNormal);*/
	     
	     getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(securityNormal1));
	     securityNormal1.click();
	     
	     Thread.sleep(1000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(okButtonInEditCompanyScreen));
	     okButtonInEditCompanyScreen.click();
	     
	     
	    new WebDriverWait(getDriver(), 300).until(ExpectedConditions.alertIsPresent());
	    String actAlertAfterEdit=getAlert().getText();
	    String expAlertAfterEdit="Company updated successfully";
	    
	    getAlert().accept();
	     Thread.sleep(2000);
	     
	     getDriver().navigate().refresh();
	     Thread.sleep(2000);
	     
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
	     userNameDisplayLogo.click();
	     
	   
		
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
	     logoutOption.click();

	     Thread.sleep(2000);
	     
	  System.out.println("Actual Alert Text After Edit		"		+	actAlertAfterEdit);
	  System.out.println("Expect Alert Text After Edit		"		+	expAlertAfterEdit);
	  if(actAlertAfterEdit.equalsIgnoreCase(expAlertAfterEdit))
	  {
		  return true;
	  }
	  else
	  {
		  return false;
	  }
	   

	}
	 
	 
	public  boolean checkSavingItemsinItemMaster() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
    {
		System.out.println("***********checkItemMasterInMasterMenu Method Excutes **********");
		
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
    	Thread.sleep(6000);
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterNewBtn));
		 masterNewBtn.click();
		 Thread.sleep(8000);
		 
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
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(productTypeSelect));
			Select s1=new Select(productTypeSelect);
		s1.selectByVisibleText("Raw material");
		productTypeSelect.sendKeys(Keys.TAB);
		 
		 Thread.sleep(2000);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 
		 Thread.sleep(6000);
		 
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
		 
		 //Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(productTypeSelect));
			//Select s1=new Select(productTypeSelect);
		s1.selectByVisibleText("Raw material");
		productTypeSelect.sendKeys(Keys.TAB);
		 
		 Thread.sleep(2000);
		 
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
		 saveBtn.click();
		 
		 Thread.sleep(6000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Item3");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("Item3");
		 codeTxt.sendKeys(Keys.TAB);
		 
		 //Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(productTypeSelect));
			//Select s1=new Select(productTypeSelect);
		s1.selectByVisibleText("Raw material");
		productTypeSelect.sendKeys(Keys.TAB);
		 
		 Thread.sleep(2000);
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
	
	
	
	
	@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[12]")
	private static List<WebElement> masterItemsList;
	
	
	@FindBy(xpath="//tbody[@id='LandingGridBody']/tr[18]/td")
	private static List<WebElement> masterItems18thRowList;
	
	@FindBy (xpath="//*[@id='8']/span")
	public static WebElement passwordPolicySubMenu;
	
	@FindBy (xpath="//*[@id='btnSave']")
	public static WebElement passwordPolicySaveBtn;
	
	public static boolean checkSavingPasswordPolicy() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));	
			securityMenu.click();
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicySubMenu));
			passwordPolicySubMenu.click();
			
			Thread.sleep(4000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicyAddBtn));
			passwordPolicyAddBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(policyName));
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(policyName));
			policyName.sendKeys("Test");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordLength));
			passwordLength.sendKeys("2");
			passwordLength.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(passwordPolicySaveBtn));
			
			passwordPolicySaveBtn.click();
			
			Thread.sleep(2000);
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
	
	
	
	public boolean CheckSavingProfilewithAlloptionsEnabled() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfile));
		createProfile.click();
	        
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileProfileNameCombo));
		createProfileProfileNameCombo.click();
		createProfileProfileNameCombo.sendKeys("profilepv");
		
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
	    
	    Thread.sleep(2000);
	    
	    System.out.println("***********************************************CheckSavingAllProfileSTwithAlloptionsEnabled************************************************");
	    
	    String expMessage="Profile saved Successfully";
		
	    String actMessage=checkValidationMessage(expMessage);
	    
	    
	    System.out.println("Saving Message of Profile : " +actMessage+" Value Expected : "+expMessage);
	    
	    Thread.sleep(3000);
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createProfileCloseIcon));
	    createProfileCloseIcon.click();
	    
	   if (actMessage.equalsIgnoreCase(expMessage)) 
	   {
			return true;
		} 
		else 
		{
			return false;
		}
	}
	
	
	@FindBy(xpath="//*[contains(text(),'profilepv')]")
	public static WebElement avaliableProfilePV ;
	
	public boolean checkSavingRoleWithAllFieldsinAllTabs() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		securityMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeSecurityCreateRoleMenu));
		homeSecurityCreateRoleMenu.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRoleNameCombo));
		createRoleRoleNameCombo.click();
		
		createRoleRoleNameCombo.sendKeys("Testing");
		createRoleRoleNameCombo.sendKeys(Keys.TAB);
		Thread.sleep(2000);
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRolePasswordPolicyDropdown));
		createRolePasswordPolicyDropdown.click();
		Select s=new Select(createRolePasswordPolicyDropdown);
		s.selectByVisibleText("Test");
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(avaliableProfilePV));
		avaliableProfilePV.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleMOveFiledsFromLeftSideToRightSide));
		createRoleMOveFiledsFromLeftSideToRightSide.click();
		 
		Thread.sleep(2000);
		
		 int actAssignedProfilesListCount=assignedProfileList.size();
		 
		 Set<String> AssignedProfilesList=new HashSet<String>();
		  
		  for (int i = 0; i < actAssignedProfilesListCount; i++) 
		  {
			  String profile=assignedProfileList.get(i).getText();
			  Thread.sleep(2000);
			  AssignedProfilesList.add(profile);
		  }
		 
		  String actAssignedProfilesList = AssignedProfilesList.toString();
		  
		  String expAssignedProfilesList = "[profilepv]";
		  
		  System.out.println("Act assigned profile" + actAssignedProfilesList);
		  System.out.println("exp assigned profile" + expAssignedProfilesList);
		
		  //Additions Tab
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(additionTab));
		  additionTab.click();
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleAddTabFinancialMenu));
		  createRoleAddTabFinancialMenu.click();
		  
		  Thread.sleep(2000);
		  
		  int actFinancialMenusCount=additionTabFinancialMenuList.size();
		  	 
		  Set<String> FinancialMenusList=new HashSet<String>();
		  
		  for (int i = 0; i < actFinancialMenusCount; i++) 
		  {
			  String profile=additionTabFinancialMenuList.get(i).getText();
			  
			  FinancialMenusList.add(profile);
		  }
		  
		  String actFinancialMenusList = FinancialMenusList.toString();
		  
		  String expFinancialMenusList = "[Transactions, Final Accounts, Credit Management, Reports, Receivable and Payable Analysis, Budget, Online Payments]";
		 
		  
		  System.out.println(actFinancialMenusList);
		  System.out.println(expFinancialMenusList);
		  
		  
		  for (int i = 0; i < actFinancialMenusCount; i++) 
		  {
			  String text=additionTabFinancialMenuList.get(i).getText();
			  
			  if (text.equalsIgnoreCase("Transactions"))
			  {
				  additionTabFinancialMenuList.get(i).click();
			  }
		  }
		  
		  Thread.sleep(2000);
		  
		  int actTransactionsMenusCount=additionTabFinancialTransactionMenusList.size();
		  
		  Set<String> TransactionsMenusList=new HashSet<String>();
		  
		  for (int i = 0; i < actTransactionsMenusCount; i++) 
		  {
			  String profile=additionTabFinancialTransactionMenusList.get(i).getText();
			  
			  TransactionsMenusList.add(profile);
		  }
		  
		  String actTransactionsMenusList = TransactionsMenusList.toString();
		  
		  String expTransactionsMenusList = "[Cash and Bank, Sales, Journals, Purchases, Auto Postings]";
		 
		  System.out.println(actTransactionsMenusList);
		  System.out.println(expTransactionsMenusList);

		  
		  for (int i = 0; i < actTransactionsMenusCount; i++) 
		  {
			  String text=additionTabFinancialTransactionMenusList.get(i).getText();
			  
			  if (text.equalsIgnoreCase("Purchases")) 
			  {
				  additionTabFinancialTransactionMenusList.get(i).click();
			  }
		  }
		  
		  Thread.sleep(2000);
		  
		  int actPurchasesMenusCount=additionTabFinancialTransactionsPurchasesMenusList.size();
		 
		  Set<String> PurchasesMenusList=new HashSet<String>();
		  
		  for (int i = 0; i < actPurchasesMenusCount; i++) 
		  {
			  String profile=additionTabFinancialTransactionsPurchasesMenusList.get(i).getText();
			  
			  PurchasesMenusList.add(profile);
		  }
		  
		  String actPurchasesMenusList = PurchasesMenusList.toString();
		  
		  String expPurchasesMenusList = "[Purchase Returns, Purchase Vouchers]";
		  
		
		  System.out.println(actPurchasesMenusList);
		  System.out.println(expPurchasesMenusList);
		  

		  
		  for (int i = 0; i < actPurchasesMenusCount; i++) 
		  {
			  String text=additionTabFinancialTransactionsPurchasesMenusList.get(i).getText();
			  
			  if (text.equalsIgnoreCase("Purchase Vouchers")) 
			  {
				  additionTabFinancialTransactionsPurchasesMenusList.get(i).click();
			  }
		  }
		
		  Thread.sleep(2000);


		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleAddTabSelectAllChkBox));
		  createRoleAddTabSelectAllChkBox.click();
		  
		  Thread.sleep(2000);
		  
		
		  //Exclusions Tab
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(exclusionsTab));
		  exclusionsTab.click();
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleExcluTabInventoryMenu));
		  createRoleExcluTabInventoryMenu.click();
		  
		  Thread.sleep(2000);
		  
		  int actInventoryMenusCount=NewExcluTabInventoryMenusList.size();
		  
		  Set<String> InventoryMenusList=new HashSet<String>();
		  
		  for (int i = 0; i < actInventoryMenusCount; i++) 
		  {
			  String menu=NewExcluTabInventoryMenusList.get(i).getText();
			  
			  InventoryMenusList.add(menu);
		  }

		  String actInventoryMenusList = InventoryMenusList.toString();
		  
		  String expInventoryMenusList = "[Transactions, Order Management, Reports]";
		  
		 
		  System.out.println(actInventoryMenusList);
		  System.out.println(expInventoryMenusList);

		  
		  for (int i = 0; i < actInventoryMenusCount; i++) 
		  {
			  String menu=NewExcluTabInventoryMenusList.get(i).getText();
			  Thread.sleep(1000);
			  
			  if (menu.equalsIgnoreCase("Transactions")) 
			  {
				  Thread.sleep(1000);
				  NewExcluTabInventoryMenusList.get(i).click();
				  Thread.sleep(1000);
			  }
		  }
		  
		  
		  Thread.sleep(2000);
		  
		  ClickUsingJs(InvTransaction_expand);
		  
		  Thread.sleep(2000);
		  
		  int actInventoryTransactionsMenusCount=excluTabInventoryTransactionsMenusList.size();
		  
		  Set<String> InventoryTransactionsMenusList=new HashSet<String>();
		  
		  for (int i = 0; i < actInventoryTransactionsMenusCount; i++) 
		  {
			  String menu=excluTabInventoryTransactionsMenusList.get(i).getText();
			  
			  InventoryTransactionsMenusList.add(menu);
		  }

		  String actInventoryTransactionsMenusList = InventoryTransactionsMenusList.toString();
		  
		  String expInventoryTransactionsMenusList = "[Sales, Stocks, Hold/Unhold Stock, Stock Reconciliation, Purchases]";
		  
		  
		  System.out.println(actInventoryTransactionsMenusList);
		  System.out.println(expInventoryTransactionsMenusList);
		  
		 
		  for (int i = 0; i < actInventoryTransactionsMenusCount; i++) 
		  {
			  String menu=excluTabInventoryTransactionsMenusList.get(i).getText();
			  
			  if (menu.equalsIgnoreCase("Stock Reconciliation")) 
			  {
				  excluTabInventoryTransactionsMenusList.get(i).click();
			  }
		  }
		  
		  Thread.sleep(2000);
		  
		
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleExclTabUnSelectAllChkBox));
		  createRoleExclTabUnSelectAllChkBox.click();

		
		//Restriction For Entry Tab
		  
		  Thread.sleep(2000);
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleRestrictionForEntryTab));
		  createRoleRestrictionForEntryTab.click();
		  
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
		  
		  
		

		  for (int i = 0; i < actRestrictionForEntryMastersListCount; i++) 
		  {
			  String masters=createRoleRestrictionForEntryTabMastersList.get(i).getText();
			  
			  if (masters.equalsIgnoreCase("Account")) 
			  {
				  createRoleRestrictionForEntryTabMastersList.get(i).click();  
			  }
		  }
		  
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableSelectRow2Col1));
		  restrictionForEntryTabTableSelectRow2Col1.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionForEntryTabTableEnterMasterTxt));
		  String c  = "Customer B";
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
		  

		  
		  //Restriction For Trees
		  Thread.sleep(2000);
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restrictionforTreesTab));
		  restrictionforTreesTab.click();
		  
		  Thread.sleep(2000);
		  
		  int actRestrictionForTreesMastersListCount=createRoleRestrictionForTreesTabMastersList.size();

		  for (int i = 0; i < actRestrictionForTreesMastersListCount; i++) 
		  {
			  String masters=createRoleRestrictionForTreesTabMastersList.get(i).getText();
			  
			  if (masters.equalsIgnoreCase("Account")) 
			  {
				  createRoleRestrictionForTreesTabMastersList.get(i).click();  
			  }
		  }
		  
		 
		  restrictionForTreesTabDefaultChkBox.click();
		  
		  // Transaction Rights Tab
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionRightsTab));
		  transactionRightsTab.click();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(allowBillWiseOnAccountChkBox));
		  allowBillWiseOnAccountChkBox.click();
		  
		  Select budgetLimitWarning= new Select(budgetLimitWarningDropdown);
		  budgetLimitWarning.selectByVisibleText("Allow");
		  
		  Select creditLimitWarning= new Select(creditLimitWarningDropdown);
		  creditLimitWarning.selectByVisibleText("Alert");
		  
		  Select negativeCashCheck= new Select(negativeCashCheckDropdown);
		  negativeCashCheck.selectByVisibleText("Allow");

		  Select negativeStockCheck= new Select(negativeStockCheckDropdown);
		  negativeStockCheck.selectByVisibleText("Allow");
		  
		  cantaddfutureTransChkbox.click();
		  
		  cannotAddTransThatAreMoreThanTxt.click();
		  cannotAddTransThatAreMoreThanTxt.clear();
		  cannotAddTransThatAreMoreThanTxt.sendKeys("5");
		  
		  cannotEditPreviousMonthEntriesAfterTxt.click();
		  cannotEditPreviousMonthEntriesAfterTxt.clear();
		  cannotEditPreviousMonthEntriesAfterTxt.sendKeys("15");
		  
		  CantEditTransthatAreremorethanTxt.click();
		  CantEditTransthatAreremorethanTxt.clear();
		  CantEditTransthatAreremorethanTxt.sendKeys("50");
		  
		  CantAddPrevMnthEntriesaftTxt.click();
		  CantAddPrevMnthEntriesaftTxt.clear();
		  CantAddPrevMnthEntriesaftTxt.sendKeys("1");
		  
		  CantPrintAfterValueTxt.click();
		  CantPrintAfterValueTxt.clear();
		  CantPrintAfterValueTxt.sendKeys("12");
		  
		  Select CantPrintAfterValue= new Select(cantPrintAfterValueDropdown);
		  CantPrintAfterValue.selectByVisibleText("Hours");
		  
		  CantRePrintAfterValueTxt.click();
		  CantRePrintAfterValueTxt.clear();
		  CantRePrintAfterValueTxt.sendKeys("1");
		  
		  Select CantRePrintAfterValue= new Select(CantRePrintAfterDropdown);
		  CantRePrintAfterValue.selectByVisibleText("Days");
		  
		  doNotAllowMasterCustomization.click();
		  
		  /*
		  
		  //AI Rights
		  Thread.sleep(2000);
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(AIRightsTab));
		  AIRightsTab.click();
		  
		  Thread.sleep(2000);
		  
		  AIRightsEnableVUIChkBox.click();
		  
		  AIRightsEnableAIChkBox.click();
		  
		  NoOfVoiceCommandsPermittedTxt.click();
		  NoOfVoiceCommandsPermittedTxt.clear();
		  NoOfVoiceCommandsPermittedTxt.sendKeys("3");
		  
		  NoOfTextCommandsPermittedTxt.click();
		  NoOfTextCommandsPermittedTxt.clear();
		  NoOfTextCommandsPermittedTxt.sendKeys("5");
		  
		  AIRightsSalesChkBox.click();
		  AIRightsSalesOrderChkBox.click();
		  AIRightsStockChkBox.click();
		  AIRightsPurchasesChkBox.click();
		  AIRightsPurchasesOrdersChkBox.click();
		  Thread.sleep(1000);*/
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleSaveicon));
		  createRoleSaveicon.click();

		  String expMessage="Role saved Successfully";
		  
		  String actMessage=checkValidationMessage(expMessage);
		  
		 
		 /* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createRoleCloseIcon));
		  createRoleCloseIcon.click();*/
		  
		  System.out.println("actAssignedProfilesList : "+actAssignedProfilesList);
		  System.out.println("expAssignedProfilesList : "+expAssignedProfilesList);
		  
		
		  
		  System.out.println("actFinancialMenusList : "+actFinancialMenusList);
		  System.out.println("expFinancialMenusList : "+expFinancialMenusList);
		  
		  System.out.println("actTransactionsMenusList : "+actTransactionsMenusList);
		  System.out.println("expTransactionsMenusList : "+expTransactionsMenusList);
		  
		  System.out.println("actPurchasesMenusList : "+actPurchasesMenusList);
		  System.out.println("expPurchasesMenusList : "+expPurchasesMenusList);
		  
		  System.out.println("actInventoryMenusList : "+actInventoryMenusList);
		  System.out.println("expInventoryMenusList : "+expInventoryMenusList);
		  
		  System.out.println("actInventoryTransactionsMenusList : "+actInventoryTransactionsMenusList);
		  System.out.println("expInventoryTransactionsMenusList : "+expInventoryTransactionsMenusList);
		  
		  
		  System.out.println("Error Message : "+actMessage+"  value expected  "+expMessage);

		if (actAssignedProfilesList.equalsIgnoreCase(expAssignedProfilesList) &&   actFinancialMenusList.equalsIgnoreCase(expFinancialMenusList) &&
			 actTransactionsMenusList.equalsIgnoreCase(expTransactionsMenusList) && actPurchasesMenusList.equalsIgnoreCase(expPurchasesMenusList) && 
			 actInventoryMenusList.equalsIgnoreCase(expInventoryMenusList) && actInventoryTransactionsMenusList.equalsIgnoreCase(expInventoryTransactionsMenusList)&& actMessage.equalsIgnoreCase(expMessage)) 
	   {
		  
			return true;
		} 
		else 
		{
			return false;
		}

	}
	
	
	public boolean checkInputMandatoryFieldsInCreateUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		 homeMenu.click();
		 Thread.sleep(2000);	
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(securityMenu));
		 securityMenu.click();
		 Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserMenu));
		createUserMenu.click();
		
		Thread.sleep(5000);
		
	    
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserIcon));
		 createUserIcon.click();
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginNameTxt));
		 createUserLoginNameTxt.click();
		 createUserLoginNameTxt.sendKeys("User1");

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserERPRoleDropDown));
		 createUserERPRoleDropDown.click();
		 createUserERPRoleDropDown.sendKeys("Testing");
		 
	     int count = createUserERPRolesDropdownOptions.size()-1;
		 System.out.println("Count "+ count);
	    
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserPasswordTxt));
		 createUserPasswordTxt.sendKeys("12345");

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserConfirmPasswordTxt));
		 createUserConfirmPasswordTxt.sendKeys("12345");
		 	
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserNameTxt));
		 createUserNameTxt.sendKeys("UserCreationTest");
		 	
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserLoginAbbrivationTxt));
		 createUserLoginAbbrivationTxt.sendKeys("UCT");
		 
		 createUserSecurityQuestionDropDown.click();
		 Select s = new Select(createUserSecurityQuestionDropDown);
		 s.selectByVisibleText("In which country were you born?");

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
			createUserEmailTxt.sendKeys("testing@focussoftnet.com");

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(createUserSaveIcon));
		 createUserSaveIcon.click();
		 
		 String expMessage = "new user created successfully";
		 String actMessage = checkValidationMessage(expMessage);
		 
		 System.out.println(actMessage);
		 System.out.println(expMessage);
		 
	     if (actMessage.equalsIgnoreCase(expMessage)) 
	     {
	    	 return true;
	 	 } 
	 	 else 
	 	 {
			return false;
	 	 }
		
	}


	
	public boolean checkSavingVoucherinPurchseVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(2000);
	
	click(financialsMenu);
	
	 Thread.sleep(2000);
	
	 click(financialsTransactionMenu);
	 Thread.sleep(2000);
	 
	 click(financialsTransactionsPurchaseMenu);
	 Thread.sleep(2000);
	 click(purchasesVoucherBtn1);
	 Thread.sleep(10000);
	 
	 click(newBtn);
	 Thread.sleep(4000);
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
		purchaseAccountTxt.sendKeys("Purchase");
		Thread.sleep(2000);
		purchaseAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		click(vendorAccountTxt);
		vendorAccountTxt.sendKeys("Vendor A");
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		click(select1stRow_1stColumn);
		click(enter_ItemTxt);
		enter_ItemTxt.sendKeys("Item1");
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);
		
		//click(select1stRow_4thColumn);
		click(enter_Quantity);
		enter_Quantity.sendKeys("5");
		enter_Quantity.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		//click(select1stRow_6thColumn);
		click(enter_Rate);
		enter_Rate.clear();
		enter_Rate.sendKeys("5");
		Thread.sleep(2000);
		
		click(voucherSaveBtn);
		Thread.sleep(10000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/
		
		String expMessage="Saving in background.";
		String actMessage=checkValidationMessage(expMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		
		 
		  
		  if(actMessage.equalsIgnoreCase(expMessage) ) 
		  { 
			  return true; 
			 } 
		  else 
		  { 
			  return false;
		  }
	
	}
	
	
	public static boolean checkEditingVoucher1inPurchaseVouchers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("1");
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);
		Thread.sleep(10000);
		Thread.sleep(10000);
		
		click(vendorAccountTxt);
		
		vendorAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		click(select1stRow_2ndColumn);
		
		click(enter_Quantity);
		enter_Quantity.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_Quantity.sendKeys("10");
		enter_Quantity.sendKeys(Keys.TAB);
		
		click(voucherSaveBtn);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		Thread.sleep(2000);
		
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 1");
		expMessage.add("Saving in background.");*/
		
		String expMessage="Saving in background.";
		String actMessage=checkValidationMessage(expMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
	
		
		
		click(new_CloseBtn);
		
		  if(actMessage.equalsIgnoreCase(expMessage) ) 
		  { 
			  return true; 
			 } 
		  else 
		  { 
			  return false;
		  }
	
		
	}
	
	@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr/td[5]")
	public static List<WebElement> grid_VoucherNoList;
	
	
	@FindBy(xpath = "//tbody[@id='tblBodyTransRender']/tr/td[2]")
	public static List<WebElement> grid_VoucherChckBoxList;
	
	public static boolean checkSuspendtheVoucherinPurchaseVouchers() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int vcount = grid_VoucherNoList.size();

		for (int i = 0; i < vcount; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase("1")) 
			{
				getAction().click(grid_VoucherChckBoxList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		click(suspendBtn);
		
		String expMessage="Voucher Suspended successfully";
	
		String actMessge=checkValidationMessage(expMessage);
		
		
		
		if(actMessge.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
		
}	
		
	
	public static boolean checkLogoutAndLoginWithNewlySavedUser() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  Thread.sleep(2000);
		  
		  getDriver().navigate().refresh();
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		  userNameDisplayLogo.click();
		  
		  Thread.sleep(3000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		  logoutOption.click();
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		  username.sendKeys("User1");
		  getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		  Thread.sleep(2000);
		  
	      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
	      password.sendKeys("12345");
	      getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
	      Thread.sleep(2000);
	      
	      String compname = "New Company";

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

	      
	      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
	      signIn.click();
	      
	      Thread.sleep(5000);
	      
	      //getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameUser1Display));
	      //userNameUser1Display.click();
	             	
	  	  String actUserInfo = userNameUser1Display.getText();
	  	  String expUserInfo = "User1";

	      Thread.sleep(2000);
		
		  String getCompanyTxt=Company_Name.getText();
		  String actLoginCompanyName = getCompanyTxt.substring(0, 11);
		  String expLoginCompanyName = "New Company";
		 // companyLogo.click();
		  Thread.sleep(2000);
		  
	      System.out.println("UserInfo               : ."+actUserInfo          +".  value expected  ."+expUserInfo+".");
	      System.out.println("LoginCompanyName       : ."+actLoginCompanyName  +".  value expected  ."+expLoginCompanyName+".");
			
		  if(actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName))  
	      {
			return true;
		  }
	      else
	      {
			return false;
		  }
		}	
	
	
	public static boolean checkEditingVoucherinPurchaseVouchersatUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		Thread.sleep(2000);
		
		click(financialsMenu);
		
		 Thread.sleep(2000);
		
		 click(financialsTransactionMenu);
		 Thread.sleep(2000);
		 
		 click(financialsTransactionsPurchaseMenu);
		 Thread.sleep(2000);
		 click(purchasesVoucherBtn1);
		 Thread.sleep(2000);
		 
		 click(newBtn);
		 Thread.sleep(4000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			documentNumberTxt.sendKeys("1");
			Thread.sleep(2000);
			documentNumberTxt.sendKeys(Keys.TAB);
			Thread.sleep(8000);
			
			click(vendorAccountTxt);
			
			vendorAccountTxt.sendKeys(Keys.TAB);
			Thread.sleep(2000);
			
			
			click(select1stRow_1stColumn);
			click(enter_ItemTxt);
			enter_ItemTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			enter_ItemTxt.sendKeys("Item2");
			Thread.sleep(2000);
			enter_ItemTxt.sendKeys(Keys.TAB);
			
			click(voucherSaveBtn);
			Thread.sleep(2000);
			
			/*
			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * pickBtn)); pickBtn.click();
			 * 
			 * Thread.sleep(2000);
			 * 
			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * Bill_OkBtn)); Bill_OkBtn.click();
			 */
			
		/*	
			HashSet<String> actMessage = new HashSet();

			for (int i = 0; i < 2; i++) 
			{
				String data = checkValidationMessage("");
				actMessage.add(data);
			}

			HashSet<String> expMessage = new HashSet();

			expMessage.add("Voucher saved successfully : 1");
			expMessage.add("Saving in background.");*/
			

			String expMessage="Saving in background.";
			String actMessage=checkValidationMessage(expMessage);


			System.out.println("Actual Message    : "+actMessage);
			System.out.println("Expected Message  : "+expMessage);
			
			
			//click(new_CloseBtn);
			
			  if(actMessage.equalsIgnoreCase(expMessage) ) 
			  { 
				  return true; 
				 } 
			  else 
			  { 
				  return false;
			  }
		
		 
		 
	}
	 
	
	public static boolean checkSupendtheVoucherinPurchaseVouchersfromEntryPageatUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			documentNumberTxt.sendKeys("1");
			Thread.sleep(2000);
			documentNumberTxt.sendKeys(Keys.TAB);
			Thread.sleep(5000);
			checkValidationMessage("");
			click(new_SuspendBtn);
			Thread.sleep(4000);
			String expMessage="Voucher saved successfully";
			
			String actMessge=checkValidationMessage(expMessage);
			
			
			
			if(actMessge.startsWith(expMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
			
	}
	
	
	public boolean checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		  userNameDisplayLogo.click();
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		  logoutOption.click();
		  
		  Thread.sleep(2000);
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
		  username.sendKeys("su");
		  getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		  Thread.sleep(2000);
		  
	      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(password));
	      password.sendKeys("su");
	      getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
	      Thread.sleep(2000);
	      
	      String compname = "New Company";

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

	      
	      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(signIn));
	      signIn.click();
	      
	      Thread.sleep(5000);
	      
	      //getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
	      //userNameDisplay.click();
	             	
	  	  String actUserInfo = userNameDisplay.getText();
	  	  String expUserInfo = "SU";

	      Thread.sleep(2000);
		
		  String getCompanyTxt=Company_Name.getText();
		  String actLoginCompanyName = getCompanyTxt.substring(0, 11);
		  String expLoginCompanyName = "New Company";
		//  companyLogo.click();
		  Thread.sleep(2000);
		  
	      System.out.println("UserInfo               : ."+actUserInfo          +".  value expected  ."+expUserInfo+".");
	      System.out.println("LoginCompanyName       : ."+actLoginCompanyName  +".  value expected  ."+expLoginCompanyName+".");
			
		  if(actUserInfo.equalsIgnoreCase(expUserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName))  
	      {
			return true;
		  }
	      else
	      {
			return false;
		  }
		}	
	
	
	
	public boolean checkEditCompanyWithSecurityasLow() throws InterruptedException
	{
		

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Edit Company");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(5000);
		
		
	     
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(securityLow));
		securityLow.click();
		/*String  actSecurityLow=securityLow.getAttribute("class");    
		String expSecurityLow="security-lavel-radio radio1 theme_background-color title-top security-lavel-radio-hover security-lavel-radio-selected";
		System.out.println("Security Low	"	+	actSecurityLow);*/
		Thread.sleep(1000);
		 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(okButtonInEditCompanyScreen));
	     okButtonInEditCompanyScreen.click();
	     
	     
	    new WebDriverWait(getDriver(), 300).until(ExpectedConditions.alertIsPresent());
	    String actAlertAfterEdit=getAlert().getText();
	    String expAlertAfterEdit="Company updated successfully";
	    
	    getAlert().accept();
	    
	    Thread.sleep(2000);
	    getDriver().navigate().refresh();
	     Thread.sleep(2000);
	     
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
	     userNameDisplayLogo.click();
	     
	   
		
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
	     logoutOption.click();

	     Thread.sleep(2000);
	  System.out.println("Actual Alert Text After Edit		"		+	actAlertAfterEdit);
	  System.out.println("Expect Alert Text After Edit		"		+	expAlertAfterEdit);
	  if(actAlertAfterEdit.equalsIgnoreCase(expAlertAfterEdit))
	  {
		  return true;
	  }
	  else
	  {
		  return false;
	  }

	}
	
	
	public boolean checkSavingVoucher2inPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		

		Thread.sleep(2000);
	
	click(financialsMenu);
	
	 Thread.sleep(2000);
	
	 click(financialsTransactionMenu);
	 Thread.sleep(2000);
	 
	 click(financialsTransactionsPurchaseMenu);
	 Thread.sleep(2000);
	 click(purchasesVoucherBtn1);
	 Thread.sleep(2000);
	 
	 click(newBtn);
	 Thread.sleep(4000);
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
		purchaseAccountTxt.sendKeys("Purchase");
		Thread.sleep(2000);
		purchaseAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		click(vendorAccountTxt);
		vendorAccountTxt.sendKeys("Vendor A");
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		click(select1stRow_1stColumn);
		click(enter_ItemTxt);
		enter_ItemTxt.sendKeys("Item2");
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);
		
		//click(select1stRow_4thColumn);
		click(enter_Quantity);
		enter_Quantity.sendKeys("5");
		enter_Quantity.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		//click(select1stRow_6thColumn);
		click(enter_Rate);
		enter_Rate.clear();
		enter_Rate.sendKeys("5");
		Thread.sleep(2000);
		
		click(select2ndRow_1stColumn);
		click(enter_ItemTxt);
		enter_ItemTxt.sendKeys("Item2");
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);
		
		//click(select1stRow_4thColumn);
		click(enter_Quantity);
		enter_Quantity.sendKeys("8");
		enter_Quantity.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		//click(select1stRow_6thColumn);
		click(enter_Rate);
		enter_Rate.clear();
		enter_Rate.sendKeys("6.25");
		Thread.sleep(2000);
		
		
		click(voucherSaveBtn);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		Thread.sleep(2000);
		/*
		HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 2");
		expMessage.add("Saving in background.");*/
		

		String expMessage="Saving in background.";
		String actMessage=checkValidationMessage(expMessage);


		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		
		 
		  
		  if(actMessage.equalsIgnoreCase(expMessage) ) 
		  { 
			  return true; 
			 } 
		  else 
		  { 
			  return false;
		  }
	
	
	}
	
	
	public boolean checkEditingandSavingVoucher2inPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	
	{
		

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("2");
		Thread.sleep(4000);
		documentNumberTxt.sendKeys(Keys.TAB);
		Thread.sleep(12000);
		
		click(vendorAccountTxt);
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		click(select1stRow_2ndColumn);
		
		click(enter_Quantity);
		enter_Quantity.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_Quantity.sendKeys("25");
		enter_Quantity.sendKeys(Keys.TAB);
		
		click(select2ndRow_2ndColumn);
		
		click(enter_Quantity);
		enter_Quantity.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_Quantity.sendKeys("10");
		enter_Quantity.sendKeys(Keys.TAB);
		
		click(voucherSaveBtn);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		Thread.sleep(2000);
	/*	
		HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 2");
		expMessage.add("Saving in background.");*/
		

		String expMessage="Saving in background.";
		String actMessage=checkValidationMessage(expMessage);


		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		
		//click(new_CloseBtn);
		
		  if(actMessage.equalsIgnoreCase(expMessage) ) 
		  { 
			  return true; 
			 } 
		  else 
		  { 
			  return false;
		  }
	
		
	
	}
	
	
	
	public boolean checkSuspendVoucher2FromEntryPageinPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			documentNumberTxt.sendKeys("2");
			Thread.sleep(2000);
			documentNumberTxt.sendKeys(Keys.TAB);
			Thread.sleep(8000);
			checkValidationMessage("");
			
			click(new_SuspendBtn);
			Thread.sleep(4000);
			String expMessage="Voucher saved successfully";
			
			String actMessge=checkValidationMessage(expMessage);
			
			
			
			if(actMessge.startsWith(expMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
			
	
	}
	
	
	public boolean checkEditingVoucher2inPurchaseVoucheratUserLevel() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		
			Thread.sleep(2000);
			
			click(financialsMenu);
			
			 Thread.sleep(2000);
			
			 click(financialsTransactionMenu);
			 Thread.sleep(2000);
			 
			 click(financialsTransactionsPurchaseMenu);
			 Thread.sleep(2000);
			 click(purchasesVoucherBtn1);
			 Thread.sleep(2000);
			 
			 click(newBtn);
			 Thread.sleep(4000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
				documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
				documentNumberTxt.sendKeys("2");
				Thread.sleep(2000);
				documentNumberTxt.sendKeys(Keys.TAB);
				Thread.sleep(8000);
				
				click(vendorAccountTxt);
				
				vendorAccountTxt.sendKeys(Keys.TAB);
				Thread.sleep(2000);
				
				
				click(select1stRow_1stColumn);
				click(enter_ItemTxt);
				enter_ItemTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
				enter_ItemTxt.sendKeys("Item2");
				Thread.sleep(2000);
				enter_ItemTxt.sendKeys(Keys.TAB);
				
				click(voucherSaveBtn);
				Thread.sleep(2000);
				
				
				  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
				  pickBtn)); pickBtn.click();
				  
				  Thread.sleep(2000);
				  
				  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
				  Bill_OkBtn)); Bill_OkBtn.click();
				  Thread.sleep(2000);
				
				
			/*	HashSet<String> actMessage = new HashSet();

				for (int i = 0; i < 2; i++) 
				{
					String data = checkValidationMessage("");
					actMessage.add(data);
				}

				HashSet<String> expMessage = new HashSet();

				expMessage.add("Voucher saved successfully : 2");
				expMessage.add("Saving in background.");
*/
				  

					String expMessage="Saving in background.";
					String actMessage=checkValidationMessage(expMessage);

				System.out.println("Actual Message    : "+actMessage);
				System.out.println("Expected Message  : "+expMessage);
				
				
				click(new_CloseBtn);
				
				  if(actMessage.equalsIgnoreCase(expMessage) ) 
				  { 
					  return true; 
					 } 
				  else 
				  { 
					  return false;
				  }
			
			 
			 
		}
		 
	
	
	public static boolean checkSuspendtheVoucher2inPurchaseVouchers() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int vcount = grid_VoucherNoList.size();

		for (int i = 0; i < vcount; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase("2")) 
			{
				getAction().click(grid_VoucherChckBoxList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		click(suspendBtn);
		
		String expMessage="Voucher Suspended Successfully";
	
		String actMessge=checkValidationMessage(expMessage);
		
		
		
		if(actMessge.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
		
}	
		
	
	
	public boolean checkEditCompanyWithSecurityasHigh() throws InterruptedException
	{
		

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Edit Company");
		Thread.sleep(2000);
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3500);
		
	     
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(securityHigh));
		securityHigh.click();
		Thread.sleep(1000);
		/*String  actSecurityHigh=securityHigh.getAttribute("class");    
		String expSecurityHigh="security-lavel-radio theme_background-color radio3 title-top security-lavel-radio-hover security-lavel-radio-selected";
		System.out.println("Actual Security High	"	+	actSecurityHigh);
		System.out.println("Expect Security High	"	+	expSecurityHigh);*/
		
		 getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(okButtonInEditCompanyScreen));
	     okButtonInEditCompanyScreen.click();
	     
	     
	    new WebDriverWait(getDriver(), 300).until(ExpectedConditions.alertIsPresent());
	    String actAlertAfterEdit=getAlert().getText();
	    String expAlertAfterEdit="Company updated successfully";
	    
	    getAlert().accept();
	    
	    Thread.sleep(2000);
	    
	    getDriver().navigate().refresh();
	     Thread.sleep(2000);
	     
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
	     userNameDisplayLogo.click();
	     
	   
		
	     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
	     logoutOption.click();

	     Thread.sleep(2000);
	     
	  System.out.println("Actual Alert Text After Edit		"		+	actAlertAfterEdit);
	  System.out.println("Expect Alert Text After Edit		"		+	expAlertAfterEdit);
	  if(actAlertAfterEdit.equalsIgnoreCase(expAlertAfterEdit) /*&& actSecurityHigh.equalsIgnoreCase(expSecurityHigh)*/)
	  {
		  return true;
	  }
	  else
	  {
		  return false;
	  }

	}
	
	
	public boolean checkSavingVoucher3inPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		

		Thread.sleep(2000);
	
	click(financialsMenu);
	
	 Thread.sleep(2000);
	
	 click(financialsTransactionMenu);
	 Thread.sleep(2000);
	 
	 click(financialsTransactionsPurchaseMenu);
	 Thread.sleep(2000);
	 click(purchasesVoucherBtn1);
	 Thread.sleep(2000);
	 
	 click(newBtn);
	 Thread.sleep(4000);
		
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
		purchaseAccountTxt.sendKeys("Purchase");
		Thread.sleep(2000);
		purchaseAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		click(vendorAccountTxt);
		vendorAccountTxt.sendKeys("Vendor A");
		Thread.sleep(2000);
		vendorAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		click(select1stRow_1stColumn);
		click(enter_ItemTxt);
		enter_ItemTxt.sendKeys("Item2");
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);
		
		//click(select1stRow_4thColumn);
		click(enter_Quantity);
		enter_Quantity.sendKeys("5");
		enter_Quantity.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		//click(select1stRow_6thColumn);
		click(enter_Rate);
		enter_Rate.clear();
		enter_Rate.sendKeys("5");
		Thread.sleep(2000);
		
		click(select2ndRow_1stColumn);
		click(enter_ItemTxt);
		enter_ItemTxt.sendKeys("Item2");
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);
		
		//click(select1stRow_4thColumn);
		click(enter_Quantity);
		enter_Quantity.sendKeys("8");
		enter_Quantity.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		//click(select1stRow_6thColumn);
		click(enter_Rate);
		enter_Rate.clear();
		enter_Rate.sendKeys("6.25");
		Thread.sleep(2000);
		
		click(select3rdRow_1stColumn);
		click(enter_ItemTxt);
		enter_ItemTxt.sendKeys("Item3");
		Thread.sleep(2000);
		enter_ItemTxt.sendKeys(Keys.TAB);
		
		//click(select1stRow_4thColumn);
		click(enter_Quantity);
		enter_Quantity.sendKeys("9");
		enter_Quantity.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		//click(select1stRow_6thColumn);
		click(enter_Rate);
		enter_Rate.clear();
		enter_Rate.sendKeys("5.22");
		Thread.sleep(2000);
		
		click(voucherSaveBtn);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		Thread.sleep(2000);
	/*	HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 3");
		expMessage.add("Saving in background.");
*/
		

		String expMessage="Saving in background.";
		String actMessage=checkValidationMessage(expMessage);

		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		
		 
		  
		  if(actMessage.equalsIgnoreCase(expMessage) ) 
		  { 
			  return true; 
			 } 
		  else 
		  { 
			  return false;
		  }
	
	
	}
	
	
	@FindBy(xpath="//*[@id='id_transaction_entry_login_popup_username']")
	public static WebElement loginValidation_UsernameTxt;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_login_popup_password']")
	public static WebElement loginValidation_PasswordTxt;
	
	@FindBy(xpath="//*[@id='id_transaction_login_popup_password']")
	public static WebElement loginValidation_PasswordTxtfromHome;
	
	@FindBy(xpath="//*[@id='id_transaction_login_popup_username']")
	public static WebElement loginValidation_UsernameTxtfromHome;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_login_popup_ok']")
	public static WebElement loginValidation_ValidateBtn;
	
	@FindBy(xpath="//*[@id='id_transaction_login_popup_ok']")
	public static WebElement loginValidation_ValidateBtnfromHome;
	
	@FindBy(xpath="//*[@id='id_transaction_entry_login_popup']/div[2]//input[2]")
	public static WebElement loginValidation_CancelBtn;
	
	public boolean checkEditingandSavingVoucher3inPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	
	{
		

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		documentNumberTxt.sendKeys("3");
		Thread.sleep(2000);
		documentNumberTxt.sendKeys(Keys.TAB);
		Thread.sleep(10000);
		
		click(vendorAccountTxt);
		
		vendorAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		
		click(select1stRow_2ndColumn);
		
		click(enter_Quantity);
		enter_Quantity.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_Quantity.sendKeys("25");
		enter_Quantity.sendKeys(Keys.TAB);
		
		click(select2ndRow_2ndColumn);
		
		click(enter_Quantity);
		enter_Quantity.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		enter_Quantity.sendKeys("10");
		enter_Quantity.sendKeys(Keys.TAB);
		
		click(voucherSaveBtn);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		Thread.sleep(8000);
		
		
		//Login Validation
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_ValidateBtn));
		boolean actLoginValid=loginValidation_ValidateBtn.isDisplayed();
		System.out.println("Login Validation  is Displayed     "     +     actLoginValid);
		
		ClickUsingJs(loginValidation_CancelBtn);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		boolean actVoucherSaveBtn=voucherSaveBtn.isDisplayed();
		System.out.println("Voucher Save Btn is Displayed After Cancel	"	+	actVoucherSaveBtn);
		
		click(voucherSaveBtn);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_UsernameTxt));
		loginValidation_UsernameTxt.click();
		loginValidation_UsernameTxt.sendKeys("user1");
		loginValidation_UsernameTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_PasswordTxt));
		loginValidation_PasswordTxt.sendKeys("12345");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_ValidateBtn));
		loginValidation_ValidateBtn.click();
		
		String expMsg="User name must match login user.";
		String actMsg=checkValidationMessage(expMsg);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_UsernameTxt));
		loginValidation_UsernameTxt.click();
		loginValidation_UsernameTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		loginValidation_UsernameTxt.sendKeys("su");
		loginValidation_UsernameTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_PasswordTxt));
		loginValidation_PasswordTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		loginValidation_PasswordTxt.sendKeys("12345");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_ValidateBtn));
		loginValidation_ValidateBtn.click();
		
		String expMsg1="Eitehr user name or password is incorrect.";
		String actMsg1=checkValidationMessage(expMsg1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_PasswordTxt));
		loginValidation_PasswordTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
		loginValidation_PasswordTxt.sendKeys("su");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_ValidateBtn));
		loginValidation_ValidateBtn.click();
		
		Thread.sleep(2000);
		/*HashSet<String> actMessage = new HashSet();

		for (int i = 0; i < 2; i++) 
		{
			String data = checkValidationMessage("");
			actMessage.add(data);
		}

		HashSet<String> expMessage = new HashSet();

		expMessage.add("Voucher saved successfully : 3");
		expMessage.add("Saving in background.");*/
		

		String expMessage="Saving in background.";
		String actMessage=checkValidationMessage(expMessage);


		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage);
		
		
		//click(new_CloseBtn);
		
		  if(actMessage.equalsIgnoreCase(expMessage) ) 
		  { 
			  return true; 
			 } 
		  else 
		  { 
			  return false;
		  }
	
		
	
	}
	
	
	
	public boolean checkSuspendVoucher3FromEntryPageinPurchaseVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			documentNumberTxt.sendKeys("3");
			Thread.sleep(2000);
			documentNumberTxt.sendKeys(Keys.TAB);
			Thread.sleep(10000);
			click(new_SuspendBtn);
			
			/*
			 * getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
			 * loginValidation_UsernameTxt)); loginValidation_UsernameTxt.click();
			 * loginValidation_UsernameTxt.sendKeys("su");
			 * loginValidation_UsernameTxt.sendKeys(Keys.TAB);
			 */
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_PasswordTxt));
			loginValidation_PasswordTxt.click();
			loginValidation_PasswordTxt.sendKeys("su");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_ValidateBtn));
			loginValidation_ValidateBtn.click();
			Thread.sleep(2000);
			
			String expMessage="Voucher saved successfully";
			
			String actMessge=checkValidationMessage(expMessage);
			
			
			
			if(actMessge.startsWith(expMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
			
	
	}
	
	
	public boolean checkEditingVoucher3inPurchaseVoucheratUserLevel() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		
			Thread.sleep(2000);
			
			click(financialsMenu);
			
			 Thread.sleep(2000);
			
			 click(financialsTransactionMenu);
			 Thread.sleep(2000);
			 
			 click(financialsTransactionsPurchaseMenu);
			 Thread.sleep(2000);
			 click(purchasesVoucherBtn1);
			 Thread.sleep(2000);
			 
			 click(newBtn);
			 Thread.sleep(4000);
			 
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
				documentNumberTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
				documentNumberTxt.sendKeys("3");
				Thread.sleep(2000);
				documentNumberTxt.sendKeys(Keys.TAB);
				Thread.sleep(8000);
				
				click(vendorAccountTxt);
				
				vendorAccountTxt.sendKeys(Keys.TAB);
				Thread.sleep(2000);
				
				
				click(select1stRow_2ndColumn);
				click(enter_Quantity);
				enter_Quantity.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
				enter_Quantity.sendKeys("15.22");
				enter_Quantity.sendKeys(Keys.TAB);
				
				click(voucherSaveBtn);
				Thread.sleep(2000);
				
				
				  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
				  pickBtn)); pickBtn.click();
				  
				  Thread.sleep(2000);
				  
				  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(
				  Bill_OkBtn)); Bill_OkBtn.click();
				 
				  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_UsernameTxt));
					loginValidation_UsernameTxt.click();
					loginValidation_UsernameTxt.sendKeys("user1");
					loginValidation_UsernameTxt.sendKeys(Keys.TAB);
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_PasswordTxt));
					loginValidation_PasswordTxt.sendKeys("12345");
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_ValidateBtn));
					loginValidation_ValidateBtn.click();
					Thread.sleep(2000);
				
			/*	HashSet<String> actMessage = new HashSet();

				for (int i = 0; i < 2; i++) 
				{
					String data = checkValidationMessage("");
					actMessage.add(data);
				}

				HashSet<String> expMessage = new HashSet();

				expMessage.add("Voucher saved successfully : 3");
				expMessage.add("Saving in background.");
*/
					

					String expMessage="Saving in background.";
					String actMessage=checkValidationMessage(expMessage);

				System.out.println("Actual Message    : "+actMessage);
				System.out.println("Expected Message  : "+expMessage);
				
				
				click(new_CloseBtn);
				
				  if(actMessage.equalsIgnoreCase(expMessage) ) 
				  { 
					  return true; 
					 } 
				  else 
				  { 
					  return false;
				  }
			
			 
			 
		}
		 
	
	@FindBy(xpath="//*[@id='id_transaction_homescreen_Suspend']/i")
	public static WebElement  suspendBtn;
	
	
	public static boolean checkSuspendtheVoucher3inPurchaseVouchers() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(2000);
		
		checkLogoutAndLoginWithNewlySavedUser();
		
		Thread.sleep(2000);
		
		Thread.sleep(2000);
		
		click(financialsMenu);
		
		 Thread.sleep(2000);
		
		 click(financialsTransactionMenu);
		 Thread.sleep(2000);
		 
		 click(financialsTransactionsPurchaseMenu);
		 Thread.sleep(2000);
		 click(purchasesVoucherBtn1);
		 Thread.sleep(2000);
		 
		 
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_ChkBox1));

		int vcount = grid_VoucherNoList.size();

		for (int i = 0; i < vcount; i++) 
		{
			String VoucherNo = grid_VoucherNoList.get(i).getText();

			if (VoucherNo.equalsIgnoreCase("3")) 
			{
				getAction().click(grid_VoucherChckBoxList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		click(suspendBtn);
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_UsernameTxtfromHome));
		loginValidation_UsernameTxtfromHome.sendKeys("User1");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_PasswordTxtfromHome));
		loginValidation_PasswordTxtfromHome.sendKeys("12345");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(loginValidation_ValidateBtnfromHome));
		loginValidation_ValidateBtnfromHome.click();
		
		
		String expMessage="Voucher Suspended Successfully";
	
		String actMessge=checkValidationMessage(expMessage);
		
		
		
		if(actMessge.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
		
}	
	
	
	public boolean checkLogoutCompanyCreationPage() throws InterruptedException
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
		
	
	
		public  CompanyCreatePage(WebDriver driver)
	     {
	    	PageFactory.initElements(driver, this);
	    	
	     }

}
