package com.focus.base;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.server.handler.GetAlertText;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.testng.ITestResult;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Optional;
import org.testng.annotations.Parameters;
import org.testng.log4testng.Logger;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;
import com.focus.Pages.LoginPage;
import com.focus.elements.WebElements;
import com.focus.repository.ElementRepository;
import com.focus.utilities.DriverUtility;
import com.focus.utilities.ScreenshotUtility;



public class BaseEngine extends ElementRepository
{
	private static WebDriver driver;
	private static String baseDir;
	private static String tcName;
	private static ExtentReports extentReports;
	private static ExtentTest extentTest;
	private static Logger logger;
	private static WebDriverWait wait;

	private String filePath;
	private FileInputStream fip;
	private FileOutputStream fop;
	private Workbook workbook;
	private Sheet sheet;
	private Cell cell;
	private CellStyle style;
	private Row row;
	private String res=null;
	
	private static Alert alert;
	
	//private String xlfile="\\E:\\ZFocusAI\\TestData.xls";
	
	
	  @Parameters("browser")
	  @BeforeSuite
	  public void openBrowser(@Optional("chrome")String browser)
	  {
		
		baseDir=System.getProperty("user.dir");
		PropertyConfigurator.configure(baseDir + "\\log4j.properties");
		
		if(browser.equalsIgnoreCase("chrome"))
		{	
			System.setProperty(DriverUtility.chromeKey, baseDir+DriverUtility.chromeValue);
			String downloadFilepath = getBaseDir()+"\\autoIt\\ExportFiles";
	        HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
	        chromePrefs.put("profile.default_content_settings.popups", 0);
	        chromePrefs.put("download.default_directory", downloadFilepath);
	        chromePrefs.put("safebrowsing.enabled", "true"); 
	        ChromeOptions options = new ChromeOptions();
	        options.setExperimentalOption("prefs", chromePrefs);
	        DesiredCapabilities cap = DesiredCapabilities.chrome();
	        cap.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
	        cap.setCapability(ChromeOptions.CAPABILITY, options);
			
			driver=new ChromeDriver(cap);
			initActivities();
		}
		
		/*if(browser.equalsIgnoreCase("chrome"))
		{	
			System.setProperty(DriverUtility.chromeKey, baseDir+DriverUtility.chromeValue);
			driver=new ChromeDriver();
			initActivities();
		}*/
		
		if(browser.equalsIgnoreCase("ie"))
		{	
			System.setProperty(DriverUtility.ieKey, baseDir+DriverUtility.ieValue);
			driver=new InternetExplorerDriver();
			initActivities();
		}
		
		if(browser.equalsIgnoreCase("firefox"))
		{
			System.setProperty(DriverUtility.firefoxKey, baseDir+DriverUtility.firefoxValue);
			driver=new FirefoxDriver();
			initActivities();
		}  
	  }
	
	  
	  public void initActivities()
	  {
		 driver.manage().window().maximize();
		 driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
		 driver.manage().deleteAllCookies();
	  }
	  
	  
	  @BeforeMethod
      public void beforeMethodExecution(Method method)
      {
		  tcName=method.getName();
		  //System.out.println("tc started");
		  logger = Logger.getLogger(BaseEngine.class);
		  
		  getLogger().info("Currently Executing Test Case Name is : "+ tcName);
		  extentTest=extentReports.startTest(tcName);
		  
		  System.out.println("This Is Before Method Execution And Trying To Executing The Method: "+ tcName);
		  //Excel Update is Written as below
      }
	  
	  
	  @AfterMethod
	  public void afterMethodExecution(ITestResult result) throws IOException, EncryptedDocumentException, InvalidFormatException
	  {
		  //System.out.println(res);
		  //System.out.println(result.getStatus());
		  //System.out.println("This is After method");
		  
		  if (result.getStatus()==ITestResult.SUCCESS)
		  {
			  res="Pass";
			  getLogger().info("Test Case is Passed : " + tcName);
			  //excelReader.setCellData(xlfile, "Sheet3", 0, 0, res);
			  extentTest.log(LogStatus.PASS, "Test Case is Passed  : " + tcName);
			  System.out.println("Pass : Test Case is Passed from after Method : " + tcName);
		  }
		  
		  else if(result.getStatus()==ITestResult.FAILURE)
		  {
			  res="Fail";
			  getLogger().info("Test Case is Failed : " + tcName + "So Taking the Screenshot");
			  //excelReader.setCellData(xlfile, "Sheet3", 0, 1, res);
			  extentTest.log(LogStatus.FAIL, "Test Case is Failed : " + tcName);
			  ScreenshotUtility.screenshot();
			  System.out.println("Fail : Test Case Is Failed After Method Execution : " + tcName);
			  //System.out.println("Log Message:: @AfterMethod: Method-"+methodName+"- has Failed");
			  
		  }
		  else if(result.getStatus()==ITestResult.SKIP)
		  {
			  res="Skip";
			  getLogger().info("Test Case is Skipped : " + tcName + "So Taking the Screenshot");
			  //excelReader.setCellData(xlfile, "Sheet3", 0, 2, res);
			  extentTest.log(LogStatus.SKIP, "Test Case is Skipped : " + tcName);
			  ScreenshotUtility.screenshot();
			  System.out.println("Skip : Test Case Is Skipped After Method Execution : " + tcName);
		  }

	  }
	  
	  
	  @AfterSuite
	  public void closeBrowser()
	  {
		  if(driver==null)
		  {
			  getLogger().info("Driver is pointing to Null : @AfterSuite");
		  }
		  else
		  {
			  driver.close();
		  }
		  
	  }
	  
	
	  
	  @BeforeTest
	  public void intialiseReports()
	  {
		  extentReports =new ExtentReports(baseDir+"\\reports\\FocusAutomationReport.html");
		  extentReports.assignProject("FocusAI");
	  }
	    
	  
	  @AfterTest
	  public void generateReports()
	  {
		  extentReports.flush();
		  extentReports.endTest(extentTest);
	  }
	  
	  

	  public static String getTCName()
	  {
		  return tcName;
	  }
	  
	  public static String getBaseDir()
	  {
		  return baseDir;
	  }
	  
	  public static WebDriver getDriver()
	  {
		  return driver;
	  }
	  
	  public static ExtentTest getExtentTest()
	  {
		  return extentTest;
	  }
	  
	  public static Logger getLogger()
	  {
		  return logger;
	  }
	  
	   public static void enterUrl(String url) 
	   {
			getDriver().get(url);
	   }
	   
	   
	   public static boolean isElementNotDisplayed(WebElement we)
	   {
		   try
		   {
			   if(we.isDisplayed())
			   {
				   return true;
			   }
			   
			   return false;
		   }
		   catch(Exception e)
		   {
			   return false;
		   }
	   }
	    
	   
	   public static Actions getAction()
	   {
		   Actions action=new Actions(driver);
		   return action;
	
	   }
	  
	   public static void getWaitForAlert()
	   {
			new WebDriverWait(getDriver(), 800).ignoring(NoAlertPresentException.class).until(ExpectedConditions.alertIsPresent());
	   }
		
		
	   public static Alert getAlert()
	   {
			alert=getDriver().switchTo().alert();
			return alert;
	   }
	   
	   public static boolean getIsAlertPresent() 
	   { 
	       try 
	       { 
	           driver.switchTo().alert(); 
	           return true; 
	       } 
	       catch (NoAlertPresentException Ex) 
	       { 
	           return false; 
	       }
	   }  
	  
	   
	  
	   public static Wait getWebDriverWait()
	   {
		    wait = new WebDriverWait(getDriver(), 90);
		    return wait;
		   
	   }
	   
	   public static Wait getFluentWebDriverWait()
	   {
		   Wait<WebDriver> wait = new FluentWait<WebDriver>(driver)							
					.withTimeout(40, TimeUnit.SECONDS) 			
					.pollingEvery(5, TimeUnit.SECONDS) 			
					.ignoring(NoSuchElementException.class);
		   return wait;
	   }
		
	   
	   public static String getCurrentDateF2() 
		{
			Date date = Calendar.getInstance().getTime();
			DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
			String strDate = dateFormat.format(date);
			// System.out.println(strDate);
			return strDate;
		}
		
	   
	   
	   
	   public static String checkDownloadedFileName(WebDriver driver) throws InterruptedException, AWTException
	   {
			/*String mainWindow = driver.getWindowHandle();

			JavascriptExecutor js = (JavascriptExecutor)driver;
			js.executeScript("window.open()");
			
			for(String winHandle : driver.getWindowHandles())
			{
				driver.switchTo().window(winHandle);
			}

			driver.get("chrome://downloads");*/

			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_CONTROL);
			robot.keyPress(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_J);
			robot.keyRelease(KeyEvent.VK_CONTROL);
			
			Thread.sleep(2000);
			
			ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

			int count = openTabs.size();
			
		 	System.out.println("openTabs : "+count);

		 	getDriver().switchTo().window(openTabs.get(count-2));
		 	
		 	getDriver().switchTo().window(openTabs.get(count-1));
		 	
		 	Thread.sleep(2000);
			
			JavascriptExecutor js1 = (JavascriptExecutor)driver;

			String fileName = (String) js1.executeScript("return document.querySelector('downloads-manager').shadowRoot.querySelector('#downloadsList downloads-item').shadowRoot.querySelector('div#content #file-link').text");

			System.err.println("Download deatils");
			System.out.println("File Name : " + fileName);

			Thread.sleep(2000);
			
			/*driver.close();

			driver.switchTo().window(mainWindow);*/
			
			getDriver().switchTo().window(openTabs.get(count-1)).close();
		 	
		 	Thread.sleep(2000);
		 	
		 	getDriver().switchTo().window(openTabs.get(count-2));
			
			return fileName;
		}
	   
	   
	   public static void dropDown(WebElement element, String value) 
	   {
			new Select(element).selectByValue(value);
	   }

	   public static void dropDown(WebElement element, int index) 
	   {
		   new Select(element).selectByIndex(index);
	   }
	   
	   
	   
	   public static void ScrollIntoView(WebElement element)
	   {
		   JavascriptExecutor js = (JavascriptExecutor) getDriver(); 
		   js.executeScript("arguments[0].scrollIntoView();",element);
	   }
	   
	   
	   
	   
	   
	   
	   public static String SavingInBackground(int count) throws InterruptedException
	   {
		   Thread.sleep(2500);
		   
		   String message;
		   
		   HashSet<String> actMessage = new HashSet();
		   
		   for (int i = 0; i < count; i++) 
		   {
			   String data = errorMessageList.get(i).getText();
			   System.err.println(data);
			   actMessage.add(data);
			   errorMessageCloseBtnList.get(i).click();
		   }
		   
		   message = actMessage.toString();
		   
		   return message;
	   }
	   
	   
	   
	   
	   
	   public static String checkValidationMessage(String ExpMessage) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	   {
			Thread.sleep(2000);
			
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
	   
	   
	   
	   
	   
	   public static void checkDeleteLinkStatus() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	   {
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homePannelOpenBtn));
			homePannelOpenBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeBtn));

			if (customizeBtn.isDisplayed()==true) 
			{
				customizeBtn.click();
			}

			if (custRearrangeFormulaBtn.isDisplayed()==true) 
			{
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(custSaveBtn));

				int custBodyHeaderListCount = custBodyHeaderList.size();

				for (int i = 0; i < custBodyHeaderListCount; i++)
				{
					String data = custBodyHeaderList.get(i).getText();

					System.out.println(i+" "+data);

					/*int temp=i+1;

						JavascriptExecutor js = (JavascriptExecutor) getDriver();
						js.executeScript("arguments[0].scrollIntoView();",custBodyHeaderList.get(temp));*/

					if (data.equalsIgnoreCase("Link status")) 
					{
						custBodyHeaderList.get(i).click();

						Thread.sleep(2000);

						getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custRemoveBtn));
						custRemoveBtn.click();

						Thread.sleep(1000);

						break;
					}
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custSaveBtn));
				custSaveBtn.click();

				String expMessage  = "Data saved successfully";
				String actSaveMessage = checkValidationMessage(expMessage);
			} 
			else 
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizeBtn));
				customizeBtn.click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(custSaveBtn));

				int custBodyHeaderListCount = custBodyHeaderList.size();

				for (int i = 0; i < custBodyHeaderListCount; i++)
				{
					String data = custBodyHeaderList.get(i).getText();

					System.out.println(i+" "+data);

					/*int temp=i+1;

						JavascriptExecutor js = (JavascriptExecutor) getDriver();
						js.executeScript("arguments[0].scrollIntoView();",custBodyHeaderList.get(temp));*/

					if (data.equalsIgnoreCase("Link status")) 
					{
						custBodyHeaderList.get(i).click();

						Thread.sleep(2000);

						getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custRemoveBtn));
						custRemoveBtn.click();

						Thread.sleep(1000);

						break;
					}
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(custSaveBtn));
				custSaveBtn.click();

				String expMessage  = "Data saved successfully";
				String actSaveMessage = checkValidationMessage(expMessage);
			}
		}

	   
	  
	   
	   
	   public static void ClickSubMenu(List<WebElement> Element, String ElementName) throws InterruptedException
	   {
		   int count = Element.size();
			
			for (int i = 0; i < count; i++) 
			{
				String data = Element.get(i).getText();
				
				System.err.println(data);
				
				ScrollIntoView(Element.get(i));
				
				if (data.equalsIgnoreCase(ElementName)) 
				{
					Element.get(i).click();
					break;
				}
			}

			Thread.sleep(2000);
	   }
	   
	   
	   
	   
	   
	   public static void ClickUsingJs(WebElement element) 
	   {
			JavascriptExecutor js = (JavascriptExecutor) getDriver();
			js.executeScript("arguments[0].click();", element);
	   }

	   
	   
	   
	   
	   public void moveToElement(WebElement element) 
	   {
           try 
           {
               Actions action = new Actions(getDriver()).moveToElement(element).click();

               action.build().perform();

               System.out.println("Move To element clicked the element");
           } 
          
           catch (StaleElementReferenceException e) 
           {
               System.out.println("Element is not attached to the page document " + e.getStackTrace());
           } 
        
           catch (NoSuchElementException e) 
           {
               System.out.println("Element " + element + " was not found in DOM " + e.getStackTrace());
           } 
           
           catch (Exception e) 
           {
               System.out.println("Element " + element + " was not clickable " + e.getStackTrace());
           }
	   }
	   
  
	   public static void click(WebElement element) {

			new WebDriverWait(getDriver(), 50).until(ExpectedConditions.elementToBeClickable(element));

			element.click();
		}
	   
	   
	   public static void ScrollToElement(WebElement element) {

	   		JavascriptExecutor js = (JavascriptExecutor) getDriver();
	   		js.executeScript("arguments[0].scrollIntoView();", element);
	   	}
	   
	   
	   
	   
	  ////Restore
	   
	   
	   
	   public static String restoreCompany(String backupName,String companyName) throws InterruptedException, IOException, AWTException
		{
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
			dataMangementMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restore));
			restore.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(folderpathExpandBtn));
			folderpathExpandBtn.click();
			
			Thread.sleep(3000);
			
			 Robot rb = new Robot();
		     StringSelection str = new StringSelection(getBaseDir()+"\\backUp\\"+backupName+".fbak");
		     Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str, null);
		 
		     rb.keyPress(KeyEvent.VK_CONTROL);
		     rb.keyPress(KeyEvent.VK_V);
		 
		    rb.keyRelease(KeyEvent.VK_CONTROL);
		    rb.keyRelease(KeyEvent.VK_V);
		 
		    rb.keyPress(KeyEvent.VK_ENTER);
		    rb.keyRelease(KeyEvent.VK_ENTER);

		     
		     Thread.sleep(3000);
		     
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtn));
		     restoreCompanyBtn.click();
		     
		     Thread.sleep(3000);
		     String alertText="";
		     try
		     {
		    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(overRideYesBtn));
		 		 overRideYesBtn.click();
		 		 
		 		 System.err.println("COMPANY EXISTS");
		 		 
		 		Thread.sleep(130000);
		 		 
		     }
		     catch(Exception e)
		     {
		    	 System.err.println("NO OLDER COMPANY EXISTS");
		     }
		     
		     if(getIsAlertPresent())
		     {
		    	 System.err.println("Alert Displayed");
		    	 getWaitForAlert();
		    	 Thread.sleep(2000);
		    	 alertText=getAlert().getText();
		    	 
		    	 getAlert().accept();
		     }
		     
		     Thread.sleep(3000);
			
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		     userNameDisplayLogo.click();
			
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		     logoutOption.click();

		     Thread.sleep(3000);
		     
		     checkLoginToSelectedCompany(companyName,"su","su");
		     
		     Thread.sleep(30000);
		     

		     getDriver().navigate().refresh();
		     
	         Thread.sleep(5000);
		     
	         
	         checkLoginToSelectedCompany(companyName,"su","su");
	         Thread.sleep(3000);
	         
	         return alertText;
		   
		}



	public static void checkLoginToSelectedCompany(String cname,String username,String Password) throws InterruptedException
		{
			  Thread.sleep(3000);
			     
			     LoginPage lp= new LoginPage(getDriver());
			     
			     lp.enterUserName(username);
			     
			     lp.enterPassword(Password);
			     
					String compname = cname;

					Select oSelect = new Select(companyDropDownList);

					List<WebElement> elementCount = oSelect.getOptions();

					int cqSize = elementCount.size();

					System.out.println("CompanyDropdownList Count :" + cqSize);

					int i;

					for (i = 0; i < elementCount.size(); i++) {

						elementCount.get(i).getText();

						String optionName = elementCount.get(i).getText();
						if (optionName.toUpperCase().startsWith(compname.toUpperCase())) {
							System.out.println("Logined Company" + elementCount.get(i).getText());
							elementCount.get(i).click();
						}

					}

					Thread.sleep(2000);

					lp.clickOnSignInBtn();
		}
		
	   
	   
	   
	   
	   
	   
}
