package com.focus.base;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.StringSelection;
import java.awt.event.KeyEvent;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Method;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoAlertPresentException;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.PageLoadStrategy;
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

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.markuputils.ExtentColor;
import com.aventstack.extentreports.markuputils.MarkupHelper;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;
/*import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;*/
import com.focus.Pages.LoginPage;
import com.focus.elements.WebElements;
import com.focus.repository.ElementRepository;
import com.focus.utilities.DriverUtility;
import com.focus.utilities.ScreenshotUtility;

import io.github.bonigarcia.wdm.WebDriverManager;



public class BaseEngine extends ElementRepository
{
	private static WebDriver driver;
	protected static String baseDir;
	private static String tcName;
	public static ExtentReports extentReports;
	private static ExtentTest extentTest;
	private static Logger logger;
	private static WebDriverWait wait;
	private static ITestResult result;

	protected String filePath;
	private FileInputStream fip;
	private FileOutputStream fop;
	protected Workbook workbook;
	protected Sheet sheet;
	private Cell cell;
	private CellStyle style;
	protected Row row;
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
			
			

			
			  WebDriverManager.chromedriver().setup();
			  
			  ChromeOptions options1 = new ChromeOptions();
			  options1.addArguments("--start-maximized");
			  options1.addArguments("--disable-notifications");
			  
			   driver = new ChromeDriver(options1); initActivities();
			 

			
			/*
			System.setProperty(DriverUtility.chromeKey, baseDir + DriverUtility.chromeValue);
			String downloadFilepath = getBaseDir() + "\\autoIt\\ExportFiles";
			HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
			chromePrefs.put("profile.default_content_settings.popups", 0);
			chromePrefs.put("download.default_directory", downloadFilepath);
			chromePrefs.put("safebrowsing.enabled", "true");
			ChromeOptions options = new ChromeOptions();
			options.setExperimentalOption("prefs", chromePrefs);		
		
			DesiredCapabilities cap = DesiredCapabilities.chrome();
			cap.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
			cap.setCapability(ChromeOptions.CAPABILITY, options);

			driver = new ChromeDriver(cap);
			initActivities();*/
			
			
			
			
			
			
			
			
			
			
			
			
			
			/*
			 * System.setProperty(DriverUtility.chromeKey,
			 * baseDir+DriverUtility.chromeValue); String downloadFilepath =
			 * getBaseDir()+"\\autoIt\\ExportFiles";
			 * 
			 * HashMap<String, Object> chromePrefs = new HashMap<String, Object>();
			 * chromePrefs.put("profile.default_content_settings.popups", 0);
			 * chromePrefs.put("download.default_directory", downloadFilepath);
			 * //chromePrefs.put("smartWait", 40); chromePrefs.put("safebrowsing.enabled",
			 * "false"); // chromePrefs.put("profile.password_manager_enabled", "false");
			 * ChromeOptions options = new ChromeOptions();
			 * //options.setPageLoadStrategy(PageLoadStrategy.NORMAL);
			 * //options.setExperimentalOption("prefs", chromePrefs); DesiredCapabilities
			 * cap = DesiredCapabilities.chrome();
			 * cap.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
			 * cap.setCapability(ChromeOptions.CAPABILITY, options);
			 * 
			 * 
			 * driver=new ChromeDriver(cap); initActivities();
			 */
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
	  
	
	  
	  
	  
	 /* @BeforeMethod
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

	  }*/
	  
	  
	  
	  @BeforeMethod
	    public void beforeMethodExecution(Method method) {
	        tcName = method.getName();
	        // System.out.println("tc started");
	        logger = Logger.getLogger(BaseEngine.class);

	        getLogger().info("Currently Executing Test Case Name is : " + tcName);
	        //extentTest = extentReports.startTest(tcName);
	        extentTest  =extentReports.createTest(tcName);

	        System.out.println("This Is Before Method Execution And Trying To Executing The Method: " + tcName);
	        // Excel Update is Written as below
	    }
	  
	  
	  
	
	  
	  
	  @AfterMethod
	    public void afterMethodExecution(ITestResult result)
	            throws IOException, EncryptedDocumentException, InvalidFormatException {
	        // System.out.println(res);
	        // System.out.println(result.getStatus());
	        // System.out.println("This is After method");

	        if (result.getStatus() == ITestResult.SUCCESS) {
	            res = "Pass";
	            getLogger().info("Test Case is Passed : " + tcName);
	            // excelReader.setCellData(xlfile, "Sheet3", 0, 0, res);
	            //extentTest.log(LogStatus.PASS, "Test Case is Passed  : " + tcName);
	            extentTest.log(Status.PASS, MarkupHelper.createLabel(result.getName()+" PASSED ", ExtentColor.GREEN));
	            System.out.println("Pass : Test Case is Passed from after Method : " + tcName);
	        }

	        else if (result.getStatus() == ITestResult.FAILURE) {
	            res = "Fail";
	            getLogger().info("Test Case is Failed : " + tcName + "So Taking the Screenshot");
	            // excelReader.setCellData(xlfile, "Sheet3", 0, 1, res);
	            //extentTest.log(LogStatus.FAIL, "Test Case is Failed : " + tcName);
	            extentTest.log(Status.FAIL, MarkupHelper.createLabel(result.getName()+" FAILED ", ExtentColor.RED));
	            ScreenshotUtility.screenshot();
	            System.out.println("Fail : Test Case Is Failed After Method Execution : " + tcName);
	            // System.out.println("Log Message:: @AfterMethod:
	            // Method-"+methodName+"- has
	            // Failed");

	        } else if (result.getStatus() == ITestResult.SKIP) {
	            res = "Skip";
	            getLogger().info("Test Case is Skipped : " + tcName + "So Taking the Screenshot");
	            // excelReader.setCellData(xlfile, "Sheet3", 0, 2, res);
	            //extentTest.log(LogStatus.SKIP, "Test Case is Skipped : " + tcName);            
	            extentTest.log(Status.SKIP, MarkupHelper.createLabel(result.getName()+" SKIP ", ExtentColor.ORANGE));
	            
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
	  
	
	  
	/*  @BeforeTest
	  public void intialiseReports()
	  {
		  extentReports =new ExtentReports(baseDir+"\\reports\\FocusAutomationReport.html");
		  extentReports.assignProject("FocusAI");
	  }
	    
	  */
	  
	  
	/*  @BeforeTest
	     public void intialiseReports()
	     {
	             
	             
	               extentReports = new ExtentReports(baseDir +"\\reports\\"+getCurrentDateFormatWithTime()+"---FocusAutomationReport.html");
	              
	               extentReports.assignProject("FocusAI");
	              
	     } 
		   
		   
	  
	  
	  @AfterTest
	  public void generateReports()
	  {
		  extentReports.flush();
		  extentReports.endTest(extentTest);
	  }
	  
	  */
	 
	 
	 
	 @BeforeTest
	    public void intialiseReports() {
	        //extentReports = new ExtentReports(baseDir + "\\reports\\FocusAutomationReport.html");
	        //extentReports.assignProject("FocusAI");
	        
	        extentReports =new ExtentReports();
	        ExtentSparkReporter spark =new ExtentSparkReporter(baseDir +"\\reports\\"+getCurrentDateFormatWithTime()+"---FocusAutomationReport.html");
	        extentReports.attachReporter(spark);
	        
	      
	    }
	 

		   
		   
	
	 
	 @AfterTest
	    public void generateReports() {
	        extentReports.flush();
	        //extentReports.endTest(extentTest);
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
		 	Thread.sleep(2000);
		 	
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
		 	Thread.sleep(2000);
			
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

	   public static void focusMainSearch(String text) throws InterruptedException {

			click(focusMainSearch);
			Thread.sleep(1000);
			focusMainSearch.sendKeys(text);
			Thread.sleep(2000);
			focusMainSearch.sendKeys(Keys.ENTER);

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
	   
	   public static void IsVisible(WebElement element) {
			new WebDriverWait(getDriver(), 250).until(ExpectedConditions.visibilityOf(element));		
			System.err.println("element is visible");
		}
	   
	   public static void ProghornRestart(String CompanyCode) throws AWTException, InterruptedException 
       
       {
               
               Robot robot=new Robot();
   robot.keyPress(KeyEvent.VK_CONTROL);
   robot.keyPress(KeyEvent.VK_T);
   robot.keyRelease(KeyEvent.VK_CONTROL);
   robot.keyRelease(KeyEvent.VK_T);

   Thread.sleep(4000);

   ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());

   System.out.println("openTabs"+openTabs);

   getDriver().switchTo().window(openTabs.get(0));

   getDriver().switchTo().window(openTabs.get(1));


   //getDriver().get("http://localhost:8011/closecompany?coid=0D0");
   
   getDriver().get("http://localhost:8011/closecompany?coid="+CompanyCode+"");

   Thread.sleep(6000);


  getDriver().switchTo().window(openTabs.get(1));
  Thread.sleep(2500);
                   
  getDriver().switchTo().window(openTabs.get(1)).close();
  Thread.sleep(2000);
  
  getDriver().switchTo().window(openTabs.get(0));
  Thread.sleep(2500);
               
       }
	   
	   
	  ////Restore
	   
	   @FindBy(xpath="//*[@id='ReIndexingError_Modal']//div[3]/button")
		public static WebElement  reindexLoginCancelBtn;
	   
	   public static String restoreCompany(String backupName,String companyName) throws InterruptedException, IOException, AWTException
		{
			
			Thread.sleep(1000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();

			Thread.sleep(1000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
			dataMangementMenu.click();

			Thread.sleep(1000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restore));
			restore.click();

			Thread.sleep(4000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(folderpathExpandBtn));
			folderpathExpandBtn.click();
			
			Thread.sleep(6000);
			
			 Robot rb = new Robot();
		     StringSelection str = new StringSelection(getBaseDir()+"\\backUp\\"+backupName+".fbak");
		     Toolkit.getDefaultToolkit().getSystemClipboard().setContents(str, null);
		 
		     rb.keyPress(KeyEvent.VK_CONTROL);
		     rb.keyPress(KeyEvent.VK_V);
		 
		    rb.keyRelease(KeyEvent.VK_CONTROL);
		    rb.keyRelease(KeyEvent.VK_V);
		 
		    rb.keyPress(KeyEvent.VK_ENTER);
		    rb.keyRelease(KeyEvent.VK_ENTER);

		     
		     Thread.sleep(4000);
		     
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(restoreCompanyBtn));
		     restoreCompanyBtn.click();
		     
		     Thread.sleep(10000);
		    // String alertText="";
		     try
		     {
		    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(overRideYesBtn));
		 		 overRideYesBtn.click();
		 		 
		 		 System.err.println("COMPANY EXISTS");
		 		 
		 	
		 		 
		     }
		     catch(Exception e)
		     {
		    	 System.err.println("NO OLDER COMPANY EXISTS");
		     }
		     
		     new WebDriverWait(getDriver(), 360).until(ExpectedConditions.alertIsPresent());
		     
		     String alertText  =getAlert().getText();
		    		     
		     getAlert().accept();
		     Thread.sleep(6000);
			
		     
		  
			
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		     userNameDisplayLogo.click();
		     
		   
			
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		     logoutOption.click();

		     Thread.sleep(6000);
		     
		     checkLoginToSelectedCompany(companyName,"su","su");
		     Thread.sleep(25000);
		     
		    // new WebDriverWait(getDriver(), 500).until(ExpectedConditions.visibilityOf(reindexLoginCancelBtn));
		     try
		     {
		   
		             if(reindexLoginCancelBtn.isEnabled())
		             { 
		            	 
		            	 System.out.println("try block");
		                     Thread.sleep(4000);
		                     getDriver().navigate().refresh(); 
		                     Thread.sleep(6000);
		                     checkLoginToSelectedCompany(companyName,"su","su");
		                     Thread.sleep(10000);
		             }
		     }
		     catch (Exception e) {
				
			}
	         return alertText;
		   
		}
	   
	   public static String getText(WebElement element) {

			new WebDriverWait(getDriver(), 50).until(ExpectedConditions.elementToBeClickable(element));

			return element.getText();
		}
	   
	   public static String getcurrentSystemTime()
		{
			 String systemtime = new SimpleDateFormat("HH"+":"+"mm").format(Calendar.getInstance().getTime());
	         
	         System.out.println(systemtime );
	        
	         
	         return systemtime;
		}
	   
	   public static void contextClick(WebElement element) {

			try {
				Actions action = new Actions(getDriver()).contextClick(element);

				action.build().perform();

				System.out.println("right clicked the element");
			} catch (StaleElementReferenceException e) {
				System.out.println("Element is not attached to the page document " + e.getStackTrace());
			} catch (NoSuchElementException e) {
				System.out.println("Element " + element + " was not found in DOM " + e.getStackTrace());
			} catch (Exception e) {
				System.out.println("Element " + element + " was not clickable " + e.getStackTrace());
			}
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
		
	 public static String getCurrentDateFormatWithTime() 
     {

             String timeStamp = new SimpleDateFormat("MMM"+" "+"dd"+" "+"yyyy"+" "+"HH"+" "+"mm").format(Calendar.getInstance().getTime());
             //long a = result.getEndMillis()-result.getStartMillis();
             
             System.out.println(timeStamp );
             System.out.println(System.currentTimeMillis());
             return timeStamp;
     }


	 public double checkImageComparison(String file1,String file2)
	 {
	 	
	 	 BufferedImage imgA = null; 
	      BufferedImage imgB = null; 

	     
	      try { 

	          // Reading file from local directory by 
	          // creating object of File class 
	          File fileA 
	              = new File(file1); 
	          File fileB 
	              = new File(file2); 

	          // Reading files 
	          imgA = ImageIO.read(fileA); 
	          imgB = ImageIO.read(fileB); 
	      } 

	     
	      catch (IOException e) { 
	        
	          System.out.println(e); 
	      } 

	      // Assigning dimensions to image 
	      int width1 = imgA.getWidth(); 
	      int width2 = imgB.getWidth(); 
	      int height1 = imgA.getHeight(); 
	      int height2 = imgB.getHeight(); 
	      double percentage = 1;
	      // Checking whether the images are of same size or 
	      // not 
	      if ((width1 != width2) || (height1 != height2)) 

	          // Display message straightaway 
	          System.out.println("Error: Images dimensions"
	                             + " mismatch"); 
	      else { 

	          // By now, images are of same size 

	          long difference = 0; 

	          // treating images likely 2D matrix 

	          // Outer loop for rows(height) 
	          for (int y = 0; y < height1; y++) { 

	              // Inner loop for columns(width) 
	              for (int x = 0; x < width1; x++) { 

	                  int rgbA = imgA.getRGB(x, y); 
	                  int rgbB = imgB.getRGB(x, y); 
	                  int redA = (rgbA >> 16) & 0xff; 
	                  int greenA = (rgbA >> 8) & 0xff; 
	                  int blueA = (rgbA)&0xff; 
	                  int redB = (rgbB >> 16) & 0xff; 
	                  int greenB = (rgbB >> 8) & 0xff; 
	                  int blueB = (rgbB)&0xff; 

	                  difference += Math.abs(redA - redB); 
	                  difference += Math.abs(greenA - greenB); 
	                  difference += Math.abs(blueA - blueB); 
	              } 
	          } 

	          // Total number of red pixels = width * height 
	          // Total number of blue pixels = width * height 
	          // Total number of green pixels = width * height 
	          // So total number of pixels = width * height * 
	          // 3 
	          double total_pixels = width1 * height1 * 3; 

	          // Normalizing the value of different pixels 
	          // for accuracy 

	          // Note: Average pixels per color component 
	          double avg_different_pixels 
	              = difference / total_pixels; 

	          // There are 255 values of pixels in total 
	           percentage 
	              = (avg_different_pixels / 255) * 100; 

	          // Lastly print the difference percentage 
	          System.out.println("Difference Percentage-->"
	                             + percentage); 
	          
	          
	      }
	 	return percentage; 
	 }

	 public static String addingDays(int i)
		{
			LocalDate futureDate = LocalDate.now().plusDays(i);		
		     DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
		     String dateStr = futureDate.format(formatter);	     
		     System.out.println("after adding "+i+" days  :"+ dateStr);
			return dateStr;

		}
	 
	 public static  String listOfElements(List<WebElement> elementList)
	 {
	     
	     int count= elementList.size();
	     
	     ArrayList<String> list = new ArrayList<String>();
	     
	     for (int i = 0; i < count; i++) 
	     {
	         String data=elementList.get(i).getText();
	         
	         if (data.isEmpty() == false) 
	         {
	             list.add(data);
	         }
	             
	     
	         
	     }

	     
	     String actList=list.toString();
	     
	     return actList;
	     
	 }
	 
	 public static void clearValueFromTextBox(WebElement element) throws InterruptedException {

			new WebDriverWait(getDriver(), 50).until(ExpectedConditions.elementToBeClickable(element));
			
			element.sendKeys(Keys.CONTROL,"a");
			Thread.sleep(1500);
			element.sendKeys(Keys.BACK_SPACE);
			Thread.sleep(1500);

			//element.clear();
		}
	 
	 @FindBy(xpath = "(//*[@id='16']/span)[1]")
	private static WebElement SettingsmenuBtn;
	 
	 @FindBy(xpath = "(//span[text()='Preferences'])[2]")
		public static WebElement Setting_PerferenceMenu;
	 
	 @FindBy(xpath = "//a[text()='Inventory']")
		private static WebElement InventoryBtn;

	 @FindBy(xpath = "//*[@id='DonotUseRealTimeRate']//following-sibling::span")
		public static WebElement DonotUseRealTimeRateChkBOx;
		
		@FindBy(xpath = "//*[@id='DonotUseRealTimeRate']")
		public static WebElement donotUseRealTimeRateChkBOxIsSelected;

		
		public static void  logout() throws InterruptedException
		{
			LoginPage lp = new LoginPage(getDriver());

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Usernamearrow));
			Usernamearrow.click();

			Thread.sleep(3000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			logoutOption.click();


		}
	 
	 public static boolean donotuserealtimerate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
		
			ClickUsingJs(SettingsmenuBtn);

			Thread.sleep(1000);
			click(Setting_PerferenceMenu);

			Thread.sleep(1000);

			click(InventoryBtn);

			Thread.sleep(1000);
			
			if(donotUseRealTimeRateChkBOxIsSelected.isSelected()==false)
			{
				ClickUsingJs(DonotUseRealTimeRateChkBOx);
				
				Thread.sleep(1000);;
			
				getWaitForAlert();
			
			
			String actalert=getAlert().getText();
			String expalert="Do you want to enable the same option in Report?";
			
			System.err.println("ACT alert Displayed:  "+actalert);
			System.err.println("EXP alert Displayed:  "+expalert);
			
			
			Thread.sleep(2569);
		
			getAlert().accept();
			
			}
		
			boolean actDonotUseRealTimeRateChkBOx=donotUseRealTimeRateChkBOxIsSelected.isSelected();
			boolean expDonotUseRealTimeRateChkBOx=true;
			
			System.err.println("DonotUseRealTimeRateChkBOx Status: "+actDonotUseRealTimeRateChkBOx+"---------"+expDonotUseRealTimeRateChkBOx);
			
			
			Thread.sleep(1000);
			ClickUsingJs(updateBtn);

			getWaitForAlert();

			getAlert().accept();

			String expMessage = "Data saved Successfully";

			String actMessage = checkValidationMessage(expMessage);

			System.err.println("********* Message  Text: " + actMessage + "  value expected  " + expMessage);
			
			if((actDonotUseRealTimeRateChkBOx==expDonotUseRealTimeRateChkBOx) && (actMessage.equalsIgnoreCase(expMessage)))
			{
				return true;
			}
			
			else 
			{
				return false;
			}
	 }
  
	 
	 public static String getCurrentDate() {
			Date date = Calendar.getInstance().getTime();
			DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
			String strDate = dateFormat.format(date);
			// System.out.println(strDate);
			return strDate;

		}
	 
	 public static  boolean ListComparisionWOOrder(List<WebElement> elementList,String expRowList)
	 {
	     
	     String actRowList = listOfElements(elementList);
	      List List1 = new ArrayList<String>(Arrays.asList(actRowList.replace("[", "").replace("]", "").replace(" ", "").split(",")));
	      List List2 = new ArrayList<String>(Arrays.asList(expRowList.replace("[", "").replace("]", "").replace(" ", "").split(",")));
	     
	      System.out.println("actList:"+List1); 
	 	   System.out.println("expList:"+List2);
	    
	      
	         Collections.sort(List1);
	         Collections.sort(List2);
	         
	    
	    
	    
	     return List1.size() == List2.size()&& List1.equals(List2)?true:false;
	 }   
	   
		
		  public static void prongHornStopAtAdminLevel() throws AWTException,
		  InterruptedException, IOException {
		  
		  
		  
				
				
				
				/*  Thread.sleep(2000);
				  
				  String batCommand
				  ="cmd /c start C:\\Users\\dell\\Desktop\\PronghornStopF10.lnk";
				  
				  
				  Thread.sleep(2000); Runtime.getRuntime().exec(batCommand);
				  
				  Thread.sleep(10000);
				  
				  System.err.println("Pronghorn stopped -----------FOCUS 10");*/
				 
				 
		  
		  
		  }
		  
		  
		  
		  public static void prongHornStartAtAdminLevel() throws AWTException,
		  InterruptedException, IOException {
		  
		  
		  
				
				
				
			/*	  Thread.sleep(2000);
				  
				  String batCommand2 = "cmd /c start C:\\Users\\dell\\Desktop\\IISRESET.lnk";
				  Thread.sleep(2000); Runtime.getRuntime().exec(batCommand2);
				  
				  Thread.sleep(10000);
				  
				  System.err.println("InetManagerRestart");
				  
				  
				  Thread.sleep(2000);
				  
				  String batCommand =
				  "cmd /c start C:\\Users\\dell\\Desktop\\PronghornstartF10.lnk";
				  Thread.sleep(2000); Runtime.getRuntime().exec(batCommand);
				  
				  Thread.sleep(10000);
				  
				  System.err.println("Pronghorn Started FOcus 10 ");*/
				 
				 
		  
		  }
		  
		  public static boolean isElementPresent(By locatorKey) {
			    try {
			        getDriver().findElement(locatorKey);
			        return true;
			    } catch (org.openqa.selenium.NoSuchElementException e) {
			        return false;
			    }
			}
		 
}
