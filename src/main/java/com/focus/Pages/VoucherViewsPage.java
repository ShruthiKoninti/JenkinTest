package com.focus.Pages;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

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
import org.openqa.selenium.support.ui.WebDriverWait;

import com.focus.Pages.LoginPage;
import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.DriverUtility;
import com.focus.utilities.POJOUtility;

public class VoucherViewsPage extends BaseEngine{
	
	@FindBy(xpath="//span[@class='dropdown icon-menuicon1 icon-font6 theme_color-inverse pull-right']")
	public static WebElement  toogleExpandBtn;
	
	
	public static boolean checkLogin()
			throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException {
		Thread.sleep(3000);

		getDriver().navigate().refresh();
		Thread.sleep(6000);

		LoginPage lp = new LoginPage(getDriver());

		lp.checkLoginPageTitleByURLInputInBrowser();
		

		String unamelt = "su";

		String pawslt = "su";

		lp.enterUserName(unamelt);
		Thread.sleep(1000);
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

		Thread.sleep(8000);

		String actUserInfo1 = userNameDisplay.getText();

		System.out.println("User Info  : " + actUserInfo1);

		System.out.println("User Info Capture Text  :  " + userNameDisplay.getText());

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();*/

		String getCompanyTxt1 = Company_Name.getText();
		String getLoginCompanyName1 = getCompanyTxt1.substring(0, 31);
		System.out.println("company name  :  " + getLoginCompanyName1);
		//companyLogo.click();

		String expUserInfo1 = "SU";
		String expLoginCompanyName1 = "Automation Company : 08/10/2020";

		System.out.println("UserInfo1             : " + actUserInfo1 + " Value Expected : " + expUserInfo1);
		System.out.println(
				"LoginCompanyName1     : " + getLoginCompanyName1 + " Value Expected : " + expLoginCompanyName1);

		if (actUserInfo1.equalsIgnoreCase(expUserInfo1) && getLoginCompanyName1.contains(expLoginCompanyName1)) {
			return true;
		} else {
			return false;
		}

	}
	  
	@FindBy(xpath="//*[@id='61']")
	public static WebElement transctionsbtn;

	@FindBy(xpath="//*[@id='2007']/span")
	public static WebElement purchacebtn;
	
	@FindBy(xpath = "//*[@class='icon-menudots hiconright2']")
	public static WebElement entryPageToggleBtn;
	
	
	public boolean checkEditingtheFieldsinViewWithoutSavingView() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(2000);
		
		click(finacinalsMenu);
		
		 Thread.sleep(2000);
		
		 click(transctionsbtn);
		
		 
		 click(purchacebtn);
		
		 click(purchaseVouchersBtn);
		 Thread.sleep(2000);
		 
		 click(newBtn);
		 Thread.sleep(2000);
		 
		 click(entryPageToggleBtn);
		 Thread.sleep(2000);
		 
		 click(settingsOption);
		 Thread.sleep(2000);
		 	
		
		  click(viewsBtn);
		  Thread.sleep(2000);
		  
		click(ViewExistingViewTxt);
		Thread.sleep(1500);
			
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			
		
		ViewExistingViewTxt.sendKeys("Test View");
		
		Thread.sleep(2000);
		
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(8000);
		int Count = viewsGridFiledsList.size();
				
		System.out.println("Count    : "+Count);
		
		ArrayList<String> actviewsGridFiledsList  = new ArrayList<String>();
		
		for (int i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

			System.out.println("data:"+data);
    		if(data.equalsIgnoreCase("Date"))
			{
    			
    			viewsGridFiledsEditList.get(i).click();
    			//ClickUsingJs(viewsGridFiledsEditList.get(i));
				break;
			}
		}
		
	
		
       Thread.sleep(4000);
       
    
       
		
       String expMessage="Select view from existing view dropdown before editing field.";
       
       String actMessage=checkValidationMessage(expMessage);
       
       System.out.println("Actual Message	"	+	actMessage);
       System.out.println("Expected Message	"	+	expMessage);
       
       getAction().moveToElement(settings_closeBtn).build().perform();
       click(settings_closeBtn);
       
       if(actMessage.equalsIgnoreCase(expMessage))
       {
    	   return true;
       }
       else
       {
    	   return false;
       }
	}
	
	@FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_Caption']")
	public static WebElement viewEditCaptionTxt;
	
	
	//@FindBy(xpath="//*[@id='callbackbtn']/i")
	@FindBy(xpath="//*[@id='configTransMenu']/ul/li[2]/span[2]")
	public static WebElement  closeBtn;
	
	@FindBy(xpath="//*[@id='id_transactionentry_header1_section']//label[text()='Department']")
	public static WebElement DepartmentCaption;
	
	@FindBy(xpath="//*[@id='id_transactionentry_header1_section']//label[text()='Financial Tag']")
	public static WebElement Department1Caption;
	
	
	@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_control_heading_1']/div[1]")
	public static WebElement  ItemCaption;
	
	
	
	public boolean checkCaptionBeforeEditingViewinVoucher() throws InterruptedException
	{
		Thread.sleep(2000);
			 
		 click(newBtn);
		 Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(DepartmentCaption));
		 String actDeptCationTxt=DepartmentCaption.getText();
		 String expDeptCationTxt="Department";
		 
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemCaption));
		 String actItemCationTxt=ItemCaption.getText();
		 String expItemCationTxt="Item";
		 
		 
		 System.out.println("Actual		"	+	actDeptCationTxt	+	"Expected		"	+	expDeptCationTxt);
		 System.out.println("Actual		"	+	actItemCationTxt	+	"Expected		"	+	expItemCationTxt);
		 
		 if(actDeptCationTxt.equalsIgnoreCase(expDeptCationTxt) && actItemCationTxt.equalsIgnoreCase(expItemCationTxt))
		 {
			 return true;
		 }
		 else
		 {
			 return false;
		 }
		
	}
	
	@FindBy(xpath="//*[contains(text(),'View Settings')]")//(//*[@id='panelsStayOpen-headingOne']/button)[2]
	private static WebElement viewSettingsBtn;
	
	public boolean  checkCreatingViewinViewTabandChangeCaption() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		 Thread.sleep(5000);
		 click(entryPageToggleBtn);
		 Thread.sleep(2000);
		 
		 click(settingsOption);
		 Thread.sleep(2000);
		 
		
		
		  click(viewsBtn);
		  Thread.sleep(2000);
		  
		
		Thread.sleep(2000);
		click(ViewExistingViewTxt);
	
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys("Test View");
		Thread.sleep(4000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
				
        Thread.sleep(5500);
	
		
		ClickUsingJs(viewSettingsBtn);
		
		if(viewSuChkboxIsSelected.isDisplayed()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		ClickUsingJs(viewTabLoginBtn);
		}
		Thread.sleep(2000);

		if(viewSuChkboxIsSelected.isSelected()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSuChkbox));
		viewSuChkbox.click();
		}
		Thread.sleep(2000);
		
       ClickUsingJs(viewSaveView);
       Thread.sleep(5500);
       
       checkValidationMessage("Data saved successfully");
       
       ClickUsingJs(ViewExistingViewTxt);
   	
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys("Test View");
		Thread.sleep(4000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);
	
		
       Thread.sleep(8000);
     
		int Count = viewsGridFiledsList.size();
	
		for (int i = 0; i < Count; i++) 
		{
			getAction().moveToElement(viewsGridFiledsList.get(i)).build().perform();
			String data		  = viewsGridFiledsList.get(i).getText();

    		if(data.equalsIgnoreCase("Department"))
			
			{	
    			
    			viewsGridFiledsEditList.get(i).click();
				
				break;
			}
		}
		
		Thread.sleep(4000);
       
      
	click(viewEditCaptionTxt)	;
	Thread.sleep(2000);
	viewEditCaptionTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	viewEditCaptionTxt.sendKeys("Financial Tag");
	Thread.sleep(1000);
	viewEditCaptionTxt.sendKeys(Keys.TAB);
	
      getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_ApplyBtn));
      fieldDetails_ApplyBtn.click();
		
      
      Thread.sleep(8000);
      
      for (int i = 0; i < Count; i++) 
		{
    		getAction().moveToElement(viewsGridFiledsList.get(i)).build().perform();
			String data		  = viewsGridFiledsList.get(i).getText();

  		if(data.equalsIgnoreCase("Item"))
			
			{
  			viewsGridFiledsEditList.get(i).click();
				
				break;
			}
		}
		
		Thread.sleep(8000);
     
		click(viewEditCaptionTxt)	;
		Thread.sleep(2000);
		viewEditCaptionTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		viewEditCaptionTxt.sendKeys("Product");
		Thread.sleep(1000);
		viewEditCaptionTxt.sendKeys(Keys.TAB);
		
		click(fieldDetails_ApplyBtn);
	    
	      Thread.sleep(8000);
      ClickUsingJs(viewSaveView);
     
      Thread.sleep(2000);
       String expMessage= "Data saved successfully";
       
       String actMessage=checkValidationMessage(expMessage);
       
       Thread.sleep(4000);
       
      
       ClickUsingJs(updateBtn);
       Thread.sleep(6500);

       ClickUsingJs(settings_closeBtn);
     
       
       if (actMessage.equalsIgnoreCase(expMessage) ) 
       { 
			return true;
		  }
		  else
		  {
			
			return false;
		  }
	}
	
	public boolean checkCaptionAfterEditingViewinVoucher() throws InterruptedException
	{
		 Thread.sleep(5000);
		 click(newBtn);
		 Thread.sleep(2000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Department1Caption));
		 String actDeptCationTxt=Department1Caption.getText();
		 String expDeptCationTxt="Financial Tag";
		 
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ItemCaption));
		 String actItemCationTxt=ItemCaption.getText();
		 String expItemCationTxt="Product";
		 
		 
		 System.out.println("Actual		"	+	actDeptCationTxt	+	"Expected		"	+	expDeptCationTxt);
		 System.out.println("Actual		"	+	actItemCationTxt	+	"Expected		"	+	expItemCationTxt);
		 
		 if(actDeptCationTxt.equalsIgnoreCase(expDeptCationTxt) && actItemCationTxt.equalsIgnoreCase(expItemCationTxt))
		 {
			 return true;
		 }
		 else
		 {
			 return false;
		 }
		
	}
	
	
	@FindBy(xpath="//*[@placeholder='Search User']")
	public static WebElement viwes_seaarchUserTxt;
	
	@FindBy(xpath="//li[@data-text='UserAllOptionsST']//label")
	public static WebElement viwes_UseralloptionsstLabel;
	
	public boolean checkSavingViewbySearchingUserfromSerchTxtinViewsTab() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 Thread.sleep(5500);
		 click(entryPageToggleBtn);
		 Thread.sleep(2000);
		 
		 click(settingsOption);
		 Thread.sleep(2000);
		
		  click(viewsBtn);
		  Thread.sleep(4000);
		  
		
		Thread.sleep(2000);
		click(ViewExistingViewTxt);
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys("Test View");
		Thread.sleep(4000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
				
        Thread.sleep(8000);
        
        
        
       // ((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,1000)","");
       // getAction().moveToElement(viewSettingsBtn).build().perform();
		//Thread.sleep(2000);
		
		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSettingsBtn));
		ClickUsingJs(viewSettingsBtn);
		Thread.sleep(2000);
		
		if(viewsUserAllOptionsSTChkboxSelected.isDisplayed()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		}
		Thread.sleep(2000);

		if(viewsUserAllOptionsSTChkboxSelected.isDisplayed()==true)
		{
			click(viwes_seaarchUserTxt);
			viwes_seaarchUserTxt.sendKeys("UserAllOptionsST");
			viwes_seaarchUserTxt.sendKeys(Keys.TAB);
			Thread.sleep(2000);
		}
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viwes_UseralloptionsstLabel));
		 String actUserTxt=viwes_UseralloptionsstLabel.getText();
		 String expUserTxt="UserAllOptionsST";
		 
		 System.out.println("Actual		"	+	actUserTxt		+	"Expected	"	+	expUserTxt);
		 
		if(viewsUserAllOptionsSTChkboxSelected.isSelected()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		viewsUserAllOptionsSTChkbox.click();
		}
		Thread.sleep(2000);

      //getAction().moveToElement(viewSaveView).build().perform(); 
        
	
	       click(viewSaveView);
	       Thread.sleep(4000);
	       
	       String expMsg="Data saved successfully";
	       String actMsg=checkValidationMessage(expMsg);
	       
	       if(actUserTxt.equalsIgnoreCase(expUserTxt)&&actMsg.equalsIgnoreCase(expMsg))
	       {
	    	   return true;
	       }
	       else
	       {
	    	   return false;
	       }
	}
	
	@FindBy(xpath = "//*[@id='editScreen_FieldsCustomization_DataType']")
	public static WebElement fieldDetails_DataTypeDrpdwn;

	
	public boolean checkEditingEditScreenFieldsinViews() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		click(ViewExistingViewTxt);
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys("Test View");
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
				
        Thread.sleep(10000);
        
        int Count = viewsGridFiledsList.size();
    	
		for (int i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

    		if(data.equalsIgnoreCase("RD"))
			
			{
    			viewsGridFiledsEditList.get(i).click();
				
				break;
			}
		}
		
		Thread.sleep(4000);
        click(fieldDetails_DataTypeDrpdwn);
		
		Select data1=new Select(fieldDetails_DataTypeDrpdwn);
		data1.selectByVisibleText("Read Only");
		
		click(fieldDetailsEditScreen_ApplyBtn);
		
         Thread.sleep(10000);
    
 		for (int i = 0; i < Count; i++) 
 		{
 			String data		  = viewsGridFiledsList.get(i).getText();

     		if(data.equalsIgnoreCase("FD%"))
 			
 			{
     			viewsGridFiledsEditList.get(i).click();
 				
 				break;
 			}
 		}
 		
 		Thread.sleep(8000);
         
 		 click(fieldDetails_DataTypeDrpdwn);
 		
 		Select data2=new Select(fieldDetails_DataTypeDrpdwn);
 		data2.selectByVisibleText("Read Only");
 		
 		click(fieldDetailsEditScreen_ApplyBtn);
          Thread.sleep(8000);
          
          
          ClickUsingJs(viewSaveView);
	       Thread.sleep(4000);
	       
	       String expMsg="Data saved successfully";
	       String actMsg=checkValidationMessage(expMsg);
	       
	       if(actMsg.equalsIgnoreCase(expMsg))
	       {
	    	   return true;
	       }
	       else
	       {
	    	   return false;
	       }
        
	}
	
	@FindBy(xpath="(//*[@class='icon-up-arrow Fbutton'])[2]")
	public static WebElement upwardArrowBtn;
	
	@FindBy(xpath="(//*[@class='icon-down-arrow Fbutton'])[2]")
	public static WebElement downwardArrowBtn;
	
	
	@FindBy(xpath="//*[@id='views_FieldsTable']/tbody/tr/td[2]")
	public static List<WebElement>  viewsGridFiledsPositionList;
	
	public boolean checkUpwardandDownwardArrowsinViews() throws InterruptedException
	{
		Thread.sleep(5000);
		click(ViewExistingViewTxt);
		Thread.sleep(1000);
		
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys("Test View");
		Thread.sleep(4000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
				
        Thread.sleep(7000);
        
        int Count = viewsGridFiledsList.size();
    	int i;
		for ( i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

    		if(data.equalsIgnoreCase("Department"))
			
			{
    			viewsGridFiledsList.get(i).click();
				
				break;
			}
		}
	
		String actPosition=viewsGridFiledsPositionList.get(i).getText();
		String expPosition="9";
		
		System.out.println("Actual		"	+	actPosition		+	"Expected	"	+	expPosition);
		
		
		ClickUsingJs(upwardArrowBtn);
		ClickUsingJs(upwardArrowBtn);
		
		
		for ( i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

    		if(data.equalsIgnoreCase("Department"))
			
			{
    			viewsGridFiledsList.get(i).click();
				
				break;
			}
		}
	
		String actPositionAfter=viewsGridFiledsPositionList.get(i).getText();
		String expPositionAfter="7";
		
		System.out.println("Actual		"	+	actPositionAfter		+	"Expected	"	+	expPositionAfter);
		
		//Downward
		
		for ( i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

    		if(data.equalsIgnoreCase("Item"))
			
			{
    			viewsGridFiledsList.get(i).click();
				
				break;
			}
		}
	
		String actPositionBeforeDown=viewsGridFiledsPositionList.get(i).getText();
		String expPositionBeforeDown="11";
		
		System.out.println("Actual		"	+	actPositionBeforeDown		+	"Expected	"	+	expPositionBeforeDown);
		
		
		ClickUsingJs(downwardArrowBtn);
		ClickUsingJs(downwardArrowBtn);
		
		
		for ( i = 0; i < Count; i++) 
		{
			String data	= viewsGridFiledsList.get(i).getText();

    		if(data.equalsIgnoreCase("Item"))
			
			{
    			viewsGridFiledsList.get(i).click();
				
				break;
			}
		}
	
		String actPositionAfterDown=viewsGridFiledsPositionList.get(i).getText();
		String expPositionAfterDown="13";
		
		System.out.println("Actual		"	+	actPositionAfterDown		+	"Expected	"	+	expPositionAfterDown);
		
		
		/*click(updateBtn);
		Thread.sleep(2000);
		
		click(settings_closeBtn);
		Thread.sleep(2000);*/
		
		
		if(actPosition.equalsIgnoreCase(expPosition) && actPositionAfter.equalsIgnoreCase(expPositionAfter)
				&& actPositionBeforeDown.equalsIgnoreCase(expPositionBeforeDown)&&actPositionAfterDown.equalsIgnoreCase(expPositionAfterDown))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	
	}
	
	public boolean checkAddingNewFieldinEditLayoutTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
			Thread.sleep(2000);
			 click(newBtn);
			 Thread.sleep(2000);
	
			 click(entryPageToggleBtn);
			 Thread.sleep(2000);
			 
			 click(settingsOption);
			 Thread.sleep(2000);
			 
			
			click(editLayoutTab);
			Thread.sleep(2000);
			click(editLayoutAddFieldsBtn);
			getWaitForAlert();
			Thread.sleep(2000);
			getAlert().accept();
			Thread.sleep(10000);
			
			click(editLayoutCaptionTxt);
			
			editLayoutCaptionTxt.sendKeys("EditLayoutHeaderTrial");
			Thread.sleep(8000);
			editLayoutCaptionTxt.sendKeys(Keys.TAB);
			Thread.sleep(4000);
			
			//getAction().moveToElement(editLayoutApplyBtn).build().perform();
			//Thread.sleep(2000);
			//click(editLayoutApplyBtn);
			ClickUsingJs(editLayoutApplyBtn);
			Thread.sleep(2000);
			
			checkValidationMessage("");
			ClickUsingJs(editLayoutApplyBtn);
			
			String expMessage = "Data saved successfully";
			String actMessage = checkValidationMessage(expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				return true;
			}
			else
			{
				return false;
			}
	}
	
	@FindBy(xpath="//*[@class='icon-refresh vouchericon']")
	public static WebElement viewRefreshViewBtn;
	
	@FindBy(xpath="//*[@id='navigationtab4']")
	public static WebElement viewsTab;

	
	public boolean checkRefreshViewAfterAddingNewFieldinEditLayoutTab() throws InterruptedException
	{
		click(viewsTab);
		click(ViewExistingViewTxt);
		
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys("Test View");
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
				
        Thread.sleep(7000); 
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewRefreshViewBtn));
		 viewRefreshViewBtn.click();
		 Thread.sleep(2000);
		 
		 int Count = viewsGridFiledsList.size();
	    	boolean flag=false;
			for ( int i = 0; i < Count; i++) 
			{
				String data		  = viewsGridFiledsList.get(i).getText();

	    		if(data.equalsIgnoreCase("EditLayoutHeaderTrial"))
				
				{
	    			flag=true;
					
					break;
				}
			}
		 
		 if(flag)
		 {
			 return true;
		 }
		 else
		 {
			 return false;
		 }
	 
	}
	
	
	
	public boolean checkCreatingViewinViewsTab() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(5500);
		/*click(viewsTab);
		Thread.sleep(4000);*/
		click(ViewExistingViewTxt);
	
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys("Focus View");
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
				
        Thread.sleep(8000);
	
        
        
       // ((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,1000)","");
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSettingsBtn));
		viewSettingsBtn.click();
		Thread.sleep(2000);
		
	if(viewsUserAllOptionsSTChkboxSelected.isDisplayed()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		}
		Thread.sleep(2000);

		if(viewSuChkboxIsSelected.isSelected()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSuChkbox));
		viewSuChkbox.click();
		}
		Thread.sleep(2000);
		
		getAction().moveToElement(viewSaveView).build().perform();
		              
      
      click(viewSaveView);
     
      Thread.sleep(2000);
       String expMessage= "Data saved successfully";
       
       String actMessage=checkValidationMessage(expMessage);
       
       getAction().moveToElement(settings_closeBtn).build().perform();
       click(updateBtn);
       Thread.sleep(8000);
      
       click(settings_closeBtn);
     
       Thread.sleep(4000);
       if (actMessage.equalsIgnoreCase(expMessage) ) 
       { 
			return true;
		  }
		  else
		  {
			
			return false;
		  }
		
	}
	
	@FindBy(xpath="//*[@id='selectView']")
	public static WebElement multiViewsSelect;
	
	@FindBy(xpath="//*[@class='icon-ok hiconright2']")
	public static WebElement multiViewsOkBtn;
	
	public boolean checkVoucherAfterMultipleViewsCreatedatHomeScreen() throws InterruptedException
	{
		Thread.sleep(5500);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(multiViewsSelect));
		boolean multiViewDisplayed=multiViewsSelect.isDisplayed();
		System.out.println("Multiple View Displayed		"	+	multiViewDisplayed);
		
		Select s=new Select(multiViewsSelect);
	
		for(int i=0;i<s.getOptions().size();i++)
		{
			System.out.println(s.getOptions().get(i).getText());
		}
		
		s.selectByVisibleText("Test View");
		
		click(multiViewsOkBtn);
		Thread.sleep(4000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		boolean newBtnDisplayed=newBtn.isDisplayed();
		System.out.println("New Button Displayed	"	+	newBtnDisplayed);
		
		click(newBtn);
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		boolean documentNumTxDisplayed=documentNumberTxt.isDisplayed();
		System.out.println("Doc Num Displayed	"	+	documentNumTxDisplayed);
		
		if(multiViewDisplayed && newBtnDisplayed && documentNumTxDisplayed)
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	@FindBy(xpath="//i[@class='icon icon-export vouchericon']")
	public static WebElement exportView;
	
	
	@FindBy(xpath="//i[@class='icon icon-import vouchericon']")
	public static WebElement importView;
	
	public boolean checkImportandExportViewsinViewsTab() throws InterruptedException, IOException, AWTException, EncryptedDocumentException, InvalidFormatException
	{
		Thread.sleep(5000);
		 click(entryPageToggleBtn);
		 Thread.sleep(2000);
		 
		 click(settingsOption);
		 Thread.sleep(2000);
		 	
		
		  click(viewsBtn);
		  Thread.sleep(2000);
		  
		click(ViewExistingViewTxt);
		Thread.sleep(1500);
			
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			
		
		ViewExistingViewTxt.sendKeys("Test View");
		
		Thread.sleep(4000);
		
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(4000);
		
		click(exportView);
		Thread.sleep(2000);
		
		Robot robot = new Robot();
		robot.keyPress(KeyEvent.VK_CONTROL);
		robot.keyPress(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_J);
		robot.keyRelease(KeyEvent.VK_CONTROL);
		
		Thread.sleep(5000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		int actOpenWindowsCount = getDriver().getWindowHandles().size();
		
		System.err.println("No of tabs : "+actOpenWindowsCount);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		Thread.sleep(2000);
		
		getDriver().switchTo().window(openTabs.get(0));
		Thread.sleep(2000);
		
		click(ViewExistingViewTxt);
		Thread.sleep(1500);
			
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			
		
		ViewExistingViewTxt.sendKeys("Import View");
		
		Thread.sleep(4000);
		
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(4000);
		
		click(importView);
		Thread.sleep(2000);
		
		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\ImportView.exe");
		
		Thread.sleep(15000);
		
		click(viewSaveView);
		String expMessage="Data saved successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	@FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_defaultValue']")
	public static WebElement fieldDetails_DefaultValue;
	
	
	
	@FindBy(xpath="//*[@id='editScreen_FieldsCustomization_DefaultValue']")
	public static WebElement fieldDetailsF1_DefaultValue;
	
	public boolean checkSavingViewwithDefaultValue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		
		Thread.sleep(5000);
		
		click(ViewExistingViewTxt);
		Thread.sleep(1500);
			
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			
		
		ViewExistingViewTxt.sendKeys("Test View");
		
		Thread.sleep(4000);
		
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(8000);
		
		int Count = viewsGridFiledsList.size();
		
		for (int i = 0; i < Count; i++) 
		{
			getAction().moveToElement(viewsGridFiledsList.get(i)).build().perform();
			String data		  = viewsGridFiledsList.get(i).getText();

    		if(data.equalsIgnoreCase("PurchaseAC"))
			
			{
    			viewsGridFiledsEditList.get(i).click();
				
				break;
			}
		}
		
		Thread.sleep(8000);
		
		click(fieldDetails_DefaultValue);
		fieldDetails_DefaultValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		fieldDetails_DefaultValue.sendKeys("7");
		click(fieldDetails_ApplyBtn);
		Thread.sleep(3000);
		ClickUsingJs(viewSaveView);
		Thread.sleep(3000);
		String expMessage="Data saved successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
		
		Thread.sleep(2000);
		// getAction().moveToElement(updateBtn).build().perform();
		//Thread.sleep(2000);
		ClickUsingJs(updateBtn);
		Thread.sleep(5000);
		ClickUsingJs(settings_closeBtn);
		Thread.sleep(4000);
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	
	@FindBy(xpath="//*[@id='id_body_23']")
	public static WebElement itemtxt;
	
	@FindBy(xpath="//input[@id='id_body_26']")
	public static WebElement quantitytxt;	

	@FindBy(xpath="//*[@id='id_header_67108884']")
	public static WebElement narationtx;	

	@FindBy(xpath="//input[@id='id_body_27']")
	public static WebElement  rateTxt;
	
	public boolean checkSavingVoucherwithDefaultValueinPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(5500);
		
		Select s=new Select(multiViewsSelect);
		
		s.selectByVisibleText("Test View");
		
		click(multiViewsOkBtn);
		Thread.sleep(2000);
		click(newBtn);
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
		String actPurchaseAccTxt=purchaseAccountTxt.getAttribute("value");
		String expPurchaseAccTxt="Purchase";
		
		System.out.println("Actual		"	+	actPurchaseAccTxt		+		"Expected	"	+	expPurchaseAccTxt);
		Thread.sleep(2000);
		purchaseAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		click(vendorAccountTxt);
		vendorAccountTxt.sendKeys("Vendor A");
		vendorAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		warehouseTxt.sendKeys("Hyderabad");
		warehouseTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		departmentTxt.sendKeys("INDIA");
		Thread.sleep(4000);
		departmentTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		click(select1stRow_1stColumn);
		click(itemtxt);
		itemtxt.sendKeys("STD RATE COGS");
		Thread.sleep(2000);
		itemtxt.sendKeys(Keys.TAB);
		
		click(select1stRow_4thColumn);
		click(quantitytxt);
		quantitytxt.sendKeys("5");
		//quantitytxt.sendKeys(Keys.TAB);
		Thread.sleep(4000);
		click(select1stRow_6thColumn);
		click(rateTxt);
		rateTxt.clear();
		rateTxt.sendKeys("5");
		Thread.sleep(2000);
		
		click(voucherSaveBtn);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		
		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		
		if(actMessage.startsWith(expMessage1) && actPurchaseAccTxt.equalsIgnoreCase(expPurchaseAccTxt))
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	@FindBy(xpath="//*[@class='icon-pick']")
	private static WebElement pickBtn;

	@FindBy(xpath="//input[@id='txtNewReference']")
	private static WebElement newReferenceTxt;

	@FindBy(xpath="//*[@class='icon-ok']")
	private static WebElement Bill_OkBtn;
	
	
	@FindBy(xpath="//*[@id='views_FieldsTable']/tbody/tr/td[8]")
	public static List<WebElement>  viewsGridDefaultValueList;
	
	@FindBy(xpath="//*[@id='views_FieldsTable']//tr[4]//td[8]")
	public static WebElement  PurchaseviewsGridDefaultValueList;
	
	@FindBy(xpath="//*[@id='EditLayout_FieldsCustomization_FieldDetails_defaultValue']")
	private static WebElement Defaultvalue;
	
	
	
	public boolean checkSavingViewwithLoadDefaults() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(5000);
		 click(entryPageToggleBtn);
		 Thread.sleep(2000);
		 
		 click(settingsOption);
		 Thread.sleep(2000);
		
		  click(viewsBtn);
		  Thread.sleep(2000);
		  
		click(ViewExistingViewTxt);
		Thread.sleep(1500);
			
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			
		
		ViewExistingViewTxt.sendKeys("Test View");
		
		Thread.sleep(4000);
		
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(8000);
		
		/*int Count = viewsGridFiledsEditList.size();
		String actDefaultValue="";
		
		for (int i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsEditList.get(i).getText();

    		if(data.equalsIgnoreCase("PurchaseAC"))
			
			{	
    			
    			actDefaultValue=PurchaseviewsGridDefaultValueList.getText();
				
				break;
			}
		}*/
		
		
		String actDefaultValue=PurchaseviewsGridDefaultValueList.getText();
		String expDefaultValue="7";
		System.out.println("Default Vlaue Before Load Defaults		:"		+	"Actual		"	+	actDefaultValue		+"Expected		"	+	expDefaultValue);
		
		Thread.sleep(4000);
		
		getAction().moveToElement(views_LoadDefault).build().perform();
		Thread.sleep(2000);
		click(views_LoadDefault);
		Thread.sleep(4000);
		
		getWaitForAlert();
		getAlert().accept();
		Thread.sleep(8000);
		
		/*boolean actDefaultValue1=false;
		
		for (int i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsEditList.get(i).getText();

    		if(data.equalsIgnoreCase("PurchaseAC"))
			
			{
    			actDefaultValue1=PurchaseviewsGridDefaultValueList.getText().isEmpty();
				
				break;
			}
		}*/
		
		boolean actDefaultValue1=PurchaseviewsGridDefaultValueList.getText().isEmpty();
		boolean expDefaultValue1=true;
		System.out.println("Default Vlaue After Load Defaults		:"		+	"Actual		"	+	actDefaultValue1		+"Expected		"	+	expDefaultValue1);
		
		ClickUsingJs(viewSaveView);
		
		String expMessage="Data saved successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
		
		
		if(actDefaultValue.equalsIgnoreCase(expDefaultValue) && actDefaultValue1==expDefaultValue1)
		{
		
		return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	
	
	@FindBy(xpath="(//i[@class='icon-load vouchericon'])[2]")
	public static WebElement views_LoadDefault;
	
	
	public boolean checkSavingViewWithDefaultValueasGroupID() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(5500);
		click(ViewExistingViewTxt);
		Thread.sleep(1500);
			
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			
		
		ViewExistingViewTxt.sendKeys("Test View");
		
		Thread.sleep(4000);
		
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(5000);
		
		int Count = viewsGridFiledsList.size();
		
		for (int i = 0; i < Count; i++) 
		{
			getAction().moveToElement(viewsGridFiledsList.get(i)).build().perform();
			String data		  = viewsGridFiledsList.get(i).getText();

    		if(data.equalsIgnoreCase("VendorAC"))
			
			{	
 
    			viewsGridFiledsEditList.get(i).click();
    			
				break;
			}
		}
		
		Thread.sleep(4000);
		
		click(fieldDetails_DefaultValue);
		fieldDetails_DefaultValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		fieldDetails_DefaultValue.sendKeys("102");
		click(fieldDetails_ApplyBtn);
		Thread.sleep(2000);
		ClickUsingJs(viewSaveView);
		
		String expMessage="Data saved successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
		
		Thread.sleep(8000);
		//getAction().moveToElement(updateBtn).build().perform();
		//Thread.sleep(2000);
		ClickUsingJs(updateBtn);
		Thread.sleep(8000);
		ClickUsingJs(settings_closeBtn);
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	public boolean checkSavingPurchseVoucherwithDefaultValueasGroup() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

		Thread.sleep(9999);
		
		Select s=new Select(multiViewsSelect);
		
		s.selectByVisibleText("Test View");
		
		click(multiViewsOkBtn);
		Thread.sleep(6000);
		click(newBtn);
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		Thread.sleep(2000);
		
		click(purchaseAccountTxt);
		purchaseAccountTxt.sendKeys("Purchase");
		Thread.sleep(1000);
		purchaseAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		click(vendorAccountTxt);
		
		vendorAccountTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		
		
		ArrayList<String>  actvendorAccountList= new ArrayList<String>(); 
		for(WebElement e: vendorAccountListCount)
		{
			actvendorAccountList.add(e.getText());
		}
		String actvendorAccountList1=actvendorAccountList.toString();
	
		String expvendorAccountList="[Group One Account One, Group One Account Two]";
	
		System.out.println("Actual		"	+	actvendorAccountList1);
		System.out.println("Expected	"	+	expvendorAccountList);
		Thread.sleep(2000);
		
		vendorAccountTxt.sendKeys("Group One Account One");
		vendorAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		warehouseTxt.sendKeys("Hyderabad");
		warehouseTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		departmentTxt.sendKeys("INDIA");
		Thread.sleep(4000);
		departmentTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		click(select1stRow_1stColumn);
		click(itemtxt);
		itemtxt.sendKeys("STD RATE COGS");
		Thread.sleep(2000);
		itemtxt.sendKeys(Keys.TAB);
		
		click(select1stRow_4thColumn);
		click(quantitytxt);
		quantitytxt.sendKeys("5");
		//quantitytxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		click(select1stRow_6thColumn);
		click(rateTxt);
		rateTxt.clear();
		rateTxt.sendKeys("5");
		Thread.sleep(2000);
		
		click(voucherSaveBtn);
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		Thread.sleep(2000);
	
		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		
		if(actMessage.startsWith(expMessage1) &&  actvendorAccountList1.equals(expvendorAccountList))
		{
			return true;
		}
		else
		{
			return false;
		}

	}
	
	
	public boolean checkSavingViewWithDefaultValueasGroupandSubGroup() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		Thread.sleep(2000);
		 click(entryPageToggleBtn);
		 Thread.sleep(2000);
		 
		 click(settingsOption);
		 Thread.sleep(2000);
		 
		
		  click(viewsBtn);
		  Thread.sleep(2000);
		  
		
		click(ViewExistingViewTxt);
		Thread.sleep(1500);
			
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			
		
		ViewExistingViewTxt.sendKeys("Test View");
		
		Thread.sleep(4000);
		
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(8000);
		
		int Count = viewsGridFiledsList.size();
		
		for (int i = 0; i < Count; i++) 
			
		{
			getAction().moveToElement(viewsGridFiledsList.get(i)).build().perform();
			String data		  = viewsGridFiledsList.get(i).getText();

    		if(data.equalsIgnoreCase("VendorAC"))
			
			{
    			viewsGridFiledsEditList.get(i).click();
				
				break;
			}
		}
		
		Thread.sleep(4000);
		
		click(fieldDetails_DefaultValue);
		fieldDetails_DefaultValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		fieldDetails_DefaultValue.sendKeys("39");
		click(fieldDetails_ApplyBtn);
		Thread.sleep(2000);
		ClickUsingJs(viewSaveView);
		
		String expMessage="Data saved successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
		Thread.sleep(8000);
		//getAction().moveToElement(updateBtn).build().perform();
		//Thread.sleep(2000);
		ClickUsingJs(updateBtn);
		Thread.sleep(10000);
		ClickUsingJs(settings_closeBtn);
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
	
	public boolean checkSavingPurchaseVoucherwithDefaultValuesofGroupandSubGroup() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		


		Thread.sleep(5000);
		
		Select s=new Select(multiViewsSelect);
		
		s.selectByVisibleText("Test View");
		
		click(multiViewsOkBtn);
		Thread.sleep(6000);
		click(newBtn);
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		Thread.sleep(2000);
		
		click(purchaseAccountTxt);
		purchaseAccountTxt.sendKeys("Purchase");
		Thread.sleep(1000);
		purchaseAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		click(vendorAccountTxt);
		
		vendorAccountTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		
		
		ArrayList<String>  actvendorAccountList= new ArrayList<String>(); 
		for(WebElement e: vendorAccountListCount)
		{
			actvendorAccountList.add(e.getText());
		}
		String actvendorAccountList1=actvendorAccountList.toString();
	
		String expvendorAccountList="[Vendor A, Vendor B]";
	
		System.out.println("Actual		"	+	actvendorAccountList1);
		System.out.println("Expected	"	+	expvendorAccountList);
		Thread.sleep(2000);
		
		vendorAccountTxt.sendKeys("Vendor A");
		vendorAccountTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		warehouseTxt.sendKeys("Hyderabad");
		warehouseTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		departmentTxt.sendKeys("INDIA");
		Thread.sleep(4000);
		departmentTxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		click(select1stRow_1stColumn);
		click(itemtxt);
		itemtxt.sendKeys("STD RATE COGS");
		Thread.sleep(2000);
		itemtxt.sendKeys(Keys.TAB);
		
		click(select1stRow_4thColumn);
		click(quantitytxt);
		quantitytxt.sendKeys("5");
		//quantitytxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		click(select1stRow_6thColumn);
		click(rateTxt);
		rateTxt.clear();
		rateTxt.sendKeys("5");
		Thread.sleep(2000);
		
		click(voucherSaveBtn);
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
		pickBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
		Bill_OkBtn.click();
		Thread.sleep(2000);
	
		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		
		if(actMessage.startsWith(expMessage1) &&  actvendorAccountList1.equals(expvendorAccountList))
		{
			return true;
		}
		else
		{
			return false;
		}

	
	}
	
	public boolean checkSavingViewforFooterFieldsWithSpecialCharacters() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

		Thread.sleep(5000);
		 click(entryPageToggleBtn);
		 Thread.sleep(2000);
		 
		 click(settingsOption);
		 Thread.sleep(2000);
		
		  click(viewsBtn);
		  Thread.sleep(2000);
		  
		
		click(ViewExistingViewTxt);
		Thread.sleep(1500);
			
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
			
		
		ViewExistingViewTxt.sendKeys("Test View");
		
		Thread.sleep(4000);
		
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(4000);
		
		int Count = viewsGridFiledsList.size();
		
		for (int i = 0; i < Count; i++) 
		{
			
			getAction().moveToElement(viewsGridFiledsList.get(i)).build().perform();
			String data		  = viewsGridFiledsList.get(i).getText();

    		if(data.equalsIgnoreCase("F1"))
			
			{
    			viewsGridFiledsEditList.get(i).click();
				
				break;
			}
		}
		
		Thread.sleep(4000);
		
		click(fieldDetailsF1_DefaultValue);
		fieldDetailsF1_DefaultValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		fieldDetailsF1_DefaultValue.sendKeys("&10#");
		
		getAction().moveToElement(fieldDetailsEditScreen_ApplyBtn).build().perform();
		Thread.sleep(2000);
		click(fieldDetailsEditScreen_ApplyBtn);
		Thread.sleep(2000);
		ClickUsingJs(viewSaveView);
		
		String expMessage="Data saved successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
		Thread.sleep(8000);
		//getAction().moveToElement(updateBtn).build().perform();
		
		//Thread.sleep(2000);
		ClickUsingJs(updateBtn);
		Thread.sleep(8000);
		//click(closeBtn);
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	
	
	}
	
	@FindBy(xpath="//*[@class='icon-insert-info-field vouchericon']")
	public static WebElement view_InserInfoField;
	
	@FindBy(xpath="//*[@class='icon-insert-info-column vouchericon']")
	public static WebElement view_InserInfoColumn;
	
	@FindBy(xpath="(//ul[@class='col-xs-12'])[1]/li/span")
	public static List<WebElement> view_InsertInfoFieldList;
	
	@FindBy(xpath="(//ul[@class='col-xs-12'])[4]/li/span")
	public static List<WebElement> view_InsertInfoFieldWarehouseList;
	
	@FindBy(xpath="//input[@value='Ok']")
	public static WebElement view_InserInfoFieldOkBtn;
	
	@FindBy(xpath="//input[@id='id_FooterClose']")
	public static WebElement view_InserInfoFieldCloseBtn;
	
	@FindBy(xpath="(//ul[@class='col-xs-12'])[1]/li/span")
	public static List<WebElement> view_InsertInfoColumnList;
	
	@FindBy(xpath="(//ul[@class='col-xs-12'])[4]/li/span")
	public static List<WebElement> view_InsertInfoColumnItemList;
	
	@FindBy(xpath="//input[@value='Ok']")
	public static WebElement view_InserInfoColOkBtn;
	
	
public boolean checkSavingViewWithInsertInformationField() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException	
{
	Thread.sleep(8000);
	click(ViewExistingViewTxt);
	Thread.sleep(4000);
		
	ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
	ViewExistingViewTxt.sendKeys("Test View");
	
	Thread.sleep(4000);
	
	ViewExistingViewTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(10000);
	
	 
    int Count = viewsGridFiledsList.size();
	
	for ( int i = 0; i < Count; i++) 
	{
		
		getAction().moveToElement(viewsGridFiledsList.get(i)).build().perform();
		String data		  = viewsGridFiledsList.get(i).getText();

		if(data.equalsIgnoreCase("Warehouse"))
		
		{
			viewsGridFiledsList.get(i).click();
			
			break;
		}
	}

	click(view_InserInfoField);
	Thread.sleep(8000);
	
int Count1 = view_InsertInfoFieldList.size();
	
	for ( int i = 0; i < Count1; i++) 
	{
		String data		  = view_InsertInfoFieldList.get(i).getText();

		if(data.equalsIgnoreCase("Warehouse"))
		
		{
			view_InsertInfoFieldList.get(i).click();
			
			break;
		}
	}
	
	
	Thread.sleep(2000);
	
int Count2 = view_InsertInfoFieldWarehouseList.size();
	
	for ( int i = 0; i < Count2; i++) 
	{
		getAction().moveToElement(view_InsertInfoFieldWarehouseList.get(i)).build().perform();
		Thread.sleep(2000);
		
		String data		  = view_InsertInfoFieldWarehouseList.get(i).getText();

		if(data.equalsIgnoreCase("Tax Code"))
		
		{
			view_InsertInfoFieldWarehouseList.get(i).click();
			
			break;
		}
	}
	
	click(view_InserInfoFieldOkBtn);
	Thread.sleep(2000);
	
	ClickUsingJs(view_InserInfoFieldCloseBtn);
	Thread.sleep(2000);
	
	String data1 = null;
	
	 int Count3 = viewsGridFiledsList.size();
		
		for ( int i = 0; i < Count3; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

			if(data.equalsIgnoreCase("Warehouse"))
			
			{
				data1  = viewsGridFiledsList.get(i + 1).getText();
				
				break;
			}
		}
	Thread.sleep(2000);
		System.out.println("After Insert Information Field	"	+	data1);
		ClickUsingJs(viewSaveView);
		
		String expMessage="Data saved successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
		
		Thread.sleep(4000);
		//getAction().moveToElement(updateBtn).build().perform();
		//Thread.sleep(2000);
		ClickUsingJs(updateBtn);
		Thread.sleep(8000);
		//Thread.sleep(2000);
		ClickUsingJs(settings_closeBtn);
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
}

@FindBy(xpath="(//*[@class='transactionentry_input_leftalign  Ftxtbox form-control'])[1]")
public static WebElement voucher_TaxcodeTxt;



	
public boolean checkSavingPurchaseVoucherwithInsertInfoFieldinView() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException, AWTException
{
	
	Thread.sleep(5000);
	
	prongHornStartAtAdminLevel();
	
	Select s=new Select(multiViewsSelect);
	
	s.selectByVisibleText("Test View");
	
	click(multiViewsOkBtn);
	Thread.sleep(6000);
	click(newBtn);
	Thread.sleep(4000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	Thread.sleep(2000);
	
	click(purchaseAccountTxt);
	purchaseAccountTxt.sendKeys("Purchase");
	Thread.sleep(1000);
	purchaseAccountTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
		
	click(vendorAccountTxt);
	
	vendorAccountTxt.sendKeys(Keys.SPACE);
	Thread.sleep(2000);
	
	vendorAccountTxt.sendKeys("Vendor A");
	Thread.sleep(2000);
	vendorAccountTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	warehouseTxt.sendKeys("Warehouse1");
	Thread.sleep(4000);
	warehouseTxt.sendKeys(Keys.TAB);
	//click(voucher_TaxcodeTxt);
	Thread.sleep(2000);
	
	
	String actTaxCode=voucher_TaxcodeTxt.getAttribute("value");
	String expTaxCode="W1@TaxCode1";
	
	System.out.println("Actual Tax Code			"		+	actTaxCode);
	System.out.println("Expected Tax Code		"		+	expTaxCode);
	
	click(departmentTxt);
	departmentTxt.sendKeys("INDIA");
	Thread.sleep(4000);
	departmentTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	click(select1stRow_1stColumn);
	click(itemtxt);
	itemtxt.sendKeys("STD RATE COGS");
	Thread.sleep(2000);
	itemtxt.sendKeys(Keys.TAB);
	
	click(select1stRow_4thColumn);
	click(quantitytxt);
	quantitytxt.sendKeys("5");
	//quantitytxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	click(select1stRow_6thColumn);
	click(rateTxt);
	rateTxt.clear();
	rateTxt.sendKeys("5");
	Thread.sleep(2000);
	
	click(voucherSaveBtn);
	Thread.sleep(6000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
	pickBtn.click();
	
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
	Bill_OkBtn.click();
	Thread.sleep(2000);

	
	String expMessage1 = "Voucher saved successfully";
	
	String actMessage = checkValidationMessage(expMessage1);
	
	if(actMessage.startsWith(expMessage1) && actTaxCode.equalsIgnoreCase(expTaxCode))
	{
		return true;
	}
	else
	{
		return false;
	}



}

@FindBy(xpath="//*[@id='id_editLayout_filtercriteria_grid_body']//tr[1]//td[2]")
public static WebElement view_Filter1stRow1stColumn;

@FindBy(xpath="//*[@id='id_editLayout_filtercriteria_grid_body']//tr[4]//td[7]")
public static WebElement view_Filter4thRow7thColumn;

@FindBy(xpath="//select[@id='id_editLayout_filtercriteria_filteroncolumn']")
public static WebElement view_FilterOnColumnSelect;


@FindBy(xpath="//select[@id='id_editLayout_filtercriteria_operator']")
public static WebElement view_FilterOnOperatorSelect;

@FindBy(xpath="//select[@id='id_editLayout_filtercriteria_comparewith']")
public static WebElement view_FilterCompareWithSelect;

@FindBy(xpath="//*[@id='editLayout_FieldsCustomization_FieldDetails_filterValTxtBx']")
public static WebElement view_FilterValue;

@FindBy(xpath="//*[@id='editLayout_FieldsCustomization_FieldDetails_linkFldTxtbx']")
public static WebElement view_FilterLinkField;

@FindBy(xpath="//select[@id='editLayout_FieldsCustomization_FieldDetails_conjuctionDrpDwn']")
public static WebElement view_FilterConjunction;

public boolean checkSavingViewWithMultipleRowsinFilterCriteria() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	Thread.sleep(5500);
	 click(entryPageToggleBtn);
	 Thread.sleep(2000);
	 
	 click(settingsOption);
	 Thread.sleep(2000);
	 
	
	  click(viewsBtn);
	  Thread.sleep(2000);
	  
	
	click(ViewExistingViewTxt);
	Thread.sleep(1500);
		
	ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
	ViewExistingViewTxt.sendKeys("Test View");
	
	Thread.sleep(4000);
	
	ViewExistingViewTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(4000);
	
	int Count = viewsGridFiledsList.size();
	
	for (int i = 0; i < Count; i++) 
	{
		String data		  = viewsGridFiledsList.get(i).getText();

		if(data.equalsIgnoreCase("PurchaseAc"))
		
		{
			viewsGridFiledsEditList.get(i).click();
			
			break;
		}
	}
	
	Thread.sleep(4000);
	
	ScrollIntoView(view_Filter1stRow1stColumn);
	click(view_Filter1stRow1stColumn);
	checkSavingFilterforRowsinViews("iMasterId","=","Value","74","OR");
	checkSavingFilterforRowsinViews("iMasterId","=","Value","9","OR");
	checkSavingFilterforRowsinViews("iMasterId","=","Value","7","OR");
	checkSavingFilterforRowsinViews("iMasterId","=","Value","11","OR");
	checkSavingFilterforRowsinViews("iMasterId","=","Value","88","OR");
	checkSavingFilterforRowsinViews("iMasterId","=","Value","12","ONLY");
	
	click(fieldDetails_ApplyBtn);
	Thread.sleep(6000);
	
	ClickUsingJs(viewSaveView);
	
	String expMessage="Data saved successfully";
	String actMessage=checkValidationMessage(expMessage);
	
	System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
	Thread.sleep(8000);
	
	//getAction().moveToElement(updateBtn).build().perform();
	//Thread.sleep(2000);
	
	ClickUsingJs(updateBtn);
	Thread.sleep(10000);
	
	ClickUsingJs(settings_closeBtn);
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
	

public void checkSavingFilterforRowsinViews(String onCol,String Operator,String value,String txt,String conjuction) throws InterruptedException
{
	
	Select s =new Select(view_FilterOnColumnSelect);
	s.selectByVisibleText(onCol);
	Thread.sleep(2000);
	view_FilterOnColumnSelect.sendKeys(Keys.TAB);
	
	Select s1 =new Select(view_FilterOnOperatorSelect);
	s1.selectByVisibleText(Operator);
	Thread.sleep(2000);
	view_FilterOnOperatorSelect.sendKeys(Keys.TAB);
	
	
	Select s2 =new Select(view_FilterCompareWithSelect);
	s2.selectByVisibleText(value);
	Thread.sleep(2000);
	view_FilterCompareWithSelect.sendKeys(Keys.TAB);
	
	view_FilterValue.sendKeys(txt);
	view_FilterValue.sendKeys(Keys.TAB);
	
	
	view_FilterLinkField.sendKeys(Keys.TAB);
	
	Select s3 =new Select(view_FilterConjunction);
	s3.selectByVisibleText(conjuction);
	Thread.sleep(2000);
	view_FilterConjunction.sendKeys(Keys.TAB);

}

@FindBy(xpath = "//tbody[@id='id_header_3_table_body']/tr/td[1]")
public static List<WebElement> purchaseAccountIdListCount;


public boolean checkSavedFilterCriteriainPurchaseVoucherforPurchaseAC() throws InterruptedException
{
	Thread.sleep(2000);
	
	Select s=new Select(multiViewsSelect);
	
	s.selectByVisibleText("Test View");
	
	click(multiViewsOkBtn);
	Thread.sleep(2000);
	click(newBtn);
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	Thread.sleep(2000);
	
	click(purchaseAccountTxt);

	purchaseAccountTxt.sendKeys(Keys.SPACE);
	Thread.sleep(2000);
	
	
	
	/* ArrayList<String> actpurchaseAccountIdList= new ArrayList<String>();
	  for(WebElement e: purchaseAccountIdListCount) {
		  
		  actpurchaseAccountIdList.add(e.getAttribute("value")); 
		  } 
	  String actpurchaseAccountIdList1=actpurchaseAccountIdList.toString();
	  
	  String exppurchaseAccountIdList="";
	  
	  System.out.println("Actual		" + actpurchaseAccountIdList1);
	  System.out.println("Expected		" + exppurchaseAccountIdList);
	 
	Thread.sleep(2000);
	*/
	
	
	  ArrayList<String> actpurchaseAccountList= new ArrayList<String>();
	  for(WebElement e: purchaseAccountListCount) {
	  actpurchaseAccountList.add(e.getText()); } String
	  actpurchaseAccountList1=actpurchaseAccountList.toString();
	  
	  String exppurchaseAccountList="[BR COGS ACC INV, Cost of goods sold - Computers, Cost of goods sold - Electronics, Inventory Trade, Inventory RM, Purchase]";
	  
	  System.out.println("Actual		" + actpurchaseAccountList1);
	  System.out.println("Expected		" + exppurchaseAccountList);
	 
	Thread.sleep(2000);
	
	

	
	if(actpurchaseAccountList1.equalsIgnoreCase(exppurchaseAccountList))
	{
		return true;
	}
	else
	{
		return false;
	}
}

/*
 * @FindBy(xpath =
 * "//tbody[@id='id_header_3_table_body']/tr/td[1]|//tbody[@id='id_header_3_table_body']/tr/td[2]")
 * public static List<WebElement> purchaseAccountListCount;
 */

public boolean checkSavingViewAfterEditingFilterCriteria() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(5500);
	 click(entryPageToggleBtn);
	 Thread.sleep(2000);
	 
	 click(settingsOption);
	 Thread.sleep(2000);
	
	// new WebDriverWait(getDriver(), 600).until(ExpectedConditions.alertIsPresent());
	 // getAlert().accept();
	 // Thread.sleep(2000);
	 
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucher_ConfirmYesBtn));
		voucher_ConfirmYesBtn.click();
		Thread.sleep(2000);
	
	
	  click(viewsBtn);
	  Thread.sleep(2000);
	  
	
	click(ViewExistingViewTxt);
	Thread.sleep(1500);
		
	ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
	ViewExistingViewTxt.sendKeys("Test View");
	
	Thread.sleep(4000);
	
	ViewExistingViewTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(4000);
	
	int Count = viewsGridFiledsList.size();
	
	for (int i = 0; i < Count; i++) 
	{
		String data		  = viewsGridFiledsList.get(i).getText();
		
		Thread.sleep(2000);
		
		

		if(data.equalsIgnoreCase("PurchaseAc"))
		
		{
			viewsGridFiledsEditList.get(i).click();
			
			break;
		}
		
	}
	
	Thread.sleep(4000);
	
	ScrollIntoView(view_Filter1stRow1stColumn);
	click(view_Filter4thRow7thColumn);
	Thread.sleep(2000);
	Select s3 =new Select(view_FilterConjunction);
	s3.selectByVisibleText("AND");
	Thread.sleep(2000);
	//view_FilterConjunction.sendKeys(Keys.TAB);
	
	click(fieldDetails_ApplyBtn);
	Thread.sleep(2000);
	
	ClickUsingJs(viewSaveView);
	
	String expMessage="Data saved successfully";
	String actMessage=checkValidationMessage(expMessage);
	
	System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
	
	Thread.sleep(8000);
	//getAction().moveToElement(updateBtn).build().perform();
	//Thread.sleep(2000);

	ClickUsingJs(updateBtn);
	Thread.sleep(8000);
	ClickUsingJs(settings_closeBtn);
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

public boolean checkEditedFilterCriteriainPurchaseVoucherforPurchaseAC() throws InterruptedException
{
	Thread.sleep(2000);
	
	Select s=new Select(multiViewsSelect);
	
	s.selectByVisibleText("Test View");
	Thread.sleep(2000);
	click(multiViewsOkBtn);
	Thread.sleep(8000);
	click(newBtn);
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	Thread.sleep(2000);
	
	click(purchaseAccountTxt);

	purchaseAccountTxt.sendKeys(Keys.SPACE);
	Thread.sleep(2000);
	
	
	
	 ArrayList<String> actpurchaseAccountIdList= new ArrayList<String>();
	  for(WebElement e: purchaseAccountIdListCount) {
		  
		  actpurchaseAccountIdList.add(e.getText()); 
		  } 
	  String actpurchaseAccountIdList1=actpurchaseAccountIdList.toString();
	  
	  String exppurchaseAccountIdList="[, , , ]";
	  
	  System.out.println("Actual		" + actpurchaseAccountIdList1);
	  System.out.println("Expected		" + exppurchaseAccountIdList);
	 
	Thread.sleep(2000);
	
	
	
	  ArrayList<String> actpurchaseAccountList= new ArrayList<String>();
	  for(WebElement e: purchaseAccountListCount) {
	  actpurchaseAccountList.add(e.getText()); } String
	  actpurchaseAccountList1=actpurchaseAccountList.toString();
	  
	  String exppurchaseAccountList="[Cost of goods sold - Computers, Cost of goods sold - Electronics, Inventory Trade, Purchase]";
	  
	  System.out.println("Actual		" + actpurchaseAccountList1);
	  System.out.println("Expected		" + exppurchaseAccountList);
	 
	Thread.sleep(2000);
	
	if(actpurchaseAccountList1.equalsIgnoreCase(exppurchaseAccountList))
	{
		return true;
	}
	else
	{
		return false;
	}
}

public boolean checkSavingViewWithInsertInformationColumn() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException	
{
	Thread.sleep(1500);
	
	Thread.sleep(2000);
	 click(entryPageToggleBtn);
	 Thread.sleep(2000);
	 
	 click(settingsOption);
	 Thread.sleep(2000);
	 
	

	  click(viewsBtn);
	  Thread.sleep(4000);
	  
	click(ViewExistingViewTxt);
	Thread.sleep(2000);
		
	ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		
	
	ViewExistingViewTxt.sendKeys("Test View");
	
	Thread.sleep(4000);
	
	ViewExistingViewTxt.sendKeys(Keys.TAB);
	
	Thread.sleep(8000);
	
	 
    int Count = viewsGridFiledsList.size();
	
	for ( int i = 0; i < Count; i++) 
	{
		getAction().moveToElement(viewsGridFiledsList.get(i)).build().perform();
		String data		  = viewsGridFiledsList.get(i).getText();

		if(data.equalsIgnoreCase("Item"))
		
		{
			viewsGridFiledsList.get(i).click();
			
			break;
		}
	}

	click(view_InserInfoColumn);
	Thread.sleep(6000);
	
int Count1 = view_InsertInfoColumnList.size();
	
	for ( int i = 0; i < Count1; i++) 
	{
		
		String data		  = view_InsertInfoColumnList.get(i).getText();

		if(data.equalsIgnoreCase("Item"))
		
		{
			view_InsertInfoColumnList.get(i).click();
			
			break;
		}
	}
	
	
int Count2 = view_InsertInfoColumnItemList.size();
	
	for ( int i = 0; i < Count2; i++) 
	{
		
		getAction().moveToElement(view_InsertInfoColumnItemList.get(i)).build().perform();
		String data		  = view_InsertInfoColumnItemList.get(i).getText();

		if(data.equalsIgnoreCase("TaxCode"))
		
		{
			view_InsertInfoColumnItemList.get(i).click();
			
			break;
		}
	}
	
	
	Thread.sleep(4000);
	
	click(view_InserInfoColOkBtn);
	Thread.sleep(4000);
	
	ClickUsingJs(view_InserInfoFieldCloseBtn);
	Thread.sleep(2000);
	
	
	String data1 = null;
	
	 int Count3 = viewsGridFiledsList.size();
		
		for ( int i = 0; i < Count3; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

			if(data.equalsIgnoreCase("Item"))
			
			{
				data1  = viewsGridFiledsList.get(i + 1).getText();
				
				break;
			}
		}
	Thread.sleep(2000);
		System.out.println("After Insert Information Field	"	+	data1);
		ClickUsingJs(viewSaveView);
		
		String expMessage="Data saved successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
		Thread.sleep(2000);
		//getAction().moveToElement(updateBtn).build().perform();
		//Thread.sleep(2000);
		ClickUsingJs(updateBtn);
		Thread.sleep(2000);
		ClickUsingJs(settings_closeBtn);
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			return true;
		}
		else
		{
			return false;
		}
}

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[3]")
public static WebElement voucherBody_TaxcodeTxt;

	
public boolean checkSavingPurchaseVoucherwithInsertInfoColumninView() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
{
	
	Thread.sleep(5000);
	
	Select s=new Select(multiViewsSelect);
	
	s.selectByVisibleText("Test View");
	
	click(multiViewsOkBtn);
	Thread.sleep(2000);
	click(newBtn);
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	Thread.sleep(2000);
	
	click(purchaseAccountTxt);
	purchaseAccountTxt.sendKeys("Purchase");
	Thread.sleep(1000);
	purchaseAccountTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	click(vendorAccountTxt);
	
	vendorAccountTxt.sendKeys(Keys.SPACE);
	Thread.sleep(2000);
	
	vendorAccountTxt.sendKeys("Vendor A");
	vendorAccountTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	warehouseTxt.sendKeys("Hyderabad");
	Thread.sleep(2000);
	warehouseTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	
	
	click(departmentTxt);
	departmentTxt.sendKeys("INDIA");
	Thread.sleep(4000);
	departmentTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	click(select1stRow_1stColumn);
	click(itemtxt);
	itemtxt.sendKeys("Item1");
	Thread.sleep(3000);
	itemtxt.sendKeys(Keys.TAB);
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherBody_TaxcodeTxt));
	String actTaxCode=voucherBody_TaxcodeTxt.getText();
	String expTaxCode="Taxcode1";
	
	System.out.println(voucherBody_TaxcodeTxt.getAttribute("data-value"));
	
	System.out.println("Actual Tax Code			"		+	actTaxCode);
	System.out.println("Expected Tax Code		"		+	expTaxCode);
	
	Thread.sleep(2000);
	click(select1stRow_5thColumn);
	click(quantitytxt);
	quantitytxt.sendKeys("5");
	//quantitytxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	click(select1stRow_6thColumn);
	click(rateTxt);
	rateTxt.clear();
	rateTxt.sendKeys("5");
	Thread.sleep(2000);
	
	click(voucherSaveBtn);
	Thread.sleep(4000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pickBtn));
	pickBtn.click();
	
	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Bill_OkBtn));
	Bill_OkBtn.click();
	Thread.sleep(2000);

	
	String expMessage1 = "Voucher saved successfully";
	
	String actMessage = checkValidationMessage(expMessage1);
	
	if(actMessage.startsWith(expMessage1) && actTaxCode.equalsIgnoreCase(expTaxCode))
	{
		return true;
	}
	else
	{
		return false;
	}



}

@FindBy(xpath="//*[@class='icon-settings hiconright2']")
public static WebElement SR_SettingsBtn;


public boolean checkSavingViewwithDefaultValueinSalesReturnsVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	Thread.sleep(2000);
	click(serachMenuHomePage);
	Thread.sleep(1000);
	
	serachMenuTextHomePage.sendKeys("Sales returns");
	Thread.sleep(1000);
	click(searchMenuTextClick);
	 
	Thread.sleep(8000);
	
	if(getIsAlertPresent())
	{
		
		getWaitForAlert();
		getAlert().accept();
	}
	
	Thread.sleep(8000);
	
	click(SR_SettingsBtn);
		
	 Thread.sleep(6000);
	  click(viewsBtn);
	  Thread.sleep(4000);
	
	click(ViewExistingViewTxt);

	ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	

	Thread.sleep(2000);
	ViewExistingViewTxt.sendKeys("Sales View");
	Thread.sleep(4000);
	ViewExistingViewTxt.sendKeys(Keys.TAB);
	
			
    Thread.sleep(10000);
    
    
  // getAction().moveToElement(viewSettingsBtn).build().perform();
	//Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSettingsBtn));
	ClickUsingJs(viewSettingsBtn);
	Thread.sleep(2000);
	
if(viewSuChkboxIsSelected.isDisplayed()==false)
	{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
	viewTabLoginBtn.click();
	}
	Thread.sleep(2000);

	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
	   if(viewsUserAllOptionsSTChkboxIsSelected.isSelected()==false)
	   {
		   viewsUserAllOptionsSTChkbox.click();
	   }
		
	   Thread.sleep(2000);
	
   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
   ClickUsingJs(viewSaveView);
   Thread.sleep(5500);

   
   
   checkValidationMessage("");
   
   click(ViewExistingViewTxt);
	
	ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	

	Thread.sleep(2000);
	ViewExistingViewTxt.sendKeys("Sales View");
	Thread.sleep(4000);
	ViewExistingViewTxt.sendKeys(Keys.TAB);

	
   Thread.sleep(10000);
 
	int Count = viewsGridFiledsList.size();

	for (int i = 0; i < Count; i++) 
	{
		getAction().moveToElement(viewsGridFiledsList.get(i)).build().perform();
		
		String data		  = viewsGridFiledsList.get(i).getText();

		if(data.equalsIgnoreCase("CustomerAc"))
		
		{
			viewsGridFiledsEditList.get(i).click();
			
			break;
		}
	}
	
	Thread.sleep(4000);
   
  
	click(fieldDetails_DefaultValue);
	fieldDetails_DefaultValue.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	fieldDetails_DefaultValue.sendKeys("#19");
	click(fieldDetails_ApplyBtn);
	Thread.sleep(4000);
	ClickUsingJs(viewSaveView);
	
	String expMessage="Data saved successfully";
	String actMessage=checkValidationMessage(expMessage);
	
	System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
	Thread.sleep(8000);
	//getAction().moveToElement(updateBtn).build().perform();
	//Thread.sleep(2000);

	ClickUsingJs(updateBtn);
	Thread.sleep(12000);
	ClickUsingJs(settings_closeBtn);
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

public boolean checkLogoutandLoginwithUseralloptionsst() throws InterruptedException
{
	getDriver().navigate().refresh();
	Thread.sleep(2000);
	click(userNameDisplayLogo);
	Thread.sleep(2000);
	click(logoutOption);
	
	Thread.sleep(3000);

	LoginPage lp = new LoginPage(getDriver());

	String unamelt = "UserAllOptionsST";

	String pawslt = "12345";

	lp.enterUserName(unamelt);

	lp.enterPassword(pawslt);

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

	
	Thread.sleep(12000);

	String actUserInfo = userNameUserDisplay.getText();

	System.out.println("User Info  : " + actUserInfo);

	System.out.println("User Info Capture Text  :  " + userNameUserDisplay.getText());

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
	//companyLogo.click();

	String getCompanyTxt = Company_Name.getText();
	String getLoginCompanyName = getCompanyTxt.substring(0, 19);
	System.out.println("company name  :  " + getLoginCompanyName);
	//companyLogo.click();

	String expUserInfo = "UserAllOptionsST";
	String expLoginCompanyName = "Automation";

	boolean actDashboard = dashboard.isDisplayed();
	boolean expDashboard = true;

	System.out.println(
			"**********************************************checkLogoutAndLoginWithUserAllOptions*********************************************");

	System.out.println("User Info               : " + actUserInfo + " Value Expected " + expUserInfo);
	System.out
			.println("Login Company           : " + getLoginCompanyName + " Value Expected " + expLoginCompanyName);
	System.out.println("Dashboard is Displaying : " + actDashboard + " Value Expected " + expDashboard);

	if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName)
			&& actDashboard == expDashboard) {
		return true;
	} else {
		return false;
	}
}

@FindBy(xpath="//*[@id='id_transaction_homescreen_new']")
public static WebElement SR_NewBtn;

@FindBy(xpath="//*[@id='id_header_268435459']")
public static WebElement departmenttxt;	


@FindBy(xpath="//*[@id='id_body_536870916']")
public static WebElement warehousetxt;	


	public boolean checkSavingSalesReturnsVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		click(serachMenuHomePage);
		
		serachMenuTextHomePage.sendKeys("Sales returns");
		click(searchMenuTextClick);
		Thread.sleep(2000);
		
		click(SR_NewBtn);
		Thread.sleep(2000);
		
		click(customerAccountTxt);
		String actCustomerTxt=customerAccountTxt.getAttribute("value");
		String expCustomerTxt="Customer A";
		
		System.out.println("Actual Customer Acc		"	+	actCustomerTxt);
		System.out.println("Expected Customer Acc	"	+	expCustomerTxt);
		
		Thread.sleep(2000);
		click(departmenttxt);
		Thread.sleep(2000);
		departmenttxt.sendKeys("INDIA");
		Thread.sleep(4000);
		departmenttxt.sendKeys(Keys.TAB);
		
		click(select1stRow_1stColumn);
		click(warehousetxt);
		warehousetxt.sendKeys("Hyderabad");
		Thread.sleep(2000);
		warehousetxt.sendKeys(Keys.TAB);
		
		itemtxt.sendKeys("STD RATE COGS ITEM");
		Thread.sleep(2000);
		itemtxt.sendKeys(Keys.TAB);
		
		click(select1stRow_6thColumn);
		click(quantitytxt);
		quantitytxt.sendKeys("10");
		
		click(select1stRow_8thColumn);
		click(rateTxt);
		rateTxt.clear();
		rateTxt.sendKeys("2.5");
		
		click(voucherSaveBtn);
		Thread.sleep(10000);
		
		click(pickBtn);
		
		Thread.sleep(2000);

		click(Bill_OkBtn);
		
		Thread.sleep(2000);

		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		
		if(actMessage.startsWith(expMessage1) && actCustomerTxt.equalsIgnoreCase(expCustomerTxt))
		{
			return true;
		}
		else
		{
			return false;
		}
		
		
	}
	
	
	public static boolean checkLogoutAndLoginWithSU() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		click(userNameDisplayLogo);
	
		Thread.sleep(2000);
		click(logoutOption);
	
		Thread.sleep(3000);

		LoginPage lp = new LoginPage(getDriver());

		String unamelt = "su";

		String pawslt = "su";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

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

		

		Thread.sleep(6000);

		String actUserInfo = userNameDisplay.getText();

		System.out.println("User Info  : " + actUserInfo);

		System.out.println("User Info Capture Text  :  " + userNameDisplay.getText());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		//companyLogo.click();

		String getCompanyTxt = Company_Name.getText();
		String getLoginCompanyName = getCompanyTxt.substring(0, 19);
		System.out.println("company name  :  " + getLoginCompanyName);
		//companyLogo.click();

		String expUserInfo = "su";
		String expLoginCompanyName = "Automation";

		boolean actDashboard = dashboard.isDisplayed();
		boolean expDashboard = true;

		System.out.println(
				"**********************************************checkLogoutAndLoginWithSU*********************************************");

		System.out.println("User Info               : " + actUserInfo + " Value Expected " + expUserInfo);
		System.out
				.println("Login Company           : " + getLoginCompanyName + " Value Expected " + expLoginCompanyName);
		System.out.println("Dashboard is Displaying : " + actDashboard + " Value Expected " + expDashboard);

		if (actUserInfo.equalsIgnoreCase(expUserInfo) && getLoginCompanyName.contains(expLoginCompanyName)
				&& actDashboard == expDashboard) {
			return true;
		} else {
			return false;
		}
	}

	
	
	public boolean checkSavingOpeningStockVoucher() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		donotuserealtimerate();
		Thread.sleep(4000);
		click(inventoryMenu);
		click(inventoryTransactionsMenu);
		click(inventoryTransactionsStocksMenu);
		click(openingStocksVoucher);
		
		Thread.sleep(2000);

		click(newBtn);
		
		//checkValidationMessage("Screen opened");


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));

		click(warehouseTxt);
		warehouseTxt.sendKeys(Keys.SPACE);

		int warehouselist=warehouseHeaderComboList.size();

		for (int i = 0; i < warehouselist; i++) 
		{
			String warehouse=warehouseHeaderComboList.get(i).getText();

			if (warehouse.equalsIgnoreCase("HYDERABAD")) 
			{
				warehouseHeaderComboList.get(i).click();
				warehouseTxt.sendKeys(Keys.TAB);
				break;
			}
		}

		click(select1stRow_1stColumn);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.SPACE);

		int itemsCount=itemComboList.size();

		for (int i = 0; i < itemsCount; i++) 
		{
			String item=itemComboList.get(i).getText();

			if (item.equalsIgnoreCase("STD RATE COGS ITEM")) 
			{
				itemComboList.get(i).click();
				enter_ItemTxt.sendKeys(Keys.TAB);
				break;
			}
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys("10");
		enter_Quantity.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.clear();
		enter_Rate.sendKeys("5");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherSaveBtn));
		voucherSaveBtn.click();

		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": 1";
		
		
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		if(actMessage.startsWith(expMessage1) || actMessage.endsWith(expMessage2) )
		{
			return true;
		}
		else
		{
			return false;
		}

		
	}
	
	
	@FindBy(xpath = "//*[@id='EditLayout_FieldsCustomization_FieldDetails_DataType']")
	public static WebElement fieldDetailsEditLayout_DataTypeDrpdwn;
	
	
	public boolean checkSavingViewinSalesInvoiceVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(2000);
		click(financialsMenu);
		
		click(financialsTransactionMenu);
		
		click(financialTransactionSalesMenu);
		
		click(salesInvoiceVATVoucher);
		
		
		Thread.sleep(4000);
	
		click(SR_SettingsBtn);
		
		 Thread.sleep(2000);
		
		  click(viewsBtn);
		  Thread.sleep(2000);
		  
		
		Thread.sleep(2000);
		click(ViewExistingViewTxt);

		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		

		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys("Sales Invoice View");
		Thread.sleep(4000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
				
	    Thread.sleep(8000);

	    
	    
	   getAction().moveToElement(viewSettingsBtn).build().perform();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSettingsBtn));
		viewSettingsBtn.click();
		Thread.sleep(2000);
		
	if(viewSuChkboxIsSelected.isDisplayed()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		}
		Thread.sleep(2000);

		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		   if(viewsUserAllOptionsSTChkboxIsSelected.isSelected()==false)
		   {
			   viewsUserAllOptionsSTChkbox.click();
		   }
			
		   Thread.sleep(2000);
		
		//getAction().moveToElement(viewSaveView).build().perform();
		
	 
	   
	  // getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
	   ClickUsingJs(viewSaveView);
	   Thread.sleep(4500);

	   
	   
	   checkValidationMessage("");
	   click(ViewExistingViewTxt);
		
		ViewExistingViewTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		

		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys("Sales Invoice View");
		Thread.sleep(4000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);

		
	   Thread.sleep(8000);
	 
		int Count = viewsGridFiledsList.size();

		for (int i = 0; i < Count; i++) 
		{
			getAction().moveToElement(viewsGridFiledsList.get(i)).build().perform();
			String data		  = viewsGridFiledsList.get(i).getText();

			if(data.equalsIgnoreCase("Item"))
			
			{
				viewsGridFiledsEditList.get(i).click();
				
				break;
			}
		}
		
		Thread.sleep(4000);
	   
	
		click(fieldDetailsEditLayout_DataTypeDrpdwn);
		Select s=new Select(fieldDetailsEditLayout_DataTypeDrpdwn);
		s.selectByVisibleText("Read Only");
		click(fieldDetails_ApplyBtn);
		Thread.sleep(2000);
		ClickUsingJs(viewSaveView);
		
		String expMessage="Data saved successfully";
		String actMessage=checkValidationMessage(expMessage);
		
		System.out.println("Actual		"		+	actMessage	+	"Expected	"	+	expMessage);
		Thread.sleep(8000);
		
		//getAction().moveToElement(updateBtn).build().perform();
		//Thread.sleep(2000);
		ClickUsingJs(updateBtn);
		Thread.sleep(4000);
		ClickUsingJs(settings_closeBtn);
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
	
	
	@FindBy(xpath="//td[@id='Id_StockReservation_FGrid_col_1-4']")
	private static WebElement stockRes_QtyToReserveRow1;
	

	@FindBy(xpath="//li[@id='btnPickQuantity']")
	private static WebElement res_PickBtn;

	@FindBy(xpath="//div[contains(text(),'Discard')]")
	private static WebElement res_DiscardBtn;

	@FindBy(xpath="(//*[contains(text(),'Ok')])[2]")
	private static WebElement res_OkBtn;

	@FindBy(xpath="//div[@class='col-sm-12 footer_button_image']//span[@class='icon-close icon-font6']")
	private static WebElement res_CloseBtn;

	@FindBy(xpath = "//tbody[@id='id_header_4_table_body']/tr/td[2]")
	public static List<WebElement> customerAccountListCount;
	
	@FindBy(xpath = "//tbody[@id='id_body_536870916_table_body']/tr/td[2]")
	public static List<WebElement> pvwareHouseListCount;

	
	@FindBy(xpath = "//tbody[@id='id_body_23_table_body']/tr/td[2]")
	public static List<WebElement> itemListCount;
	
	public boolean checkSavingSalesOrderVoucherwithStockItem() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{

		Thread.sleep(4000);
		click(inventoryMenu);
		Thread.sleep(2000);
		click(inventoryTransactionsMenu);
		Thread.sleep(2000);
		click(inventoryTransactionsSalesMenu);
		Thread.sleep(2000);
		click(salesOrdersVoucher);
		Thread.sleep(10000);
		click(newBtn);
		Thread.sleep(4000);
		
		//checkValidationMessage("Screen opened");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));


		click(customerAccountTxt);
		customerAccountTxt.sendKeys(Keys.SPACE);

		int customercount=customerAccountListCount.size();

		System.err.println(customercount);

		for(int i=0 ; i < customercount ;i++)
		{
			String data=customerAccountListCount.get(i).getText();

			if(data.equalsIgnoreCase("Customer A"))
			{
				customerAccountListCount.get(i).click();
				break;
			}
		}

		customerAccountTxt.sendKeys(Keys.TAB);
		
		click(select1stRow_1stColumn);

		click(pvWareHouseTxt);

		pvWareHouseTxt.sendKeys(Keys.SPACE);

		int warehousecount=pvwareHouseListCount.size();

		System.err.println(warehousecount);

		for(int i=0 ; i < warehousecount ;i++)
		{
			String data=pvwareHouseListCount.get(i).getText();

			if(data.equalsIgnoreCase("Hyderabad"))
			{
				pvwareHouseListCount.get(i).click();

				break;
			}
		}

		pvWareHouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.SPACE);

		int itemcount1=itemListCount.size();

		System.err.println(itemcount1);

		for(int i=0 ; i < itemcount1 ;i++)
		{
			String data=itemListCount.get(i).getText();

			if(data.equalsIgnoreCase("STD RATE COGS ITEM"))
			{
				itemListCount.get(i).click();

				break;
			}
		}
		
		Thread.sleep(2000);

		enter_ItemTxt.sendKeys(Keys.TAB);	

		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		//enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
		select1stRow_7thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys("6");
		enter_Quantity.sendKeys(Keys.TAB);

		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockRes_QtyToReserveRow1));
		stockRes_QtyToReserveRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_PickBtn));
		res_PickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));

		String actReserveColumn	=select1stRow_8thColumn.getText();
		String expReserveColumn	="Reserved";

		System.out.println("Reservation Column  : "+actReserveColumn+"  Value Expected  "+expReserveColumn);

			
		click(select2ndRow_1stColumn);
		Thread.sleep(2000);

		click(pvWareHouseTxt);
		
		pvWareHouseTxt.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
		enter_ItemTxt.sendKeys(Keys.SPACE);

		int itemcount2=itemListCount.size();

		System.err.println(itemcount1);

		for(int i=0 ; i < itemcount2 ;i++)
		{
			String data=itemListCount.get(i).getText();

			if(data.equalsIgnoreCase("STD RATE COGS ITEM"))
			{
				itemListCount.get(i).click();

				break;
			}
		}

		enter_ItemTxt.sendKeys(Keys.TAB);	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_UnitTxt));
		enter_UnitTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_7thColumn));
		select2ndRow_7thColumn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Quantity));
		enter_Quantity.sendKeys("6");
		enter_Quantity.sendKeys(Keys.TAB);


		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockRes_QtyToReserveRow1));
		stockRes_QtyToReserveRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_PickBtn));
		res_PickBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(res_OkBtn));
		res_OkBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
		enter_Rate.click();
		enter_Rate.clear();
		enter_Rate.sendKeys("10");
		enter_Rate.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Gross));
		enter_Gross.sendKeys(Keys.TAB);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(MRsaveBtn));
		MRsaveBtn.click();

		
		String expMessage1 = "Voucher saved successfully";
		
		String actMessage = checkValidationMessage(expMessage1);
		String expMessage2 = ": 1";
		
				
		System.out.println("Actual Message    : "+actMessage);
		System.out.println("Expected Message  : "+expMessage1);

		
		
		if (actReserveColumn.equalsIgnoreCase(expReserveColumn) && actMessage.startsWith(expMessage1) && actMessage.endsWith(expMessage2) ) 
		{
			return true;
		} 
		else 
		{
			return false;
		}
	
		
	}
	
	@FindBy(xpath="//*[@id='panelsStayOpen-headingOne']")
	public static WebElement SI_headerExpandBtn;
	
	
	@FindBy(xpath="//a[contains(text(),'Pending Sales Orders')]")
	public static WebElement pendingSOLink;
	
	
	@FindBy(xpath="//*[@id='trRender_1']/td[2]/input")
	public static WebElement pendingSOChkBox;
	
	
	@FindBy(xpath="//*[@id='id_transaction_homescreen_convert']")
	public static WebElement SI_ConvertBtn;
	
	public boolean checkSavingSalesInvoiceVATthroughConvertOption() throws InterruptedException
	{
		
		Thread.sleep(2000);
		click(financialsMenu);
		Thread.sleep(2000);
		click(financialsTransactionMenu);
		Thread.sleep(2000);
		click(financialTransactionSalesMenu);
		Thread.sleep(2000);
		click(salesInvoiceVATVoucher);
		Thread.sleep(4000);
		
		click(SI_headerExpandBtn);
		Thread.sleep(2000);
		click(pendingSOLink);
		Thread.sleep(2000);
		click(pendingSOChkBox);
		Thread.sleep(2000);
		click(SI_ConvertBtn);
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItemEdit=enter_ItemTxt.getAttribute("disabled");
		String expItemEdit="true";
		
		System.out.println("Item Editing	"		+	actItemEdit		+	"Expected	"	+	expItemEdit);
		
		Thread.sleep(2000);
		
		click(select1stRow_12thColumn);
		//click(quantitytxt);
		quantitytxt.sendKeys(Keys.TAB);
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_13thColumn));

		String actSalesColumn	=select1stRow_13thColumn.getAttribute("data-value");
		String expSalesColumn	="SalOrd:1";

		System.out.println("Sales Order Column  : "+actSalesColumn+"  Value Expected  "+expSalesColumn);

		if(actItemEdit.equalsIgnoreCase(expItemEdit) && actSalesColumn.equalsIgnoreCase(expSalesColumn))
		{
		
			return true;
		}
		else
		{
			return false;
		}

		
	}
	
	
	public boolean checkLogoutVoucherViewsPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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


	
	
	 public VoucherViewsPage(WebDriver driver)
	 {
	    	
	    	PageFactory.initElements(driver, this);
	    	
	 }
	

}
