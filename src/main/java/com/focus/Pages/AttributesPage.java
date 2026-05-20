package com.focus.Pages;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import com.focus.Pages.LoginPage;
import com.focus.base.BaseEngine;
import com.focus.utilities.DriverUtility;

public class AttributesPage extends BaseEngine{
	
	
	//SettingsMenu
	@FindBy(xpath="//*[@id='16']/div/span")
	public static WebElement settings;
//config transactions
	@FindBy(xpath="//*[@id='17']/span")
	public static WebElement configTrans;
	
//config transactions preferences
	@FindBy(xpath="(//a[@id='16' and //text()='Masters'])[2]")
	public static WebElement PreferencesMaster;

	//Label of Masters tab
	@FindBy(xpath="//*[@id='ScreenTitle']")
	public static WebElement LblMaster;

	@FindBy(xpath="//*[@id='ProductAttributesValues']")
	public static WebElement attributes;

	@FindBy(xpath="//*[@id='grdProductAttributesValues_body']/tr[1]/td[3]")
	public static WebElement attribute0;
	
	@FindBy(xpath="//*[@id='txtProductName']")
	public static WebElement attributeAfterclick;
	
	@FindBy(xpath="//*[@id='grdProductAttributesValues_body']/tr[2]/td[3]")
	public static WebElement attribute1;
	
	@FindBy(xpath="//*[@id='grdProductAttributesValues_body']/tr[3]/td[3]")
	public static WebElement attribute2;
	
	@FindBy(xpath="//*[@id='grdProductAttributesValues_body']/tr[4]/td[3]")
	public static WebElement attribute3;

	//update 
	@FindBy(xpath="//*[@id='updateButton']/i")
	public static WebElement updtBtn;
	
	//close
	@FindBy(xpath="//*[@id='btnCustomizeClose']")
	public static WebElement closeBtn;
	
	@FindBy(xpath="//*[@id='Dashboard_ddlList']/li/a")
	public static WebElement dashboard;
	
	
	//Home Menu
	@FindBy (xpath="//*[@id='1']")
    public static WebElement homeMenu;
	
        //Masters Menu	
		@FindBy (xpath="//a[@id='1000']//span[contains(text(),'Masters')]")
	    public static WebElement mastersMenu;
		
		//Items Menu
		@FindBy(xpath="//*[@id='221']")
	public static WebElement itemMenu;
		
		@FindBy(xpath="//*[@id='1105']")
		public static WebElement item;
	
		
		
		@FindBy(xpath="//*[@id='spnHeaderText']")
		public static WebElement LblItem;
		
		@FindBy(xpath="//*[@id='btnNew']/i")
		public static WebElement newBtn;
		
		@FindBy(xpath="//*[@id='divBtnGroup10']")
		public static WebElement attribute0Tab;
		
		@FindBy(xpath="//*[@id='divBtnGroup11']")
		public static WebElement attribute1Tab;
		
		@FindBy(xpath="//*[@id='divBtnGroup12']")
		public static WebElement attribute2Tab;
		
		@FindBy(xpath="//*[@id='divBtnGroup13']")
		public static WebElement attribute3Tab;
		
		
		public static boolean checkLogin() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		        Thread.sleep(3000);

		        getDriver().navigate().refresh();
		        Thread.sleep(3000);

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

		       /* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
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
	
	public boolean checkConfigureTransactionMastersUnderSettings() throws InterruptedException
	{
		
		Thread.sleep(2000);
		ClickUsingJs(SettingsMenu);
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(preferencesBtn));
		preferencesBtn.click();

		ClickUsingJs(PreferencesMaster);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(LblMaster));
		
		String actLabelMaster=LblMaster.getText();
		String expLabelMaster="Masters [Preferences]";
		
		System.out.println("actLabelMaster: "+actLabelMaster);
		System.out.println("expLabelMaster: "+expLabelMaster);

		if (actLabelMaster.equalsIgnoreCase(expLabelMaster)) 
		{
			return true;
		} 
		else 
		{
			return false;
		}
	}

	
	
	
	public boolean checkItemAttributeValuesinMastersUnderSettingsConfigureTransactions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(attributes));
		Select select=new Select(attributes);
		select.selectByVisibleText("4");
		
		//Entering first attribute
		
		ClickUsingJs(attribute0);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(attributeAfterclick));
		attributeAfterclick.clear();
		attributeAfterclick.sendKeys("Colors");
		
		//Entering second attribute 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(attribute1));
		attribute1.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(attributeAfterclick));
		attributeAfterclick.clear();
		attributeAfterclick.sendKeys("Fruits");
		
		Thread.sleep(2000);
		
		//Entering Third attribute 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(attribute2));
		attribute2.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(attributeAfterclick));
		attributeAfterclick.clear();
		attributeAfterclick.sendKeys("Flowers");
		
		//Entering fourth attribute 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(attribute3));
		attribute3.click();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(attributeAfterclick));
		attributeAfterclick.clear();
		attributeAfterclick.sendKeys("Electronics");
		
		Thread.sleep(2000);
		
		//clicking update button
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updtBtn));
		updtBtn.click();
		
		getWaitForAlert();
		
		String actAlert=getAlert().getText();
		String expAlrert="Do you want to save the changes?";
		
		System.out.println("actAlert : "+actAlert);
		System.out.println("expAlrert: "+expAlrert);
		
		getAlert().accept();

		String expMessage = "Data saved Successfully";
		String actMessage = checkValidationMessage(expMessage);
		
		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		}
		else 
		{
			return false;
		}
		
	}
		public boolean checkCloseOptionInMastersScreenUnderSettingsMenu() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			//clicking on close
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
			closeBtn.click();
			
			Thread.sleep(1000);
			
			
			boolean actMessage=dashboard.isDisplayed();
			
			if(actMessage)
			{
				return true;
			}
			else
			{
				return false;
			}
		}

		
		@FindBy(xpath="//*[@id='lastTab']")
		public static WebElement itemLasttab;
	
		@FindBy(xpath="//li[@class='liProdAttribute']/a")
		public static List<WebElement> attributesList;
		

	public boolean CheckAttributesTabinItemMenu() throws EncryptedDocumentException, InvalidFormatException, InterruptedException, IOException
	{
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemMenu));
		itemMenu.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item));
		item.click();
		
		Thread.sleep(4000);
		
	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
		newBtn.click();
		
		Thread.sleep(1000);
		
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Colors Attribute Item");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("Colors Attribute Item");
		 codeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLasttab));
		itemLasttab.click();
		Thread.sleep(1000);
		
		ArrayList<String>attributesName=new ArrayList<String>();
		
		
		for(WebElement e: attributesList)
		{
			attributesName.add(e.getText());
			
		}
		
		String actAttributesNameList=attributesName.toString();
		String expAttributesNameList="[Attribute-0, Attribute-1, Attribute-2, Attribute-3]";
		
		System.out.println("Actual 	 Attributes List  "	 +actAttributesNameList);
		System.out.println("Expected Attributes List  "	 +expAttributesNameList);
		
		
		if(actAttributesNameList.equalsIgnoreCase(expAttributesNameList))
		
		{
			return true;
		}
		else
			return false;
	}
	

		
	@FindBy(xpath="//*[@id='chkProductAttribute0']")
	public static WebElement checkBox0;
	
	@FindBy(xpath="//*[@id='chkProductAttribute1']")
	public static WebElement checkBox1;
	
	@FindBy(xpath="//*[@id='chkProductAttribute2']")
	public static WebElement checkBox2;
	
	@FindBy(xpath="//*[@id='chkProductAttribute3']")
	public static WebElement checkBox3;
	
	
	@FindBy(xpath="//*[@id='txtProdAttributeName0']")
	public static WebElement Valueattribute1;
	
	
	@FindBy(xpath="//*[@id='txtProdAttributeCode0']")
	public static WebElement Codeattribute1;
	
	//*[@id="txtProdAttributeName1"]
	
	@FindBy(xpath="//*[@id='txtProdAttributeName1']")
	public static WebElement Valueattribute2;
		
	@FindBy(xpath="//*[@id='txtProdAttributeCode1']")
	public static WebElement Codeattribute2;

	
	@FindBy(xpath="//*[@id='txtProdAttributeName3']")
	public static WebElement Valueattribute3;
	
	@FindBy(xpath="//*[@id='txtProdAttributeCode3']")
	public static WebElement Codeattribute3;
	
	
	@FindBy(xpath="//*[@id='txtProdAttributeName4']")
	public static WebElement Valueattribute4;
	
	@FindBy(xpath="//*[@id='txtProdAttributeCode4']")
	public static WebElement Codeattribute4;
	
	@FindBy(xpath="//*[@id='btnMasterSaveClick']/i")
	public static WebElement saveBtn;
	
	@FindBy(xpath="//*[@id='GenerateNewMasterModel']/div/div[1]/div[2]/div/ul/li[2]/button[6]/i")
	public static WebElement closeBtnforAttributes;
	
	
	public boolean checkItemsAddingforAttribute0inAttributesTab() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(5500);
		
		attribute0Tab.click();
		Thread.sleep(1000);
		checkBox0.click();
		Thread.sleep(1000);
		
		System.err.println("Attribute name " +attribute0Tab.getText());
		
		Valueattribute1.sendKeys("Red");
		Valueattribute1.sendKeys(Keys.TAB);
		Codeattribute1.sendKeys("R");
		Codeattribute1.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		Valueattribute2.sendKeys("Blue");
		Valueattribute2.sendKeys(Keys.TAB);
		Codeattribute2.sendKeys("B");
		Codeattribute2.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		Valueattribute3.sendKeys("Green");
		Valueattribute3.sendKeys(Keys.TAB);
		Codeattribute3.sendKeys("G");
		Codeattribute3.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		Valueattribute4.sendKeys("Yellow");
		Valueattribute4.sendKeys(Keys.TAB);
		Codeattribute4.sendKeys("Y");
		Codeattribute4.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		saveBtn.click();
		
		String expMessage = "Saved Successfully";
		String actMessage = checkValidationMessage(expMessage);
		
		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		} 
		else 
		{
			return false;
		}
	
		
	}
	
	
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[2]/tr[1]/td[4]/input")
	public static WebElement A1_1stRow1stCol;
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[2]/tr[1]/td[5]/input")
	public static WebElement A1_1stRow2ndCol;
	
	
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[2]/tr[2]/td[4]/input")
	public static WebElement A1_2ndRow1stCol;
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[2]/tr[2]/td[5]/input")
	public static WebElement A1_2ndRow2ndCol;
	
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[2]/tr[3]/td[4]/input")
	public static WebElement A1_3rdRow1stCol;
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[2]/tr[3]/td[5]/input")
	public static WebElement A1_3rdRow2ndCol;
	
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[2]/tr[4]/td[4]/input")
	public static WebElement A1_4thRow1stCol;
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[2]/tr[4]/td[5]/input")
	public static WebElement A1_4thRow2ndCol;
	
	
	//Attribute2
	
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[3]/tr[1]/td[4]/input")
	public static WebElement A2_1stRow1stCol;
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[3]/tr[1]/td[5]/input")
	public static WebElement A2_1stRow2ndCol;
	
	
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[3]/tr[2]/td[4]/input")
	public static WebElement A2_2ndRow1stCol;
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[3]/tr[2]/td[5]/input")
	public static WebElement A2_2ndRow2ndCol;
	
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[3]/tr[3]/td[4]/input")
	public static WebElement A2_3rdRow1stCol;
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[3]/tr[3]/td[5]/input")
	public static WebElement A2_3rdRow2ndCol;
	
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[3]/tr[4]/td[4]/input")
	public static WebElement A2_4thRow1stCol;
	@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[3]/tr[4]/td[5]/input")
	public static WebElement A2_4thRow2ndCol;
	
	
	//Attribute3
	
		@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[4]/tr[1]/td[4]/input")
		public static WebElement A3_1stRow1stCol;
		@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[4]/tr[1]/td[5]/input")
		public static WebElement A3_1stRow2ndCol;
		
		
		@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[4]/tr[2]/td[4]/input")
		public static WebElement A3_2ndRow1stCol;
		@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[4]/tr[2]/td[5]/input")
		public static WebElement A3_2ndRow2ndCol;
		
		@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[4]/tr[3]/td[4]/input")
		public static WebElement A3_3rdRow1stCol;
		@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[4]/tr[3]/td[5]/input")
		public static WebElement A3_3rdRow2ndCol;
		
		@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[4]/tr[4]/td[4]/input")
		public static WebElement A3_4thRow1stCol;
		@FindBy(xpath="(//tbody[@id='tblProductAttributesBody'])[4]/tr[4]/td[5]/input")
		public static WebElement A3_4thRow2ndCol;
		
	
	
	
	public boolean checkItemsAddingforAttribute1inAttributesTab() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(4000);
		
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Fruits Attribute Item");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("Fruits Attribute Item");
		 codeTxt.sendKeys(Keys.TAB);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLasttab));
		itemLasttab.click();
		Thread.sleep(1000);*/
		
		attribute1Tab.click();
		Thread.sleep(1000);
		checkBox1.click();
		Thread.sleep(1000);
		
		System.err.println("Attribute name " +attribute1Tab.getText());
		
		A1_1stRow1stCol.sendKeys("Apple");
		A1_1stRow1stCol.sendKeys(Keys.TAB);
		A1_1stRow2ndCol.sendKeys("A");
		A1_1stRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		A1_2ndRow1stCol.sendKeys("Banana");
		A1_2ndRow1stCol.sendKeys(Keys.TAB);
		A1_2ndRow2ndCol.sendKeys("B");
		A1_2ndRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		A1_3rdRow1stCol.sendKeys("Orange");
		A1_3rdRow1stCol.sendKeys(Keys.TAB);
		A1_3rdRow2ndCol.sendKeys("O");
		A1_3rdRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		A1_4thRow1stCol.sendKeys("Grapes");
		A1_4thRow1stCol.sendKeys(Keys.TAB);
		A1_4thRow2ndCol.sendKeys("G");
		A1_4thRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		saveBtn.click();
		
		String expMessage = "Saved Successfully";
		String actMessage = checkValidationMessage(expMessage);
		
		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		} 
		else 
		{
			return false;
		}
	
		
	}
	
	
	
	public boolean checkItemsAddingforAttribute2inAttributesTab() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(4000);
		
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Flowers Attribute Item");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("Flowers Attribute Item");
		 codeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLasttab));
		itemLasttab.click();
		Thread.sleep(1000);
		
		attribute2Tab.click();
		Thread.sleep(1000);
		checkBox2.click();
		Thread.sleep(1000);
		
		A2_1stRow1stCol.sendKeys("Rose");
		A2_1stRow1stCol.sendKeys(Keys.TAB);
		A2_1stRow2ndCol.sendKeys("R");
		A2_1stRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		A2_2ndRow1stCol.sendKeys("Lilly");
		A2_2ndRow1stCol.sendKeys(Keys.TAB);
		A2_2ndRow2ndCol.sendKeys("Li");
		A2_2ndRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		A2_3rdRow1stCol.sendKeys("Hibiscus");
		A2_3rdRow1stCol.sendKeys(Keys.TAB);
		A2_3rdRow2ndCol.sendKeys("H");
		A2_3rdRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		A2_4thRow1stCol.sendKeys("Lotus");
		A2_4thRow1stCol.sendKeys(Keys.TAB);
		A2_4thRow2ndCol.sendKeys("LO");
		A2_4thRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		
		saveBtn.click();
		
		String expMessage = "Saved Successfully";
		String actMessage = checkValidationMessage(expMessage);
		
		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		} 
		else 
		{
			return false;
		}
	
		
	}
	
	
	
	public boolean checkItemsAddingforAttribute3inAttributesTab() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		Thread.sleep(4000);
		
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Electronics Attribute Item");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("Electronics Attribute Item");
		 codeTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLasttab));
		itemLasttab.click();
		Thread.sleep(1000);
		
		attribute3Tab.click();
		Thread.sleep(1000);
		checkBox3.click();
		Thread.sleep(1000);
		
	
		A3_1stRow1stCol.sendKeys("Mobile");
		A3_1stRow1stCol.sendKeys(Keys.TAB);
		A3_1stRow2ndCol.sendKeys("M");
		A3_1stRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		A3_2ndRow1stCol.sendKeys("TV");
		A3_2ndRow1stCol.sendKeys(Keys.TAB);
		A3_2ndRow2ndCol.sendKeys("T");
		A3_2ndRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		A3_3rdRow1stCol.sendKeys("Computer");
		A3_3rdRow1stCol.sendKeys(Keys.TAB);
		A3_3rdRow2ndCol.sendKeys("C");
		A3_3rdRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		A3_4thRow1stCol.sendKeys("Tablet");
		A3_4thRow1stCol.sendKeys(Keys.TAB);
		A3_4thRow2ndCol.sendKeys("Tab");
		A3_4thRow2ndCol.sendKeys(Keys.TAB);
		Thread.sleep(1000);
		
		Thread.sleep(1000);
		
		saveBtn.click();
		
		String expMessage = "Saved Successfully";
		String actMessage = checkValidationMessage(expMessage);
		
		System.out.println("actMessage: "+actMessage);
		System.out.println("expMessage: "+expMessage);
		
		
		
		if (actMessage.equalsIgnoreCase(expMessage)) 
		{
			return true;
		} 
		else 
		{
			return false;
		}
	
		
	}
	
	
	@FindBy(xpath="//*[@id='id_AttributeName0']")
	public static WebElement attribute0Name;
	
	@FindBy(xpath="//*[@id='id_AttributeName1']")
	public static WebElement attribute1Name;
	
	@FindBy(xpath="//*[@id='id_AttributeName2']")
	public static WebElement attribute2Name;
	
	@FindBy(xpath="//*[@id='id_AttributeName3']")
	public static WebElement attribute3Name;
	
	@FindBy(xpath="//*[@id='divProductAttributesModal']//h4")
	public static WebElement itemAttributesTitle;
	
	@FindBy(xpath="//*[@id='prodAttrModalBody']//table//tr/td/label")
	public static List<WebElement> itemAttributeList;
	
	@FindBy(xpath="//*[@id='prodAttrModalBody']//table//tr/td/label/input//..//span")
	public static List<WebElement> itemAttributeListChkBox;
	
	@FindBy(xpath="//*[@id='selectProdAttr']/../span")
	public static WebElement itemAttributes_SelectAllChkBox;
	
	
	@FindBy(xpath="//*[@id='divProductAttributesModal']//button[text()='Ok']")
	public static WebElement itemAttributes_OkBtn;
	
	@FindBy(xpath="//*[@id='divProductAttributesModal']//button[text()='Close']")
	public static WebElement itemAttributes_CloseBtn;
	
	
	public boolean checkSavedAttributesListofAttribute0inAttributesTab() throws InterruptedException
	{
		Thread.sleep(4000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLasttab));
		itemLasttab.click();
		Thread.sleep(1000);*/
		
		attribute0Tab.click();
		Thread.sleep(1000);
		checkBox0.click();
		Thread.sleep(1000);
		
		attribute0Name.click();
		String actAttributeName=attribute0Name.getAttribute("value");
		String expAttributeName="Colors";
		
		System.out.println("Act Attribute0 Name  "  +actAttributeName);
		System.out.println("Exp Attribute0 Name  "  +expAttributeName);
			
		attribute0Name.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		String actItemAttributeTitle= itemAttributesTitle.getText();
		String expItemAttributeTitle	=	"Item Attributes";
		
		System.out.println("actItemAttributeTitle"	 +actItemAttributeTitle);
		System.out.println("expItemAttributeTitle"	 +expItemAttributeTitle);
		
		ArrayList<String>itemAttributesList=new ArrayList<String>();
		
		
		for(int i=0;i<itemAttributeList.size();i++)
		{
			itemAttributesList.add(itemAttributeList.get(i).getText());
			
		}
		
		String actAttribute0ItemList=itemAttributesList.toString();
		String expAttribute0ItemList="[Blue, Green, Red, Yellow]";
		
		System.out.println("Actual 	 Item List in Attribute0  "  +actAttribute0ItemList);
		System.out.println("Expected Item List in Attribute0  "  +expAttribute0ItemList);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemAttributes_SelectAllChkBox));
		itemAttributes_SelectAllChkBox.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemAttributes_OkBtn));
		itemAttributes_OkBtn.click();
		Thread.sleep(8000);
		
		
		if(actAttributeName.equalsIgnoreCase(expAttributeName) && actItemAttributeTitle.equalsIgnoreCase(expItemAttributeTitle)
				&&  actAttribute0ItemList.equalsIgnoreCase(expAttribute0ItemList))
		{
			return true;
		}
		else
		{
			return false;
			
		}
		
		
		
	}
	
	
	public boolean checkSavedAttributesListofAttribute1inAttributesTab() throws InterruptedException
	{
		Thread.sleep(3000);
		
		ClickUsingJs(attribute1Tab);
		Thread.sleep(2000);
		checkBox1.click();
		Thread.sleep(2000);
		
		attribute1Name.click();
		String actAttributeName=attribute1Name.getAttribute("value");
		String expAttributeName="Fruits";
		
		System.out.println("Attribute0 Name	 "  +actAttributeName);
		System.out.println("Attribute0 Name	 "  +actAttributeName);
			
		attribute1Name.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		String actItemAttributeTitle=itemAttributesTitle.getText();
		String expItemAttributeTitle="Item Attributes";
		
		System.out.println("actItemAttributeTitle : "  +actItemAttributeTitle);
		System.out.println("expItemAttributeTitle : "  +expItemAttributeTitle);
		
		ArrayList<String>itemAttributesList=new ArrayList<String>();
		
		
		for(int i=0;i<itemAttributeList.size();i++)
		{
			itemAttributesList.add(itemAttributeList.get(i).getText());
			
		}
		
		String actAttribute0ItemList=itemAttributesList.toString();
		String expAttribute0ItemList="[Apple, Banana, Grapes, Orange]";
		
		System.out.println("Actual 	Item List in Attribute0	  "	 +	actAttribute0ItemList);
		System.out.println("Expected Item List in Attribute0  "	 +	expAttribute0ItemList);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemAttributes_SelectAllChkBox));
		itemAttributes_SelectAllChkBox.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemAttributes_OkBtn));
		itemAttributes_OkBtn.click();
		Thread.sleep(2000);
		
		
		
		if(actAttributeName.equalsIgnoreCase(expAttributeName) && actItemAttributeTitle.equalsIgnoreCase(expItemAttributeTitle)
				&&  actAttribute0ItemList.equalsIgnoreCase(expAttribute0ItemList))
		{
			return true;
		}
		else
		{
			return false;
			
		}
		
		
		
	}
	
	
	public boolean checkSavedAttributesListofAttribute2inAttributesTab() throws InterruptedException
	{
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLasttab));
		itemLasttab.click();
		Thread.sleep(1000);
		
		attribute2Tab.click();
		Thread.sleep(1000);
		checkBox2.click();
		Thread.sleep(1000);
		
		attribute2Name.click();
		String actAttributeName=attribute2Name.getAttribute("value");
		String expAttributeName="Flowers";
		
		System.out.println("Act Attribute0 Name	"	+actAttributeName);
		System.out.println("Exp Attribute0 Name	"	+expAttributeName);
			
		attribute2Name.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		String actItemAttributeTitle= itemAttributesTitle.getText();
		String expItemAttributeTitle	=	"Item Attributes";
		
		System.out.println("actItemAttributeTitle  "	+actItemAttributeTitle);
		System.out.println("expItemAttributeTitle  "	+expItemAttributeTitle);
		
		ArrayList<String>itemAttributesList=new ArrayList<String>();
		
		
		for(int i=0;i<itemAttributeList.size();i++)
		{
			itemAttributesList.add(itemAttributeList.get(i).getText());
			
		}
		
		String actAttribute0ItemList=itemAttributesList.toString();
		String expAttribute0ItemList="[Hibiscus, Lilly, Lotus, Rose]";
		
		System.out.println("Actual 	Item List in Attribute0			"		+		actAttribute0ItemList);
		System.out.println("Expected Item List in Attribute0		"		+		expAttribute0ItemList);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemAttributes_SelectAllChkBox));
		itemAttributes_SelectAllChkBox.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemAttributes_OkBtn));
		itemAttributes_OkBtn.click();
		Thread.sleep(2000);
		
		
		
		if(actAttributeName.equalsIgnoreCase(expAttributeName) && actItemAttributeTitle.equalsIgnoreCase(expItemAttributeTitle)
				&&  actAttribute0ItemList.equalsIgnoreCase(expAttribute0ItemList))
		{
			return true;
		}
		else
		{
			return false;
			
		}
		
		
		
	}
	
	
	public boolean checkSavedAttributesListofAttribute3inAttributesTab() throws InterruptedException
	{
		Thread.sleep(8000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemLasttab));
		itemLasttab.click();
		Thread.sleep(1000);
		
		attribute3Tab.click();
		Thread.sleep(1000);
		checkBox3.click();
		Thread.sleep(1000);
		
		attribute3Name.click();
		String actAttributeName=attribute3Name.getAttribute("value");
		String expAttributeName="Electronics";
		
		System.out.println("Act Attribute0 Name	 " +actAttributeName);
		System.out.println("Exp Attribute0 Name	 " +expAttributeName);
		
			
		attribute3Name.sendKeys(Keys.TAB);
		
		Thread.sleep(2000);
		
		String actItemAttributeTitle= itemAttributesTitle.getText();
		String expItemAttributeTitle	=	"Item Attributes";
		
		System.out.println("actItemAttributeTitle "	 +actItemAttributeTitle);
		System.out.println("expItemAttributeTitle "	 +expItemAttributeTitle);
		
		
		ArrayList<String>itemAttributesList=new ArrayList<String>();
		
		
		for(int i=0;i<itemAttributeList.size();i++)
		{
			itemAttributesList.add(itemAttributeList.get(i).getText());
			
		}
		
		String actAttribute0ItemList=itemAttributesList.toString();
		String expAttribute0ItemList="[Computer, Mobile, TV, Tablet]";
		
		System.out.println("Actual 	Item List in Attribute0	   "  +	actAttribute0ItemList);
		System.out.println("Expected Item List in Attribute0   "  +	expAttribute0ItemList);
		
		for(int i=0;i<itemAttributeList.size();i++)
		{
			itemAttributeListChkBox.get(i).click();
			
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemAttributes_OkBtn));
		itemAttributes_OkBtn.click();
		Thread.sleep(2000);
		
		
		if(actAttributeName.equalsIgnoreCase(expAttributeName) && actItemAttributeTitle.equalsIgnoreCase(expItemAttributeTitle)
				&&  actAttribute0ItemList.equalsIgnoreCase(expAttribute0ItemList))
		{
			return true;
		}
		else
		{
			return false;
			
		}
		
		
	}
	
	
	@FindBy(xpath="//*[@id='divBtnGroup0']")
	public static WebElement generalTab1;
	
	public boolean checkSavingItemwithSelectedAttributesList() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalTab1));
		generalTab1.click();
		Thread.sleep(2000);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
		 nameTxt.click();
		 nameTxt.clear();
		 nameTxt.sendKeys("Attribute Item");
		 nameTxt.sendKeys(Keys.TAB);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(codeTxt));
		 codeTxt.click();
		 codeTxt.clear();
		 codeTxt.sendKeys("Attribute Item");
		 codeTxt.sendKeys(Keys.TAB);
		 
		 	saveBtn.click();
		 	Thread.sleep(2000);
			
			String expMessage = "Saved Successfully";
			String actMessage = checkValidationMessage(expMessage);
			
			System.out.println("actMessage: "+actMessage);
			System.out.println("expMessage: "+expMessage);
			
			Thread.sleep(4000);
			
			closeBtnforAttributes.click(); 
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
	
	
	@FindBy(xpath="//*[@id='LandingGridBody']/tr/td[12]")
	public static List<WebElement> masterGridBodyName;
	
	public boolean checkSavedAttributeItemwithAllAttributeItemList() throws InterruptedException
	{
		 int count = masterGridBodyName.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = masterGridBodyName.get(i).getText();
			Thread.sleep(2000);
			
			if (data.equalsIgnoreCase("Attribute Item")) 
			{
				getAction().doubleClick(masterGridBodyChkbox.get(i)).build().perform();
				Thread.sleep(1000);
				break;
			}
		}
		
		
		Thread.sleep(8000);
		
		ArrayList<String>attributeList=new ArrayList<String>();
		
		 int count1 = masterGridBodyName.size();
			
			for (int i = 0; i < count1; i++) 
			{
				String data = masterGridBodyName.get(i).getText();
				Thread.sleep(2000);
				
				attributeList.add(data);
		
			}
			
			String actAttributeItemList=attributeList.toString();
			String expAttributeItemList="[Attribute Item-Blue/Apple/Hibiscus/Computer, Attribute Item-Blue/Apple/Hibiscus/Mobile, Attribute Item-Blue/Apple/Hibiscus/TV, Attribute Item-Blue/Apple/Hibiscus/Tablet, Attribute Item-Blue/Apple/Lilly/Computer, Attribute Item-Blue/Apple/Lilly/Mobile, Attribute Item-Blue/Apple/Lilly/TV, Attribute Item-Blue/Apple/Lilly/Tablet, Attribute Item-Blue/Apple/Lotus/Computer, Attribute Item-Blue/Apple/Lotus/Mobile, Attribute Item-Blue/Apple/Lotus/TV, Attribute Item-Blue/Apple/Lotus/Tablet, Attribute Item-Blue/Apple/Rose/Computer, Attribute Item-Blue/Apple/Rose/Mobile, Attribute Item-Blue/Apple/Rose/TV, Attribute Item-Blue/Apple/Rose/Tablet, Attribute Item-Blue/Banana/Hibiscus/Computer, Attribute Item-Blue/Banana/Hibiscus/Mobile, Attribute Item-Blue/Banana/Hibiscus/TV, Attribute Item-Blue/Banana/Hibiscus/Tablet, Attribute Item-Blue/Banana/Lilly/Computer, Attribute Item-Blue/Banana/Lilly/Mobile, Attribute Item-Blue/Banana/Lilly/TV, Attribute Item-Blue/Banana/Lilly/Tablet, Attribute Item-Blue/Banana/Lotus/Computer, Attribute Item-Blue/Banana/Lotus/Mobile, Attribute Item-Blue/Banana/Lotus/TV, Attribute Item-Blue/Banana/Lotus/Tablet, Attribute Item-Blue/Banana/Rose/Computer, Attribute Item-Blue/Banana/Rose/Mobile, Attribute Item-Blue/Banana/Rose/TV, Attribute Item-Blue/Banana/Rose/Tablet, Attribute Item-Blue/Grapes/Hibiscus/Computer, Attribute Item-Blue/Grapes/Hibiscus/Mobile, Attribute Item-Blue/Grapes/Hibiscus/TV, Attribute Item-Blue/Grapes/Hibiscus/Tablet, Attribute Item-Blue/Grapes/Lilly/Computer, Attribute Item-Blue/Grapes/Lilly/Mobile, Attribute Item-Blue/Grapes/Lilly/TV, Attribute Item-Blue/Grapes/Lilly/Tablet, Attribute Item-Blue/Grapes/Lotus/Computer, Attribute Item-Blue/Grapes/Lotus/Mobile, Attribute Item-Blue/Grapes/Lotus/TV, Attribute Item-Blue/Grapes/Lotus/Tablet, Attribute Item-Blue/Grapes/Rose/Computer, Attribute Item-Blue/Grapes/Rose/Mobile, Attribute Item-Blue/Grapes/Rose/TV, Attribute Item-Blue/Grapes/Rose/Tablet, Attribute Item-Blue/Orange/Hibiscus/Computer, Attribute Item-Blue/Orange/Hibiscus/Mobile]";
			
			System.out.println("Actual Attribute Item Sub List  	"	+actAttributeItemList);
			System.out.println("Expected Attribute Item Sub List	"	+expAttributeItemList);
			
			if(actAttributeItemList.equalsIgnoreCase(expAttributeItemList))
			{
				return true;
			}
			else
			{
				return false;
			}
			
	}
	
	@FindBy(xpath="//*[@id='ol_treeNavigation']/li[1]/span")
	public static WebElement ItemLbl;
	
	@FindBy(xpath="//*[@id='ddlDatatype_ExtraField']")
	public static WebElement item_DataTypeSelect;
	
	@FindBy(xpath="//*[@id='ExtraField_ControlType']")
	public static WebElement item_ControlTypeSelect;
	
	
	@FindBy(xpath="//span[@id='toggle_ribbon']")
	public static WebElement toggleRibbon;
	
	
	/*@FindBy(xpath="//i[@class='icon-font6 icon-custamize']")
	public static WebElement masterCustamizemasterBtn; */
	
	
	@FindBy(xpath="//*[@id='btnAdd_Field_MasterCust']//i")
	private static WebElement accountAddBtn;
	
	public boolean checkControlTypeofDataTypeFieldinItemMasterCustomization() throws InterruptedException
	{
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
		homeMenu.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mastersMenu));
		mastersMenu.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemMenu));
		itemMenu.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(item));
		item.click();
		
		Thread.sleep(4000);
	   	
	   	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterRibbonToExpandOptions));	
	   	masterRibbonToExpandOptions.click();
	   	Thread.sleep(2000);
	   		
		ClickUsingJs(masterCustamizemasterBtn);
		
		Thread.sleep(5500);
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(generalHeaderDetailsTab));
    	 generalHeaderDetailsTab.click();
    	 
    	 Thread.sleep(4000);
    	 
    	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(accountAddBtn));
    	 accountAddBtn.click();
    	 
    	 Thread.sleep(6000);
	   	
		Select s=new Select(item_DataTypeSelect);
		int count=s.getOptions().size();
		
		Select s1=new Select(item_ControlTypeSelect);
		
		ArrayList<String>item_CustDataTypeList=new ArrayList<String>();
		
		for(int i=0;i<count;i++)
		{
			String data= s.getOptions().get(i).getText();
			
			s.getOptions().get(i).click();
			Thread.sleep(4000);
			
			item_CustDataTypeList.add(data);
			
			item_CustDataTypeList.add("-->");
			String data1= s1.getFirstSelectedOption().getText();
			
			Thread.sleep(2000);
			item_CustDataTypeList.add(data1);
			
			System.out.println();
			
		}
		
		String actControlTypeList=item_CustDataTypeList.toString().replaceAll(",", "");
		String expControlTypeList="[Text --> TextBox Number --> TextBox Boolean --> Checkbox DateTime --> TextBox Date --> Date Picker Time --> Time Picker Fraction --> TextBox Picture --> Picture Box String List --> Dropdown Number List --> Dropdown Document Viewer --> Document Control Master --> Option Control Big Number --> TextBox External Table --> Option Control Small Number --> TextBox Tiny Number --> TextBox Gregorian --> Date Picker Hijri --> Date Picker Shamsi --> Date Picker LOV --> Option Control]";
		System.out.println("Actual   Control Type 	"	+	actControlTypeList);
		System.out.println("Expected Control Type 	"	+	expControlTypeList);
		
		if(actControlTypeList.equalsIgnoreCase(expControlTypeList))
		{
			return true;
		}
		else
		{
			return false;
		}
		
	}
	
	
	
	public boolean checkLogoutAttributesPage() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
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

		public AttributesPage(WebDriver driver)
		{
			PageFactory.initElements(driver,this);
		}
		
}
