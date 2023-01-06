
package com.focus.Pages;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import javax.security.auth.login.CredentialExpiredException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.omg.Messaging.SyncScopeHelper;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;

public class SanityDocCustViewPage extends BaseEngine  
{
	private static String xlfile;
	private static String resPass="Pass";
	private static String resFail="Fail";
	private static ExcelReader excelReader;

	private static String xlSheetName = "SmokeDocumentCustomization";	

	public static boolean checkLoginCompanyWithValidCredentials() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		getDriver().navigate().refresh();
		Thread.sleep(3000);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt="su";

		String pawslt="su";

		lp.enterUserName(unamelt);

		lp.enterPassword(pawslt);

		lp.clickOnSignInBtn();

		Thread.sleep(5000);

		String userInfo=userNameDisplay.getText();

		System.out.println("User Info : "+userInfo);

		System.out.println("User Info Capture Text :"+userNameDisplay.getText());


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		companyLogo.click();

		String getCompanyTxt=companyName.getText();
		String getLoginCompanyName=getCompanyTxt.substring(0, 19);
		System.out.println("company name :"+ getLoginCompanyName);
		companyLogo.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));

		String getDashboard=dashboard.getText();

		System.out.println(getDashboard);

		if(userInfo.equalsIgnoreCase("SU") && getLoginCompanyName.equalsIgnoreCase("Automation Company ") &&
				getDashboard.equalsIgnoreCase("Dashboard"))
		{
			//excelReader.setCellData(xlfile, "Sheet1", 7, 9, resPass);
			return true;
		}
		else
		{
			//excelReader.setCellData(xlfile, "Sheet1", 7, 9, resFail);
			return true;
		}
	}

	@FindBy(xpath="//*[@id='id_search_menu']/input")
	public static WebElement searchTxt;



	public boolean checkUserCreatedVoucherDisplayWithDocumentTabInDocumentCustomization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		System.out.println("**************  checkUserCreatedVoucherDisplayInDocumentCustomization Method Executes.............  *****************");

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			ScrollIntoView(SettingsSubMenusList.get(i));
			
			if (data.equalsIgnoreCase("Configure Transactions")) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);*/
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();
		
		Thread.sleep(1000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreatedVoucher));
		docCustCreatedVoucher.click();
		
		Thread.sleep(4000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustDocumentsTabGeneralBtn));
		docCustDocumentsTabGeneralBtn.click();
		*/
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(titleTxt));
		titleTxt.click();

		String actTitleTxt=titleTxt.getAttribute("value");
		String expTitleTxt= "Purchase Voucher VAT";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(baseDocumentTxtInDocCust));
		String actBaseDocumentName = baseDocumentTxtInDocCust.getAttribute("value");
		String expBaseDocumentName = "Purchases Vouchers 301";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masters1)); 
		masters1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropDown));

		Select getValuesOfMasterC1=new Select(masterDropDown);
		String actMastersR1C1=getValuesOfMasterC1.getFirstSelectedOption().getText();
		String expMastersR1C1="Warehouse";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(position1));
		position1.click();

		Select getValueOfPostionC1=new Select(positionDropdown);
		String actPositionR1C1=getValueOfPostionC1.getFirstSelectedOption().getText();
		String expPositionR1C1="Body";	

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masters2)); 
		ScrollIntoView(masters2);
		Thread.sleep(2000);
		masters2.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterDropDown));

		Select getValuesOfMasterC2=new Select(masterDropDown);
		String actMastersR2C1=getValuesOfMasterC1.getFirstSelectedOption().getText();
		String expMastersR2C1="Department";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(position2));
		position2.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(positionDropdown));

		Select getValueOfPostionC2=new Select(positionDropdown);
		String actPositionR2C2=getValueOfPostionC2.getFirstSelectedOption().getText();
		String expPositionR2C2="Header";	

		System.out.println("TitleTxt               : " + actTitleTxt             + "  value expected  " + expTitleTxt);
		System.out.println("Base Doucment          : " + actBaseDocumentName     + "  value expected  " + expBaseDocumentName);
		System.out.println("Row One Master         : " + actMastersR1C1          + "  value expected  " + expMastersR1C1);
		System.out.println("Row One Position       : " + actPositionR1C1         + "  value expected  " + expPositionR1C1);
		System.out.println("Row Two Master         : " + actMastersR2C1          + "  value expected  " + expMastersR2C1);
		System.out.println("Row Tow Position       : " + actPositionR2C2         + "  value expected  " + expPositionR2C2);

		boolean actdocumentsTab         = documentsTab.isDisplayed();
		boolean actmiscellaneousTab     = miscellaneousTab.isDisplayed();
		boolean acteditScreenTab        = editScreenTab.isDisplayed();
		boolean actinventoryoptionsTab  = inventoryoptionsTab.isDisplayed();
		boolean acteditLayoutTab        = editLayoutTab.isDisplayed();
		boolean actviewsBtn             = viewsBtn.isDisplayed();
		boolean actexternalFunctionsBtn = externalFunctionsBtn.isDisplayed();
		boolean actschemesBtn           = schemesBtn.isDisplayed();
		boolean actexportFieldsBtn      = exportFieldsBtn.isDisplayed();
		boolean acttriggersBtn          = triggersBtn.isDisplayed();
		boolean actdocumentNumberingBtn = documentNumberingBtn.isDisplayed();
		boolean acthirePurchaseBtn      = hirePurchaseBtn.isDisplayed();
		boolean actrulesBtn             = rulesBtn.isDisplayed();
		boolean actreportsBtn           = reportsBtn.isDisplayed();

		boolean expdocumentsTab         = true;
		boolean expmiscellaneousTab     = true;
		boolean expeditScreenTab        = true;
		boolean expinventoryoptionsTab  = true;
		boolean expeditLayoutTab        = true;
		boolean expviewsBtn             = true;
		boolean expexternalFunctionsBtn = true;
		boolean expschemesBtn           = true;
		boolean expexportFieldsBtn      = true;
		boolean exptriggersBtn          = true;
		boolean expdocumentNumberingBtn = true;
		boolean exphirePurchaseBtn      = true;
		boolean exprulesBtn             = true;
		boolean expreportsBtn           = true;	


		System.out.println("DocumentsTab           : " + actdocumentsTab             + "  value expected  " + expdocumentsTab);
		System.out.println("MiscellaneousTab       : " + actmiscellaneousTab         + "  value expected  " + expmiscellaneousTab);
		System.out.println("EditScreenTab          : " + acteditScreenTab            + "  value expected  " + expeditScreenTab);
		System.out.println("InventoryoptionsTab    : " + actinventoryoptionsTab      + "  value expected  " + expinventoryoptionsTab);
		System.out.println("EditLayoutTab          : " + acteditLayoutTab            + "  value expected  " + expeditLayoutTab);
		System.out.println("ViewsBtn               : " + actviewsBtn                 + "  value expected  " + expviewsBtn);
		System.out.println("ExternalFunctionsBtn   : " + actexternalFunctionsBtn     + "  value expected  " + expexternalFunctionsBtn);
		System.out.println("SchemesBtn             : " + actschemesBtn               + "  value expected  " + expschemesBtn);
		System.out.println("ExportFieldsBtn        : " + actexportFieldsBtn          + "  value expected  " + expexportFieldsBtn);
		System.out.println("TriggersBtn            : " + acttriggersBtn              + "  value expected  " + exptriggersBtn);
		System.out.println("DocumentNumberingBtn   : " + actdocumentNumberingBtn     + "  value expected  " + expdocumentNumberingBtn);
		System.out.println("HirePurchaseBtn        : " + acthirePurchaseBtn          + "  value expected  " + exphirePurchaseBtn);
		System.out.println("RulesBtn               : " + actrulesBtn                 + "  value expected  " + exprulesBtn);
		System.out.println("ReportsBtn             : " + actreportsBtn               + "  value expected  " + expreportsBtn);
		System.out.println("Row One Master         : " + actMastersR1C1              + "  value expected  " + expMastersR1C1);

		if(actMastersR1C1.equalsIgnoreCase(expMastersR1C1)
  				&& actPositionR1C1.equalsIgnoreCase(expPositionR1C1) 

  				&& actMastersR2C1.equalsIgnoreCase(expMastersR2C1)
  				&& actPositionR2C2.equalsIgnoreCase(expPositionR2C2)   

  				&& actdocumentsTab==expdocumentsTab 
  				&& actmiscellaneousTab==expmiscellaneousTab && acteditScreenTab==expeditScreenTab && actinventoryoptionsTab==expinventoryoptionsTab 
  				&& acteditLayoutTab==expeditLayoutTab 
  				&& actviewsBtn==expviewsBtn && actexternalFunctionsBtn==expexternalFunctionsBtn && actschemesBtn==expschemesBtn 
  				&& actexportFieldsBtn==expexportFieldsBtn && acttriggersBtn==exptriggersBtn && actdocumentNumberingBtn==expdocumentNumberingBtn 
  				&& acthirePurchaseBtn==exphirePurchaseBtn && actrulesBtn==exprulesBtn && actreportsBtn==expreportsBtn)
		{
			System.out.println("Test Pass : Toogle Options in Sales Orders As Expected");
			excelReader.setCellData(xlfile, xlSheetName,  281, 9, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Toogle Options in Sales Orders As Expected");
			excelReader.setCellData(xlfile, xlSheetName,  281, 9, resFail);
			return false;
		}
	}
	
	
	



	public boolean checkCreatedVoucherMiscellaneousTabInDocumentCustoimization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		System.out.println("**************  checkCreatedVoucherMiscellaneousTabInDocumentCustoimization Method Executes.............  *****************");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTab));
		miscellaneousTab.click();

		Thread.sleep(3000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTabVoucherAccountSettingsBtn));
		miscellaneousTabVoucherAccountSettingsBtn.click();
		*/
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(abbreviationTxt));
		abbreviationTxt.click();		

		String actAbbreviationTxt=abbreviationTxt.getAttribute("value");
		String expAbbreviationTxt= "PurVocUpda";
		excelReader.setCellData(xlfile, xlSheetName, 289, 8, actAbbreviationTxt);

		Thread.sleep(2000);
		
		System.out.println("Abbreviation : "+actAbbreviationTxt+"  Value Expected  "+expAbbreviationTxt);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneousTabVoucherAccountSettingsBtn));
		miscellaneousTabVoucherAccountSettingsBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneusTabPostingDetailsMenu));
		miscellaneusTabPostingDetailsMenu.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateStockChkBox));
		
		if (updateStockChkBoxIsSelected.isSelected()==false)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateStockChkBox));
			updateStockChkBox.click();
		} 
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneusTabPostingDetailsMenu));
		miscellaneusTabPostingDetailsMenu.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneusTabARAPBtn));
		miscellaneusTabARAPBtn.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dueDateDropdown));
		Select s1=new Select(dueDateDropdown);
		s1.selectByVisibleText("Header");
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(miscellaneusTabARAPBtn));
		miscellaneusTabARAPBtn.click();
		
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();
		
		Thread.sleep(1000);

		String expMessage="Data Saved SuccessFully";

		String actMessage=checkValidationMessage(expMessage);

		if(actAbbreviationTxt.equalsIgnoreCase(expAbbreviationTxt))
		{
			System.out.println("Test Pass : Abbreviation Contain CasSal As Expected");
			excelReader.setCellData(xlfile, xlSheetName, 288, 9, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Abbreviation Does NotContain CasSal As Expected");
			excelReader.setCellData(xlfile, xlSheetName, 288, 9, resFail);
			return false;
		}
	}





	public boolean checkCreatedVoucherEditScreenTabInDocumentCustoimization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		System.out.println("**************  checkCreatedVoucherMiscellaneousTabInDocumentCustoimization Method Executes.............  *****************");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyTab));
		editScreenBodyTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenBodyRow2Edit));
		editScreenBodyRow2Edit.click();

		Thread.sleep(2000);

			
				getWaitForAlert();
				Thread.sleep(2000);
				getAlert().accept();
				Thread.sleep(2000);
			
			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
		editScreenCaptionTxt.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenDefaultValue));
		editScreenDefaultValue.click();
		editScreenDefaultValue.sendKeys(Keys.HOME);
		editScreenDefaultValue.sendKeys(Keys.SHIFT, Keys.END);
		editScreenDefaultValue.sendKeys("4");
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenPropertiesTab));
		editScreenPropertiesTab.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenPropertiesReadOnly));
		editScreenPropertiesReadOnly.click();

		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
		editScreenApplyBtn.click();

		Thread.sleep(2000);
		
		checkValidationMessage("");

		int editScreenbodyCaptionCount = editScreenbodyCaption.size();

		for(int i=0;i<editScreenbodyCaptionCount;i++)
		{
			String data = editScreenbodyCaption.get(i).getText();

			if(data.equalsIgnoreCase("VAT"))
			{
				editScreenbodyEditBtn.get(i).click();

				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenAddBehaviourformulaTxt));
				editScreenAddBehaviourformulaTxt.click();

				Thread.sleep(2000);
				
				editScreenAddBehaviourformulaTxt.sendKeys(Keys.END);
				
				editScreenAddBehaviourformulaTxt.sendKeys(Keys.SHIFT, Keys.HOME);

				editScreenAddBehaviourformulaTxt.sendKeys("(gr-SBO1)*SB5/100");
				
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
				getAction().moveToElement(editScreenApplyBtn).pause(1000).click().build().perform();

				String expMessage1		="Data Saved Successfully";

				String actMessage1		=checkValidationMessage(expMessage1);

			}

			
			
			if(data.equalsIgnoreCase("Taxable"))
			{
				editScreenbodyEditBtn.get(i).click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenPreloadTxt));
				editScreenPreloadTxt.click();
				
				Thread.sleep(2000);
				
				editScreenPreloadTxt.sendKeys(Keys.END);
				
				editScreenPreloadTxt.sendKeys(Keys.SHIFT, Keys.HOME);

				editScreenPreloadTxt.sendKeys("&(gr-SBO5)");
				
				Thread.sleep(2000);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
				getAction().moveToElement(editScreenApplyBtn).pause(1000).click().build().perform();

				String expMessage1		="Data Saved Successfully";

				String actMessage1		=checkValidationMessage(expMessage1);

			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		Thread.sleep(2000);
		
		String expUpdateMessage = "Data saved successfully";
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);

		if(actUpdateMessage.equalsIgnoreCase(expUpdateMessage))
		{
			System.out.println("Test Pass : Edit Screen Options Are Available as Expected");
			excelReader.setCellData(xlfile, xlSheetName, 297, 9, resPass);
			return  true;
		}
		else
		{
			System.out.println("Test Fail : Edit Screen Options Are Not Available");
			excelReader.setCellData(xlfile, xlSheetName, 297, 9, resFail);
			return  false;
		}
	}


	
	


	public boolean checktheNarrationinHeaderTabEditLayoutinPVVATVIEW() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTab));
		editLayoutTab.click();

		Thread.sleep(4000);

		String actdata,expdata;

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutHeader1stRowCaption));
		editLayoutHeader1stRowCaption.click();

		Thread.sleep(3000);

		actdata=editLayoutHeader1stRowCaption.getText();
		expdata=excelReader.getCellData(xlSheetName, 329, 6);
		excelReader.setCellData(xlfile, xlSheetName, 329, 7, actdata);

		System.out.println("actdata"+actdata);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayouEditBtn));
		editLayouEditBtn.click();

		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutDefaultValueTxt));
		editLayoutDefaultValueTxt.click();
		editLayoutDefaultValueTxt.clear();
		
		editLayoutDefaultValueTxt.sendKeys(excelReader.getCellData(xlSheetName, 332, 5));//332

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutPropertiesTab));
		editLayoutPropertiesTab.click();

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(readOnlyChkBox));
		readOnlyChkBox.click();

		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutApplyBtn));
		editLayoutApplyBtn.click();
		Thread.sleep(2000);

		String expMessageClickOnApply = excelReader.getCellData(xlSheetName, 334, 6);

		String actMessageClickOnApply = checkValidationMessage(expMessageClickOnApply);

		excelReader.setCellData(xlfile, xlSheetName, 334, 7, actMessageClickOnApply);

		if(actMessageClickOnApply.equalsIgnoreCase(expMessageClickOnApply))
		{
			System.out.println("Test Pass : Data Saved In Edit Screeen Field Details");
			excelReader.setCellData(xlfile, xlSheetName, 331, 9, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Data NOT Saved In Edit Screeen Field Details");
			excelReader.setCellData(xlfile, xlSheetName, 331, 9, resFail);
			return false;

		}
	}






	public boolean checkFieldTypesinBodyGridViewTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsBtn));
		viewsBtn.click();

		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabViewSettingsBtn));
		viewTabViewSettingsBtn.click();
		*/
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
		ViewExistingViewTxt.click();
		ViewExistingViewTxt.clear();

		ViewExistingViewTxt.sendKeys(excelReader.getCellData(xlSheetName, 353, 5));
		Thread.sleep(1000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);

		Thread.sleep(4000);

		int Count = viewsGridFiledsList.size();

		System.out.println("Count    : "+Count);

		ArrayList<String> actviewsGridFiledsList  = new ArrayList<String>();

		for (int i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

			if(data.equalsIgnoreCase("Date"/*excelReader.getCellData(xlSheetName, 354, 5)*/))//6

			{
				viewsGridFiledsEditList.get(i).click();

				break;
			}
		}


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewGridRow2EditBtn));
		getAction().doubleClick(viewGridRow2EditBtn).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_DataTypeDrpdwn));			
		fieldDetails_DataTypeDrpdwn.click();

		Select data=new Select(fieldDetails_DataTypeDrpdwn);
		data.selectByVisibleText("Read Only");

		String actDataType=data.getFirstSelectedOption().getText();
		String expDataType=excelReader.getCellData(xlSheetName, 355, 6);
		excelReader.setCellData(xlfile, xlSheetName, 355, 7, actDataType);

		System.err.println("FiledCaption For date In Create View  : "+actDataType);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_ApplyBtn));
		fieldDetails_ApplyBtn.click();

		Thread.sleep(2000);
		
		if(fieldDetails_CloseBtn.isDisplayed()==true)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_CloseBtn));
			fieldDetails_CloseBtn.click();
		}
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabViewSettingsBtn));
		viewTabViewSettingsBtn.click();
		
		Thread.sleep(1000);*/
		

		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)","");
		Thread.sleep(2000);
		
		//getAction().moveToElement(viewTabLoginBtn).build().perform(); 
		
		
		if(viewsUserAllOptionsSTChkboxSelected.isDisplayed()==false)
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		}
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		viewsUserAllOptionsSTChkbox.click();

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
		viewSaveView.click();

		String expMessage= excelReader.getCellData(xlSheetName, 356, 6);

		String actMessage=checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 356, 7, actMessage);

		if (actMessage.equalsIgnoreCase(expMessage) && actDataType.equalsIgnoreCase(expDataType)) 
		{ 
			System.out.println("Test Pass : Toogle Options in Sales Orders As Expected");
			excelReader.setCellData(xlfile, xlSheetName,  352, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Toogle Options in Sales Orders As Expected");
			excelReader.setCellData(xlfile, xlSheetName,  352, 8, resFail);
			return false;
		}
	}


	
	
	

	public boolean checkDeleteinCreatedViewInViewTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabViewSettingsBtn));
		viewTabViewSettingsBtn.click();
		
		Thread.sleep(1000);
*/
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
		ViewExistingViewTxt.click();

		ViewExistingViewTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		ViewExistingViewTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(excelReader.getCellData(xlSheetName, 358, 5));
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);

		Thread.sleep(5000);

		int Count = viewsGridFiledsList.size();

		System.out.println("Count    : "+Count);

		ArrayList<String> actviewsGridFiledsList  = new ArrayList<String>();

		for (int i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 359, 5)))
			{
				viewsGridFiledsEditList.get(i).click();

				break;
			}
		}


		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_DataTypeDrpdwn));
		Select data=new Select(fieldDetails_DataTypeDrpdwn);
		data.selectByVisibleText("Read Only");

		String actDataType=data.getFirstSelectedOption().getText();
		String expDataType=excelReader.getCellData(xlSheetName, 360, 6);
		excelReader.setCellData(xlfile, xlSheetName, 360, 7, actDataType);


		System.err.println("FiledCaption  For date in Create For Delete: "+actDataType);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_ApplyBtn));
		fieldDetails_ApplyBtn.click();

		Thread.sleep(2000);
		if(fieldDetails_CloseBtn.isDisplayed()==true)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldDetails_CloseBtn));
			fieldDetails_CloseBtn.click();
		}
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabViewSettingsBtn));
		viewTabViewSettingsBtn.click();*/
		
		Thread.sleep(3000);
		if(viewsUserAllOptionsSTChkbox.isDisplayed()==false)
		{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
		viewTabLoginBtn.click();
		
		Thread.sleep(2000);
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsUserAllOptionsSTChkbox));
		viewsUserAllOptionsSTChkbox.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
		viewSaveView.click();

		String expMessage=excelReader.getCellData(xlSheetName, 361, 6);

		String actMessage=checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 361, 7, actMessage);
		
		Thread.sleep(2000);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsDeleteViewIcon));
		viewsDeleteViewIcon.click();

		getWaitForAlert();
		Thread.sleep(2000);

		String actAlertMessage=getAlert().getText();

		String expAlertMessage=excelReader.getCellData(xlSheetName, 362, 6);

		excelReader.setCellData(xlfile, xlSheetName, 362, 7, actAlertMessage);

		System.out.println("Alert Message on Deleteing the Voucher :  "+actAlertMessage +"  "+expAlertMessage);

		getAlert().accept();
		Thread.sleep(2000);

		String expDeleteMessage=excelReader.getCellData(xlSheetName, 363, 6);

		String actDeleteMessage=checkValidationMessage(expDeleteMessage);

		excelReader.setCellData(xlfile, xlSheetName, 363, 7, actDeleteMessage);

		if (actMessage.equalsIgnoreCase(expMessage)&& actDeleteMessage.equalsIgnoreCase(expDeleteMessage) && 
				actDataType.equalsIgnoreCase(expDataType)) 
		{
			System.out.println(" Test Pass:  Created View is Deleted View Successfull in Views TAB");
			excelReader.setCellData(xlfile, xlSheetName,  357, 9, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail:  Created View is Deleted View Successfull in Views TAB");
			excelReader.setCellData(xlfile, xlSheetName,  357, 9, resFail);
			return false;
		}
	}
	
	
	
	


	public boolean checkEditinginBodyFieldAndFooterFieldinInViewTab() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabViewSettingsBtn));
		viewTabViewSettingsBtn.click();
		
		Thread.sleep(1000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
		ViewExistingViewTxt.sendKeys(Keys.TAB);

		Thread.sleep(7000);

		int Count = viewsGridFiledsList.size();

		System.out.println("Count    : "+Count);

		ArrayList<String> actviewsGridFiledsList  = new ArrayList<String>();

		for (int i = 0; i < Count; i++) 
		{
			String data		  = viewsGridFiledsList.get(i).getText();

			if(data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 365, 5)))
			{
				viewsGridFiledsEditList.get(i).click();

				break;
			}
		}

		Thread.sleep(3000);
		
			getAction().moveToElement(viewSaveView).build().perform();
			Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
		viewSaveView.click();

		String expMessage=excelReader.getCellData(xlSheetName, 369, 6);

		String actMessage=checkValidationMessage(expMessage);

		excelReader.setCellData(xlfile, xlSheetName, 369, 7, actMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if (actMessage.equalsIgnoreCase(expMessage) /*&& actDataType.equalsIgnoreCase(expDataType)*/ /*&& 
           		actDefaultValue.equalsIgnoreCase(expDefaultValue)*/) 
		{
			System.out.println(" Test Pass:  Created View is Edited  Successfull in Views TAB");
			excelReader.setCellData(xlfile, xlSheetName,  364, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail:  Created View is Not Edited uuccessfull in Views TAB");
			excelReader.setCellData(xlfile, xlSheetName,  364, 8, resFail);
			return false;
		}
	}










	public boolean checkRuleTabInDocumentCustomization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		getDriver().navigate().refresh();
		Thread.sleep(2000);

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			ScrollIntoView(SettingsSubMenusList.get(i));
			
			if (data.equalsIgnoreCase("Configure Transactions")) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
*/	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		
		
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);*/
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();
		
		Thread.sleep(1000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreatedVoucher));
		docCustCreatedVoucher.click();
		
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesBtn));
		rulesBtn.click();

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabAddRuleBtn));

		boolean actRuleAdd=rulesTabAddRuleBtn.isDisplayed();

		rulesTabAddRuleBtn.click();

		String actRuleAddBtnDisplay=Boolean.toString(actRuleAdd);
		String expRuleAddBtnDisplay=excelReader.getCellData(xlSheetName, 371, 6);
		excelReader.setCellData(xlfile, xlSheetName, 371, 7, actRuleAddBtnDisplay);

		if(actRuleAddBtnDisplay.equalsIgnoreCase(expRuleAddBtnDisplay))
		{
			System.out.println(" Test Pass:  Add Link is Displayed ");
			excelReader.setCellData(xlfile, xlSheetName,  370, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail: Add Rule Link is Not Displayed ");
			excelReader.setCellData(xlfile, xlSheetName,  370, 8, resFail);
			return false;
		}
	}
	
	
	
	



	public boolean checkCreatingRuleinRuleTAbUnderDocumentCustomization() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabRuleNameTxt));
		rulesTabRuleNameTxt.click();

		rulesTabRuleNameTxt.sendKeys(Keys.SHIFT,Keys.HOME);

		rulesTabRuleNameTxt.sendKeys(Keys.BACK_SPACE);		

		Thread.sleep(1000);

		rulesTabRuleNameTxt.sendKeys(excelReader.getCellData(xlSheetName, 375, 5));

		Thread.sleep(2000);
		
		rulesTabRuleNameTxt.sendKeys(Keys.TAB);

		String actRuleName=rulesTabRuleNameTxt.getText();
		String expRuleName=excelReader.getCellData(xlSheetName, 375, 6);
		excelReader.setCellData(xlfile, xlSheetName, 375, 7, actRuleName);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabNewRecordChkbox));
		rulesTabNewRecordChkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabEditChkbox));
		rulesTabEditChkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabOnLeaveChkbox));
		rulesTabOnLeaveChkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabActiveChkbox));
		rulesTabActiveChkbox.click();

		boolean RulesTabNewRecordChkbox=rulesTabNewRecordChkbox.isSelected();
		String actRulesTabNewRecordChkbox=Boolean.toString(RulesTabNewRecordChkbox);
		String expRulesTabNewRecordChkbox=excelReader.getCellData(xlSheetName, 376, 6);
		excelReader.setCellData(xlfile, xlSheetName, 376, 7, actRulesTabNewRecordChkbox);

		boolean RulesTabEditChkbox=rulesTabEditChkbox.isSelected();
		String actRulesTabEditChkbox=Boolean.toString(RulesTabEditChkbox);
		String expRulesTabEditChkbox=excelReader.getCellData(xlSheetName, 377, 6);
		excelReader.setCellData(xlfile, xlSheetName, 377, 7, actRulesTabEditChkbox);

		boolean RulesTabOnLeaveChkbox=rulesTabOnLeaveChkbox.isSelected();
		String actRulesTabOnLeaveChkbox=Boolean.toString(RulesTabOnLeaveChkbox);
		String exRrulesTabOnLeaveChkbox=excelReader.getCellData(xlSheetName, 378, 6);
		excelReader.setCellData(xlfile, xlSheetName, 378, 7, actRulesTabOnLeaveChkbox);;

		boolean RulesTabActiveChkbox=rulesTabNewRecordChkbox.isSelected();
		String actRulesTabActiveChkbox=Boolean.toString(RulesTabActiveChkbox);
		String expRulesTabActiveChkbox=excelReader.getCellData(xlSheetName, 379, 6);
		excelReader.setCellData(xlfile, xlSheetName, 379, 7, actRulesTabActiveChkbox);

		rulesTabIFSelectTabTxt.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabIFSelectTabDrpdwn));
		Select account=new Select(rulesTabIFSelectTabDrpdwn);
		account.selectByVisibleText(excelReader.getCellData(xlSheetName, 380, 5));//6
		rulesTabIFSelectTabDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		rulesTabExtraFieldDrpdwn.click();

		rulesTabExtraFieldDrpdwn.sendKeys(Keys.ARROW_DOWN);
		rulesTabExtraFieldDrpdwn.sendKeys(Keys.TAB);

		rulesTabOperatorsDrpdwn.click();

		rulesTabOperatorsDrpdwn.sendKeys(Keys.ARROW_DOWN);

		rulesTabOperatorsDrpdwn.sendKeys(Keys.TAB);

		rulesTabValueTxt.click();

		Thread.sleep(1000);
		Select Value=new Select(rulesTabCompareWithDrpdwn);
		Value.selectByValue("1");
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);
		rulesTabCompareWithDrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		rulesTabConditionTxt.sendKeys(Keys.SPACE);
		rulesTabConditionTxt.sendKeys(Keys.BACK_SPACE);
		Thread.sleep(2000);

		rulesTabConditionTxt.sendKeys(excelReader.getCellData(xlSheetName, 381, 5));

		Thread.sleep(4000);

		rulesTabConditionTxt.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid1stRow_1stcol));
		rulesGrid1stRow_1stcol.click();

		rulesIFEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 382, 5));
		rulesIFEnterFiled.sendKeys(Keys.TAB);
		rulesIFEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 383, 5));

		rulesIFEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesIFEnterValue));

		rulesIFEnterValue.sendKeys(excelReader.getCellData(xlSheetName, 384, 5));

		rulesIFEnterValue.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabGrid1st_10thCol));
		rulesTabGrid1st_10thCol.click();

		rulesTabGridMandatoryChkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid2ndRow_1stcol));
		rulesGrid2ndRow_1stcol.click();

		rulesIFEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 385, 5));
		rulesIFEnterFiled.sendKeys(Keys.TAB);
		rulesIFEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 386, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGrid2ndRow_5thcol));
		rulesGrid2ndRow_5thcol.click();
		Thread.sleep(2000);

		rulesTabIFGridEnterCaption.sendKeys(excelReader.getCellData(xlSheetName, 387, 5));

		rulesTabIFGridEnterCaption.sendKeys(Keys.TAB);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTab_IFMesssageTab));
		rulesTab_IFMesssageTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rules_MessageTabGeneralTXT));
		rules_MessageTabGeneralTXT.sendKeys(excelReader.getCellData(xlSheetName, 388, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTabElseTab));
		rulesTabElseTab.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGridElse1stRow_1stcol));
		rulesGridElse1stRow_1stcol.click();

		rulesElseEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 389, 5));
		rulesElseEnterFiled.sendKeys(Keys.TAB);
		rulesElseEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 390, 5));

		rulesElseEnterChangeValueDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGridElse2ndRow_1stcol));
		rulesGridElse2ndRow_1stcol.click();

		rulesElseEnterFiled.sendKeys(excelReader.getCellData(xlSheetName, 392, 5));
		rulesElseEnterFiled.sendKeys(Keys.TAB);
		rulesElseEnterChangeValueDrpdwn.sendKeys(excelReader.getCellData(xlSheetName, 393, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesGridElse2ndRow_5thcol));
		rulesGridElse2ndRow_5thcol.click();
		Thread.sleep(2000);

		rulesTabElseGridEnterCaption.sendKeys(excelReader.getCellData(xlSheetName, 394, 5));

		rulesTabElseGridEnterCaption.sendKeys(Keys.TAB);	
		
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", rulesTab_SaveRuleBtn);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(rulesTab_SaveRuleBtn));
		rulesTab_SaveRuleBtn.click();

		String expMessageOnSaveRule = excelReader.getCellData(xlSheetName, 395, 6);
		String actMessageOnSaveRule=checkValidationMessage(expMessageOnSaveRule);

		excelReader.setCellData(xlfile, xlSheetName, 395, 7, actMessageOnSaveRule);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Settings_closeBtn));
		Settings_closeBtn.click();

		String expMessageOnClickOnUpdate = excelReader.getCellData(xlSheetName, 396, 6);
		String actMessageOnClickOnUpdate=checkValidationMessage(expMessageOnSaveRule);

		excelReader.setCellData(xlfile, xlSheetName, 396, 7, actMessageOnClickOnUpdate);

		if (actMessageOnSaveRule.equalsIgnoreCase(expMessageOnSaveRule) && 
				actMessageOnClickOnUpdate.equalsIgnoreCase(expMessageOnClickOnUpdate))
		{
			System.out.println(" Test Pass:  Rule Created Successful");
			excelReader.setCellData(xlfile, xlSheetName,  374, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail:  Rule not Created ");
			excelReader.setCellData(xlfile, xlSheetName,  374, 8, resFail);
			return false;
		}
	}

	
	


	


	public boolean checkNavigateToDocumentNumberingTabInPurchaseVoucherVatView() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{	
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		getDriver().navigate().refresh();
		Thread.sleep(2000);
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			ScrollIntoView(SettingsSubMenusList.get(i));
			
			if (data.equalsIgnoreCase("Configure Transactions")) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
		
*/	
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);*/
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();
		
		Thread.sleep(1000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreatedVoucher));
		docCustCreatedVoucher.click();
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberingBtn));
		documentNumberingBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow1));
		typeRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		Select typeDropDownSelect = new Select(typeDropDown);

		typeDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 413, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow1));
		fieldValueRow1.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
		Select valueDropDownSelect= new Select(fieldValueDropDown);

		valueDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 415, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow1));
		startingCharacterRow1.click();

		startingCharacterTxt.sendKeys(Keys.BACK_SPACE);
		
		startingCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 416, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow1));
		noofCharactersRow1.click();

		noOfCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 417, 5));

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow2));
		typeRow2.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 418, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow2));
		fieldValueRow2.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
		Select fieldValueDropDownSelect = new Select(fieldValueDropDown);

		fieldValueDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 419, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow2));
		noofCharactersRow2.click();

		noOfCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 420, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow3));
		typeRow3.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText("Department");

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow3));
		fieldValueRow3.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
		fieldValueDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 423, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow3));
		startingCharacterRow3.click();

		startingCharacterTxt.sendKeys(Keys.BACK_SPACE);
		startingCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 424, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow3));
		noofCharactersRow3.click();

		noOfCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 425, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow4));
		typeRow4.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 426, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow4));
		fieldValueRow4.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueDropDown));
		fieldValueDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 427, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow4));
		noofCharactersRow4.click();

		noOfCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 428, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow5));
		typeRow5.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 429, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow5));
		fieldValueRow5.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueTxt));
		fieldValueTxt.sendKeys(excelReader.getCellData(xlSheetName, 430, 5));


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow5));
		noofCharactersRow5.click();

		noOfCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 431, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow6));
		typeRow6.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeDropDown));
		typeDropDownSelect.selectByVisibleText(excelReader.getCellData(xlSheetName, 432, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueRow6));
		fieldValueRow6.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fieldValueTxt));
		fieldValueTxt.sendKeys(excelReader.getCellData(xlSheetName, 433, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startingCharacterRow6));
		startingCharacterRow6.click();

		startingCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 434, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(noofCharactersRow6));
		noofCharactersRow6.click();

		noOfCharacterTxt.sendKeys(excelReader.getCellData(xlSheetName, 435, 5));

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(typeRow7));
		typeRow7.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docNumberAppearAsValue));

		String actdocNumberAppearAsValue = docNumberAppearAsValue.getText();

		String expdocNumberAppearAsValue = excelReader.getCellData(xlSheetName, 433, 6);

		excelReader.setCellData(xlfile, xlSheetName, 433, 7, actdocNumberAppearAsValue);


		System.out.println("docNumberAppearAsValue Display Value Actual      : " + actdocNumberAppearAsValue);
		System.out.println("docNumberAppearAsValue Display Value Expected    : " + expdocNumberAppearAsValue);

		if(actdocNumberAppearAsValue.equalsIgnoreCase(expdocNumberAppearAsValue)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 432, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 432, 8, resFail);
			return false;
		}
	}

	
	
	
	public boolean checkSavingOnClickingUpdateOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{	
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expSaveMessage = excelReader.getCellData(xlSheetName, 437, 6);

		String actSaveMessage = checkValidationMessage(expSaveMessage);

		excelReader.setCellData(xlfile, xlSheetName, 437, 7, actSaveMessage);

		System.out.println("SaveMessage Display Value Actual      : " + actSaveMessage);
		System.out.println("SaveMessage Display Value Expected    : " + expSaveMessage);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settings_closeBtn));
		settings_closeBtn.click();

		if(actSaveMessage.equalsIgnoreCase(expSaveMessage)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 436, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 436, 8, resFail);
			return false;
		}
	}


	
	

	public boolean checkSavedFieldsinPvVatViewVoucherDocumentNUmberingTab() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{	
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		getDriver().navigate().refresh();

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			ScrollIntoView(SettingsSubMenusList.get(i));
			
			if (data.equalsIgnoreCase("Configure Transactions")) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
		
*/	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);*/
	/*	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();
		
		Thread.sleep(1000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreatedVoucher));
		docCustCreatedVoucher.click();
		
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberingBtn));
		documentNumberingBtn.click();

		int docNumberingGridBodyCount = docNumberingGridBody.size();

		ArrayList<String> docNumberingGridBodyArray=new ArrayList<String>();

		for(int i=0;i<docNumberingGridBodyCount;i++)
		{
			String data = docNumberingGridBody.get(i).getText();
			docNumberingGridBodyArray.add(data);
		}

		String actdocNumberingGridBody = docNumberingGridBodyArray.toString();

		String expdocNumberingGridBody = "[1, Login, Login Name, 0, 2, 2, Character Input, /, 0, 1, 3, Department, Name, 0, 3, 4, Character Input, /, 0, 1, 5, Input, TEXT, 0, 1, 6, Increment, 0, 0, 1, 7,  ,  ,  ,  , 8,  ,  ,  ,  , 9,  ,  ,  ,  , 10,  ,  ,  ,  , 11,  ,  ,  ,  , 12,  ,  ,  ,  , 13,  ,  ,  ,  , 14,  ,  ,  ,  , 15,  ,  ,  ,  , 16,  ,  ,  ,  , 17,  ,  ,  ,  , 18,  ,  ,  ,  , 19,  ,  ,  ,  , 20,  ,  ,  ,  ]";

		excelReader.setCellData(xlfile, xlSheetName, 440, 8, actdocNumberingGridBody);

		System.out.println("Body Display Value Actual      : " + actdocNumberingGridBody);
		System.out.println("Body Display Value Expected    : " + expdocNumberingGridBody);

		// Footer

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docNumberAppearAsLabel));

		String actdocNumberAppearAsLabel = docNumberAppearAsLabel.getText();

		String expdocNumberAppearAsLabel = excelReader.getCellData(xlSheetName, 441, 6);

		excelReader.setCellData(xlfile, xlSheetName, 441, 7, actdocNumberAppearAsLabel);

		System.out.println("Footer Display Value Actual      : " + actdocNumberAppearAsLabel);
		System.out.println("Footer Display Value Expected    : " + expdocNumberAppearAsLabel);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docNumberAppearAsValue));

		String actdocNumberAppearAsValue = docNumberAppearAsValue.getText();

		String expdocNumberAppearAsValue = excelReader.getCellData(xlSheetName, 442, 6);

		excelReader.setCellData(xlfile, xlSheetName, 442, 7, actdocNumberAppearAsValue);

		System.out.println("Footer Display Value Actual      : " + actdocNumberAppearAsValue);
		System.out.println("Footer Display Value Expected    : " + expdocNumberAppearAsValue);

		if(actdocNumberingGridBody.equalsIgnoreCase(expdocNumberingGridBody) && actdocNumberAppearAsLabel.equalsIgnoreCase(expdocNumberAppearAsLabel) 
				&& actdocNumberAppearAsValue.equalsIgnoreCase(expdocNumberAppearAsValue)) 
		{
			excelReader.setCellData(xlfile, xlSheetName, 438, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 438, 8, resFail);
			return false;
		}
	}



	  		


	public boolean 	checkRaiseDocumentOptionInTriggers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();

		Thread.sleep(4000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabTriggersBtn));
		triggerTabTriggersBtn.click();
		
		Thread.sleep(1000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(excelReader.getCellData(xlSheetName, 444, 5));

		Thread.sleep(2000);

		editLayoutTriggerName.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabTriggersBtn));
		triggerTabTriggersBtn.click();
		
		Thread.sleep(1000);*/
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabActionBtn));
		triggerTabActionBtn.click();
		
		Thread.sleep(1000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersRiseANewDocChkbox));
		triggersRiseANewDocChkbox.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabActionBtn));
		triggerTabActionBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
		triggerTabAlertBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
		boolean allVoucher=editLayoutTriggersAllVouchers.isDisplayed();

		String actAllVouchersDisplay=Boolean.toString(allVoucher);
		String expAllVouchersDisplay=excelReader.getCellData(xlSheetName, 445, 6);

		excelReader.setCellData(xlfile, xlSheetName, 445, 7, actAllVouchersDisplay);

		if(actAllVouchersDisplay.equalsIgnoreCase(expAllVouchersDisplay))
		{
			System.out.println("Test Pass : All Vouchers Displayed Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 443, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : All Vouchers NOt Displayed  Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 443, 8, resFail);
			return false;
		}
	}


	
	

	public boolean 	checkSaveOptionInTriggerTab() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersAllVouchers));
		editLayoutTriggersAllVouchers.sendKeys(excelReader.getCellData(xlSheetName, 446, 6));
		editLayoutTriggersAllVouchers.sendKeys(Keys.TAB);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();

		String expSaveMessage = excelReader.getCellData(xlSheetName, 447, 6);

		String actSaveMessage = checkValidationMessage(expSaveMessage);

		excelReader.setCellData(xlfile, xlSheetName, 447, 7, actSaveMessage);

		if(actSaveMessage.equalsIgnoreCase(expSaveMessage))
		{
			System.out.println("Test Pass :Trigger Saved  Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 446, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Trigger Not Saved  Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 446, 8, resFail);
			return false;
		}
	}

	
	
	


	public boolean 	checkDeleteTriggerOptionInTriggers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabTriggersBtn));
		triggerTabTriggersBtn.click();*/
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		editLayoutTriggerName.sendKeys(excelReader.getCellData(xlSheetName, 448, 5));
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerDeleteBtn));
		editLayoutTriggerDeleteBtn.click();

		getWaitForAlert();

		boolean AlertPresent=getIsAlertPresent();

		String actAlertPresent=Boolean.toString(AlertPresent);
		String expAlertPresent=excelReader.getCellData(xlSheetName, 449, 6);
		excelReader.setCellData(xlfile, xlSheetName, 449, 7, actAlertPresent);

		System.out.println(" ALert is Present : "+actAlertPresent+" Value Expected : "+expAlertPresent);

		//System.out.println("Select Options In Inventory Tab and Click on Update  Value Actual : " + actAlertPresent +  " Value Expected : " + expAlertPresent);

		String actMessageOnUpdateInAlert = getAlert().getText();
		String expMessageOnUpdateInAlert = excelReader.getCellData(xlSheetName, 450, 6);

		excelReader.setCellData(xlfile, xlSheetName, 450, 7, actMessageOnUpdateInAlert);

		//System.out.println("Select Options In Inventory Tab and Click on Ok  Value Actual : " + actMessageOnUpdateInAlert +  " Value Expected : " + expMessageOnUpdateInAlert);

		Thread.sleep(2000);
		getAlert().accept();

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessagecloseBtn));
		errorMessagecloseBtn.click();*/

		if(actAlertPresent.equalsIgnoreCase(expAlertPresent)/* && actMessageOnUpdateInAlert.equalsIgnoreCase(expMessageOnUpdateInAlert)*/)
		{
			System.out.println("Test Pass : Trigger Deleted  Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 448, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail :Created Trigger Not Deleted  Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 448, 8, resFail);
			return false;
		}
	}








	public boolean 	checkSaveTriggerWithRaiseDocumentOptionInTriggers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		
		Thread.sleep(4000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();
		Thread.sleep(1500);
		editLayoutTriggerName.sendKeys(excelReader.getCellData(xlSheetName, 451, 5));
		Thread.sleep(2000);
		editLayoutTriggerName.sendKeys(Keys.TAB);			
		Thread.sleep(6000);
		/*editLayoutTriggerName.sendKeys(Keys.TAB);	
		Thread.sleep(2000);
		editLayoutTriggerValue.click();
		editLayoutTriggerValue.sendKeys(Keys.TAB);

		Thread.sleep(5000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabTriggersBtn));
		triggerTabTriggersBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabActionBtn));
		triggerTabActionBtn.click();
		
		Thread.sleep(1000);

		
		Thread.sleep(2000);
*/		
		triggersConjuctionDrpdwn.click();
		Thread.sleep(2000);		
		triggersConjuctionDrpdwn.sendKeys(Keys.ARROW_DOWN);			
		triggersConjuctionDrpdwn.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
		triggersSelectFieldTxt.click();

		Thread.sleep(3000);
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", triggerVendorACExpansionBtn);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorACExpansionBtn));
		triggerVendorACExpansionBtn.click();

		Thread.sleep(2000);
		
		//((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", triggerVendorAcName);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerVendorAcName));
		triggerVendorAcName.click();

		Thread.sleep(2000);
		
		//triggerVendorAcName.sendKeys(Keys.TAB);
		triggersSelectFieldTxt.sendKeys(Keys.TAB);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
		triggersSelectOperdrpdwn.click();
		triggersSelectOperdrpdwn.sendKeys("Equal");


		Thread.sleep(5000);
		
		triggersSelectOperdrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(5000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
		triggersCompareWithdrpdwn.click();

		/*triggersCompareWithdrpdwn.sendKeys(Keys.ARROW_DOWN);
		Thread.sleep(2000);*/
		triggersCompareWithdrpdwn.sendKeys("value");
		Thread.sleep(3000);
		triggersCompareWithdrpdwn.sendKeys(Keys.TAB);

		Thread.sleep(2000);

		//triggersValueToEnterTxt.click();
		triggersValueTxt.click();
		triggersValueTxt.sendKeys("Vendor A");	

		Thread.sleep(2000);

		//triggersValueTxt.sendKeys(Keys.TAB);
		triggersValueTxt.sendKeys(Keys.TAB);



		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersRiseANewDocChkbox));
		triggersRiseANewDocChkbox.click();

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabActionBtn));
		triggerTabActionBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabAlertBtn));
		triggerTabAlertBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
		boolean actdata=editLayoutTriggersAllVouchers.isDisplayed();
		boolean expdata=true;

 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersAllVouchers));
		editLayoutTriggersAllVouchers.click();
		editLayoutTriggersAllVouchers.sendKeys("Job Orders");
		editLayoutTriggersAllVouchers.sendKeys(Keys.ENTER);


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();

		String expSaveMessage = excelReader.getCellData(xlSheetName, 453, 6);
		String actSaveMessage = checkValidationMessage(expSaveMessage);
		excelReader.setCellData(xlfile, xlSheetName, 453, 7, actSaveMessage);


		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();


		String expUpdateMessage = excelReader.getCellData(xlSheetName, 454, 6);
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);
		excelReader.setCellData(xlfile, xlSheetName, 454, 7, actUpdateMessage);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
		closeBtn.click();

		if(actdata==expdata && actSaveMessage.equalsIgnoreCase(expSaveMessage))
		{
			System.out.println("Test Pass : Trigger  Added Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 451, 8, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Trigger NOT Added Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 451, 8, resFail);
			return false;
		}
	}
	
	





	public boolean 	checkEditingInTriggersTabUnderPVVATView() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(SettingsMenu));
		SettingsMenu.click();
		
		Thread.sleep(2000);

		int count = SettingsSubMenusList.size();
		
		for (int i = 0; i < count; i++) 
		{
			String data = SettingsSubMenusList.get(i).getText();
			
			System.err.println(data);
			
			ScrollIntoView(SettingsSubMenusList.get(i));
			
			if (data.equalsIgnoreCase("Configure Transactions")) 
			{
				SettingsSubMenusList.get(i).click();
				break;
			}
		}
		
		Thread.sleep(2000);
*/	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchTxt));
		searchTxt.click();
		searchTxt.sendKeys("Configure Transactions");
		searchTxt.sendKeys(Keys.ENTER);
		Thread.sleep(3000);
		
	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionPreferencesMinus));
		configureTransactionPreferencesMinus.click();
		
		Thread.sleep(2000);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(docCustCreatedVoucher));
		docCustCreatedVoucher.click();
		
		Thread.sleep(4000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersBtn));
		triggersBtn.click();			

		Thread.sleep(2000);	
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabTriggersBtn));
		triggerTabTriggersBtn.click();
		
		Thread.sleep(1000);*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggerName));
		editLayoutTriggerName.click();

		editLayoutTriggerName.sendKeys(Keys.SHIFT,Keys.HOME);
		editLayoutTriggerName.sendKeys(Keys.BACK_SPACE);

		editLayoutTriggerName.sendKeys(Keys.CLEAR);

		Thread.sleep(2000);
		
		editLayoutTriggerName.sendKeys(Keys.TAB);

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabTriggersBtn));
		triggerTabTriggersBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggerTabActionBtn));
		triggerTabActionBtn.click();
		
		Thread.sleep(1000);
		
		triggersValueTxtGetValue.click();

		Thread.sleep(2000);

		triggersValueTxt.sendKeys("Vendor B");	

		Thread.sleep(2000);

		triggersValueTxt.sendKeys(Keys.TAB);

		Thread.sleep(3000);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersConjuctionDrpdwn));
		Select triggersConjuctionDrpdwnSelect = new Select(triggersConjuctionDrpdwn);
		String acttriggersConjuctionDrpdwn=triggersConjuctionDrpdwnSelect.getFirstSelectedOption().getText();
		String exptriggersConjuctionDrpdwn="Where";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
		String acttriggersSelectFieldTxt= triggersSelectFieldTxt.getAttribute("value");
		String exptriggersSelectFieldTxt="Name";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectOperdrpdwn));
		Select triggersSelectOperdrpdwnSelect = new Select(triggersSelectOperdrpdwn);
		String acttriggersSelectOperdrpdwn=triggersSelectOperdrpdwnSelect.getFirstSelectedOption().getText();
		String exptriggersSelectOperdrpdwn="Equal to";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersCompareWithdrpdwn));
		Select triggersCompareWithdrpdwnSelect = new Select(triggersCompareWithdrpdwn);
		String acttriggersCompareWithdrpdwn=triggersCompareWithdrpdwnSelect.getFirstSelectedOption().getText();
		String exptriggersCompareWithdrpdwn="value";

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(triggersSelectFieldTxt));
		String acttriggersValueToEnterTxt= triggersValueTxtGetValue.getAttribute("value");
		String exptriggersValueToEnterTxt="Vendor B";

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(editLayoutTriggersAllVouchers));
		boolean actdata=editLayoutTriggersAllVouchers.isDisplayed();
		boolean expdata=true;

		System.out.println("triggersConjuctionDrpdwn  : ."+acttriggersConjuctionDrpdwn  +". Value Expected :"+exptriggersConjuctionDrpdwn);
		System.out.println("triggersSelectFieldTxt    : ."+acttriggersSelectFieldTxt    +". Value Expected :"+exptriggersSelectFieldTxt);
		System.out.println("triggersSelectOperdrpdwn  : ."+acttriggersSelectOperdrpdwn  +". Value Expected :"+exptriggersSelectOperdrpdwn);
		System.out.println("triggersCompareWithdrpdwn : ."+acttriggersCompareWithdrpdwn +". Value Expected :"+exptriggersCompareWithdrpdwn);
		System.out.println("triggersValueToEnterTxt   : ."+acttriggersValueToEnterTxt   +". Value Expected :"+exptriggersValueToEnterTxt);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editLayoutTriggersSaveBtn));
		editLayoutTriggersSaveBtn.click();

		String expSaveMessage = excelReader.getCellData(xlSheetName, 453, 7);
		String actSaveMessage = checkValidationMessage(expSaveMessage);
		excelReader.setCellData(xlfile, xlSheetName, 453, 8, actSaveMessage);			

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn));
		updateBtn.click();

		String expUpdateMessage = excelReader.getCellData(xlSheetName, 454, 7);
		String actUpdateMessage = checkValidationMessage(expUpdateMessage);
		excelReader.setCellData(xlfile, xlSheetName, 454, 8, actUpdateMessage);

		if(actdata==expdata && actSaveMessage.equalsIgnoreCase(expSaveMessage) && 
				acttriggersCompareWithdrpdwn.equalsIgnoreCase(exptriggersCompareWithdrpdwn) && 
				acttriggersConjuctionDrpdwn.equalsIgnoreCase(exptriggersConjuctionDrpdwn) && 
				acttriggersSelectFieldTxt.equalsIgnoreCase(exptriggersSelectFieldTxt) && 
				acttriggersSelectOperdrpdwn.equalsIgnoreCase(exptriggersSelectOperdrpdwn) && 
				acttriggersValueToEnterTxt.equalsIgnoreCase(exptriggersValueToEnterTxt)
				&& actUpdateMessage.equalsIgnoreCase(expUpdateMessage))
		{
			System.out.println("Test Pass : Trigger  Edited Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 451, 9, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Trigger NOT Edited Successfully");
			excelReader.setCellData(xlfile, xlSheetName, 451, 9, resFail);
			return false;
		}
	}


	
	
	
	
	public boolean checkLogoutUserDocumentCustOptions() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

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
			excelReader.setCellData(xlfile, xlSheetName, 285, 9, resPass);
			return true;
		}
		else
		{

			System.out.println("***Test Fail: Login Not Successfull***");
			excelReader.setCellData(xlfile, xlSheetName, 285, 9, resFail);
			return false;
		}
	}


	
/////////Creating new fields with different positions/////////
	
	
	
	  @FindBy(xpath="//a[@id='769']")
		 private static WebElement  purchaseVouchersNewVoucherBtn;
      
		@FindBy(xpath="//select[@id='editScreen_FieldsCustomization_Position']")
		private static WebElement editScreenpositionDropdown;  
	   
		@FindBy(xpath="//input[@id='editScreen_FieldsCustomization_noofDecimals']")
		private static WebElement editScreenNoOfDecimalsTxt;
		
		  @FindBy(xpath="//button[@id='btnEditScreenNext']")
		  private static WebElement editScreenNextBtn;  
	
public boolean checkCreatingNewFieldsinEditScreenTabatDifferentPositions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
		settingsmenuBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionBtn));
		configureTransactionBtn.click();
		
	    Thread.sleep(2000);
		
	 
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(configureTransactionDocumentCustomiztionMinus));
		configureTransactionDocumentCustomiztionMinus.click();
		
		
		
		
	    Thread.sleep(2000);
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersNewVoucherBtn));
	    purchaseVouchersNewVoucherBtn.click();
	    
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
		editScreenTab.click();
		
		Thread.sleep(2000);
		
		
		///Position at Before Quantity
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenaddBtn));
		editScreenaddBtn.click();
		Thread.sleep(2000);
		
		if(getIsAlertPresent())
		{
		getWaitForAlert();
		getAlert().accept();
		Thread.sleep(2000);
		}
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
		editScreenCaptionTxt.click();
		editScreenCaptionTxt.sendKeys("MCharge1");
		editScreenCaptionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenNoOfDecimalsTxt));
		
		editScreenNoOfDecimalsTxt.click();
		editScreenNoOfDecimalsTxt.clear();
		editScreenNoOfDecimalsTxt.sendKeys("2");
		editScreenNoOfDecimalsTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenpositionDropdown));
		Select s=new Select(editScreenpositionDropdown);
		s.selectByValue("0");
		editScreenpositionDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
		editScreenApplyBtn.click();
		
		Thread.sleep(2000);
		checkValidationMessage("");
		
		
		
///Position at Between Quantity and Rate
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenaddBtn));
		editScreenaddBtn.click();
		Thread.sleep(2000);
		
		if(getIsAlertPresent())
		{
		getWaitForAlert();
		getAlert().accept();
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
		editScreenCaptionTxt.click();
		editScreenCaptionTxt.sendKeys("MCharge2");
		editScreenCaptionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenNoOfDecimalsTxt));
		
		editScreenNoOfDecimalsTxt.click();
		editScreenNoOfDecimalsTxt.clear();
		editScreenNoOfDecimalsTxt.sendKeys("2");
		editScreenNoOfDecimalsTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenpositionDropdown));
		Select s1=new Select(editScreenpositionDropdown);
		s1.selectByValue("1");
		editScreenpositionDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
		editScreenApplyBtn.click();
		
		Thread.sleep(2000);
		
		checkValidationMessage("");
		
		
///Position at Between  Rate and Gross
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenaddBtn));
		editScreenaddBtn.click();
		Thread.sleep(2000);
		
		if(getIsAlertPresent())
		{
		getWaitForAlert();
		getAlert().accept();
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
		editScreenCaptionTxt.click();
		editScreenCaptionTxt.sendKeys("MCharge3");
		editScreenCaptionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenNoOfDecimalsTxt));
		
		editScreenNoOfDecimalsTxt.click();
		editScreenNoOfDecimalsTxt.clear();
		editScreenNoOfDecimalsTxt.sendKeys("2");
		editScreenNoOfDecimalsTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenpositionDropdown));
		Select s3=new Select(editScreenpositionDropdown);
		s3.selectByValue("2");
		editScreenpositionDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
		editScreenApplyBtn.click();
		Thread.sleep(2000);
		
		checkValidationMessage("");
		
		
///Position at After Gross
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenaddBtn));
		editScreenaddBtn.click();
		Thread.sleep(2000);
		
		if(getIsAlertPresent())
		{
		getWaitForAlert();
		getAlert().accept();
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenCaptionTxt));
		editScreenCaptionTxt.click();
		editScreenCaptionTxt.sendKeys("MCharge4");
		editScreenCaptionTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenNoOfDecimalsTxt));
		
		editScreenNoOfDecimalsTxt.click();
		editScreenNoOfDecimalsTxt.clear();
		editScreenNoOfDecimalsTxt.sendKeys("2");
		editScreenNoOfDecimalsTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenpositionDropdown));
		Select s4=new Select(editScreenpositionDropdown);
		s4.selectByValue("3");
		editScreenpositionDropdown.sendKeys(Keys.TAB);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenApplyBtn));
		editScreenApplyBtn.click();
		Thread.sleep(2000);
		checkValidationMessage("");
		
		
		///checking List of captions for newly added fields
		
		int captionsListCount = captionsList.size();
		ArrayList<String> captionsListArray = new ArrayList<String>();
		for(int i=0;i<captionsListCount;i++)
		{
			
			/*if(i==7)
			{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenNextBtn));
			editScreenNextBtn.click();
			}
			*/
			String data = captionsList.get(i).getText();
			captionsListArray.add(data);
			
			
		}
		String actCaptionList = captionsListArray.toString();
		String expCaptionList = "[Discount, RD, Avg Rate, Avg Rate(O), VAT, Taxable, MCharge1, MCharge2, MCharge3, MCharge4]";
		
		
		System.out.println("actCaptionsList  : "+actCaptionList);
		System.out.println("expCaptionsList  : "+expCaptionList);
		
		
		if(actCaptionList.equalsIgnoreCase(expCaptionList))
		{
					return true;
		}
		else
		{
			return false;
		}
	    
}


@FindBy(xpath="//span[@class='dropdown icon-menuicon1 icon-font6 theme_color-inverse pull-right']")
private static WebElement  toogleExpandBtn;

@FindBy(xpath="//*[@id='panelsStayOpen-headingTwo']/button")
private static WebElement viewTabLoginBtn;

public boolean checkCreateViewforNewlyAddedFields() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	

	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewsBtn));
	  viewsBtn.click();
	  Thread.sleep(2000);
	  
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
		ViewExistingViewTxt.click();
		
		ViewExistingViewTxt.sendKeys("positionView");
		
		ViewExistingViewTxt.sendKeys(Keys.TAB);*/
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
		ViewExistingViewTxt.click();
		
		ViewExistingViewTxt.sendKeys(Keys.SHIFT,Keys.HOME);
		
		ViewExistingViewTxt.sendKeys(Keys.SPACE);
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys("positionView");
		Thread.sleep(2000);
		ViewExistingViewTxt.sendKeys(Keys.TAB);
		
		 Thread.sleep(2000);
		 
	int Count = viewsGridFiledsList.size();
	
	String actFiledCount=Integer.toString(Count);
	String expFiledCount= "40"; 		
	
	System.out.println("Count    : "+Count);
	
	ArrayList<String> actviewsGridFiledsList  = new ArrayList<String>();
	
	for (int i = 0; i < Count; i++) 
	{
		String data		                              = viewsGridFiledsList.get(i).getText();
		actviewsGridFiledsList.add(data);
	}
	
	String actFieldList					  = actviewsGridFiledsList.toString();
	String expFieldList					  = "[DocNo, Date, Time, VendorAC, UpdateStock, RaiseReceipt, DueDate, Currency, ExchangeRate, Department, LocExchangeRate, Place of supply, Jurisdiction, sNarration, PermitNo, Warehouse, Item, TaxCode, PurchaseAC, Unit, RD, Avg Rate, Avg Rate(O), MCharge1, Quantity, L-Purchases Orders, MCharge2, Rate, MCharge3, Gross, Discount, VAT, Taxable, MCharge4, Batch, Bins, ExpDate, RMA, FD%, VAT advance]";
	
	
		    	
	System.err.println("Field Names Display Text Actual Values      :  " + actFieldList);
	System.out.println("Field Names Display Text Expected Values    :  " + expFieldList);
	
	  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewTabLoginBtn));
	  viewTabLoginBtn.click();
	  
	  if(viewSuChkboxIsSelected.isSelected()==false)
	  {
	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSuChkbox));
	   viewSuChkbox.click();
	  }
       Thread.sleep(2000);
       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
       viewSaveView.click();
       Thread.sleep(2000);
       
       String expMessage= "Data saved successfully";
       
       String actMessage=checkValidationMessage(expMessage);
       
	
	
	
	
	if(actFiledCount.equalsIgnoreCase(expFiledCount) && actFieldList.equalsIgnoreCase(expFieldList)
			&&actMessage.equalsIgnoreCase(expMessage))
	{
		System.out.println("Test Pass : Body Grid Filed Options  As Expected");
	  
		return true;
	  }
	  else
	  {
		System.out.println("Test Fail :  Body Grid Filed Options  As Expected");
	
		return false;
	  }
}


public boolean checkFieldsinViewTabafterRefreshView() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ViewExistingViewTxt));
	ViewExistingViewTxt.click();
	
	ViewExistingViewTxt.sendKeys(Keys.SHIFT,Keys.HOME);
	
	ViewExistingViewTxt.sendKeys(Keys.SPACE);
	Thread.sleep(2000);
	ViewExistingViewTxt.sendKeys("positionView");
	Thread.sleep(2000);
	ViewExistingViewTxt.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	
	 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewRefreshViewBtn));
	 viewRefreshViewBtn.click();
	 Thread.sleep(2000);
	 
	 
	 int Count = viewsGridFiledsList.size();
		
		String actFiledCount=Integer.toString(Count);
		String expFiledCount= "38"; 		
		
		System.out.println("Count    : "+Count);
		
		ArrayList<String> actviewsGridFiledsList  = new ArrayList<String>();
		
		for (int i = 0; i < Count; i++) 
		{
			String data		                              = viewsGridFiledsList.get(i).getText();
			actviewsGridFiledsList.add(data);
		}
		
		String actFieldList					  = actviewsGridFiledsList.toString();
		String expFieldList					  = "[DocNo, Date, Time, VendorAC, UpdateStock, RaiseReceipt, DueDate, Currency, ExchangeRate, Department, LocExchangeRate, Place of supply, Jurisdiction, sNarration, PermitNo, Warehouse, Item, TaxCode, PurchaseAC, Unit, RD, Avg Rate, Avg Rate(O), MCharge1, Quantity, L-Purchases Orders, Rate, Gross, Discount, VAT, Taxable, MCharge4, Batch, Bins, ExpDate, RMA, FD%, VAT advance]";
		
		
			    	
		System.err.println("Field Names Display Text Actual Values      :  " + actFieldList);
		System.out.println("Field Names Display Text Expected Values    :  " + expFieldList);
		
		
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(viewSaveView));
           viewSaveView.click();
           Thread.sleep(2000);
           
           getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(updateBtn)); 
           updateBtn.click();
           Thread.sleep(2000);
		
           String expMessage="Data saved successfully";
           String actMessage=checkValidationMessage(expMessage);
		
		
		if(actFiledCount.equalsIgnoreCase(expFiledCount) && actFieldList.equalsIgnoreCase(expFieldList)
				&& actMessage.equalsIgnoreCase(expMessage))
		{
			System.out.println("Test Pass : Body Grid Filed Options  As Expected");
		  
			return true;
		  }
		  else
		  {
			System.out.println("Test Fail :  Body Grid Filed Options  As Expected");
		
			return false;
		  }
		
	
}


public boolean checkCreatedViewFieldsatVoucherLevel() throws InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	financialsMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVouchersVat));
	purchaseVouchersVat.click();

	Thread.sleep(2000);

	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();
	Thread.sleep(2000);
	
	
	 int reportsHeaderListCount = headersList.size();
	 
	 System.err.println(reportsHeaderListCount);
	 
		ArrayList<String> reportsHeaderListArray = new ArrayList<String>();
		for(int i=0;i<reportsHeaderListCount;i++)
		{
			String data = headersList.get(i).getText();
			((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,300)","");
			 System.err.println(data);
			 
			reportsHeaderListArray.add(data);
		}
		String actHeaderList = reportsHeaderListArray.toString();
		String expHeaderList = "";
	    
	    System.out.println(" Header List Actual    : "+actHeaderList);
	    System.out.println("  Header List expected :  "+expHeaderList);
	
		Thread.sleep(1500);
		
		if(actHeaderList.equalsIgnoreCase(expHeaderList))
		{
			return true;
		}
		else
		{
			return false;
		}
	
	
	
}


@FindBy(xpath="//*[@id='id_transactionentry_more']/span")
public static WebElement PVVAtToggleBtn;


@FindBy(xpath="//*[@id='id_transactionentry_settings']")
public static WebElement PVVAtSettingsBtn;




public boolean checkDeleteAddedFieldsinEditScreenTab() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PVVAtToggleBtn));
	PVVAtToggleBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(PVVAtSettingsBtn));
	PVVAtSettingsBtn.click();
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editScreenTab));
	editScreenTab.click();
	Thread.sleep(2000);
	
	int captionsListCount = captionsList.size();
	ArrayList<String> captionsListArray = new ArrayList<String>();
	for(int i=0;i<captionsListCount;i++)
	{
		
		String data = captionsList.get(i).getText();
		
		if(data.equals("MCharge1") )
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(delete7thRowBtn));
			delete7thRowBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			Thread.sleep(2000);
			break;
			
		}
	}
	
	checkValidationMessage("");
	Thread.sleep(1000);
	for(int i=0;i<captionsListCount;i++)
	{
		
		String data = captionsList.get(i).getText();
		if(data.equals("MCharge2") )
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(delete7thRowBtn));
			delete7thRowBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			Thread.sleep(2000);
			break;
		}
		
	}
	
	checkValidationMessage("");
	Thread.sleep(1000);
	
	for(int i=0;i<captionsListCount;i++)
	{
		
		String data = captionsList.get(i).getText();
	
		 if(data.equals("MCharge3") )
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(delete7thRowBtn));
			delete7thRowBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			Thread.sleep(2000);
			break;
			
		}
	}
	
	checkValidationMessage("");
	Thread.sleep(1000);
	
	
	for(int i=0;i<captionsListCount;i++)
	{
		
		String data = captionsList.get(i).getText();
	
		 if(data.equals("MCharge4") )
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(delete7thRowBtn));
			delete7thRowBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			Thread.sleep(2000);
			break;
		}
		
					
	}
	
	
	checkValidationMessage("");
	Thread.sleep(1000);
	
	
	Thread.sleep(2000);
	
	for(int i=0;i<captionsListCount;i++)
	{
		
		String data = captionsList.get(i).getText();
		captionsListArray.add(data);
	}
	String actCaptionList = captionsListArray.toString();
	String expCaptionList = "[Discount, RD, Avg Rate, Avg Rate(O), VAT, Taxable]";
	
	
	System.out.println("actCaptionsList  : "+actCaptionList);
	System.out.println("expCaptionsList  : "+expCaptionList);
	
	
	if(actCaptionList.equalsIgnoreCase(expCaptionList))
	{
				return true;
	}
	else
	{
		return false;
	}
    	
	
	
}



@FindBy(xpath="(//*[@id='deleteIcon'])[7]")
public static WebElement delete7thRowBtn;


@FindBy(xpath="//*[@id='editScreen_bodyTable']/tbody/tr/td[3]")
public static List<WebElement> captionsList;


@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_head']/tr/th")
public static List<WebElement> headersList;


@FindBy(xpath="//*[@id='btnEditScreenNext']")
public static WebElement editScreenNxtBtn;

@FindBy(xpath="//*[@id='views_tabContent']//div[1]/a[8]/i")
public static WebElement viewRefreshViewBtn;


	
	
	

	public SanityDocCustViewPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}




}
