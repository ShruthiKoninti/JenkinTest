package com.focus.Pages;

import java.awt.RenderingHints.Key;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import javax.sound.midi.SysexMessage;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ddf.EscherColorRef.SysIndexProcedure;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Assert;
import org.testng.SkipException;

import com.focus.base.BaseEngine;
import com.focus.supporters.ExcelReader;
import com.focus.utilities.POJOUtility;
import com.sun.tools.xjc.model.SymbolSpace;

  public class SmokeBudgetPage extends BaseEngine
  {
	
	  
		// Login Page
		@FindBy(xpath = "//*[@id='txtUsername']")
		private static WebElement username;
		

			
		private static String xlfile;
		private static String resPass="Pass";
		private static String resFail="Fail";
		private static ExcelReader excelReader;
		
		private static String xlSheetName = "SmokeVouchers";
		
		private static boolean Status;

		private String expOsVoucherGrossR4;
		
		
		//User Name Display
		/*@FindBy(xpath="//*[@id='mainHeader_MainLayout']/nav/div/ul/li[7]/a/span")
		private static WebElement userNameDisplay;*/
		
		@FindBy(xpath="//span[@class='hidden-xs']")
		private static WebElement userNameDisplay;

		@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
		private static WebElement companyName;

		@FindBy(xpath="//*[@id='companyLogo']")
		private static WebElement companyLogo;

	    @FindBy(xpath="//div[@id='Layout_Popup']//input[2]")
		private static WebElement noIn;
	    
	    @FindBy(xpath="//input[@id='donotshow']")
		private static WebElement doNotShowCheckbox;
	    
	    @FindBy(xpath="//span[@class='pull-right']")
		private static WebElement closeBtnInDemoPopupScreen;
	  
	  public boolean checkSignInToCheckBudgets() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException 
	  {
		excelReader = new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		   
		/*try
		{ */ 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(username));
	    	 
			LoginPage lp                          = new LoginPage(getDriver());
	      
	        String unamelt                        = excelReader.getCellData(xlSheetName, 7, 6);
	        String pawslt                         = excelReader.getCellData(xlSheetName, 8, 6);
			
			LoginPage.enterUserName(unamelt);
			LoginPage.enterPassword(pawslt);
			
			LoginPage.clickOnSignInBtn();
			
		
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			String actUserInfo                      = userNameDisplay.getText();
			String expUserInfo                      = excelReader.getCellData(xlSheetName, 9, 7);	
			excelReader.setCellData(xlfile, xlSheetName, 9, 8, actUserInfo);
			
			System.out.println("User Name Display Value Actual                   : " + actUserInfo +             " Value Expected : " + expUserInfo);
					
				
			if (actUserInfo.equalsIgnoreCase(expUserInfo)) 
			{
				excelReader.setCellData(xlfile, xlSheetName, 7, 9, resPass);
				return true;
			} 
			else 
			{
				excelReader.setCellData(xlfile, xlSheetName, 7, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
			 String exception =e.getMessage().substring(0, 50);
	        excelReader.setExceptionInExcel(xlfile, xlSheetName, 7, 10, e.getMessage());
	        System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ exception);
	        return false;
	    }
	 }*/
	
	  
	  
	  @FindBy(xpath="//*[@id='60']/div")
	  private static WebElement  financialsMenu; 
	  
		@FindBy(xpath="//*[@id='3301']/span")
		private static WebElement  financialsBudgetMenu; 

			@FindBy(xpath="//*[@id='3302']/span")
			private static WebElement  financialsBudgetDefineBudget; 
		  
			@FindBy(xpath="//*[@id='3303']/span")
			private static WebElement  financialsBudgetConfirmBudget; 
		  
			@FindBy(xpath="//*[@id='3304']/span")
			private static WebElement  financialsBudgetReviseBudget; 
			
				@FindBy(xpath="//span[contains(text(),'Append Budget')]")
				private static WebElement appendBudgetMenu;
				
				@FindBy(xpath="//a[@id='3306']//span[contains(text(),'Add')]")
				private static WebElement addORReduceBudgetMenu;
				
				@FindBy(xpath="//span[contains(text(),'Transfer Budget')]")
				private static WebElement transferBudgetMenu;
		  
			@FindBy(xpath="//*[@id='3309']/span")
			private static WebElement  financialsBudgetApproveBudget; 
		  
			@FindBy(xpath="//*[@id='3310']/span")
			private static WebElement  financialsBudgetRejectedBudget; 

			@FindBy(xpath="//*[@id='3328']/span")
			private static WebElement  financialsBudgetReviseBudgetDetailsBudget; 

			@FindBy(xpath="//*[@id='3321']/span")
			private static WebElement  financialsBudgetReportst; 

			
			
	  public static boolean checkBudgetsMenu() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
	 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	 	
	 	/*try
	 	{*/
            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	 		financialsMenu.click();

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetMenu));
	 		financialsBudgetMenu.click();
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetDefineBudget));     	
	     	boolean actDefineBudget       = financialsBudgetDefineBudget.isDisplayed();
	     	boolean actConfirmBudget      = financialsBudgetConfirmBudget.isDisplayed();
	     	boolean actReviseBudget       = financialsBudgetReviseBudget.isDisplayed();
	     	boolean actApprovalBudget     = financialsBudgetApproveBudget.isDisplayed();
	     	boolean actRejectedBudget     = financialsBudgetRejectedBudget.isDisplayed();
	     	boolean actReviseBudgetDetail = financialsBudgetReviseBudgetDetailsBudget.isDisplayed();
	     	boolean actBudgetReport       = financialsBudgetReportst.isDisplayed();
	     	
	     	boolean expDefineBudget       = true;
	     	boolean expConfirmBudget      = true;
	     	boolean expReviseBudget       = true;
	     	boolean expApprovalBudget     = true;
	     	boolean expRejectedBudget     = true;
	     	boolean expReviseBudgetDetail = true;
	     	boolean expBudgetReport       = true;
	     	
	    	System.out.println("Budget Menu Define Budget         : " + actDefineBudget +       "  value expected  " + expDefineBudget);
	    	System.out.println("Budget Menu Confirm Budget        : " + actConfirmBudget +      "  value expected  " + expConfirmBudget);
	    	System.out.println("Budget Menu Revise Budget         : " + actReviseBudget +       "  value expected  " + expReviseBudget);
	    	System.out.println("Budget Menu Approval Budget       : " + actApprovalBudget +     "  value expected  " + expApprovalBudget);
	    	System.out.println("Budget Menu Rejected Budget       : " + actRejectedBudget +     "  value expected  " + expRejectedBudget);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actReviseBudgetDetail + "  value expected  " + expReviseBudgetDetail);
	    	System.out.println("Budget Menu Budget Report Budget  : " + actBudgetReport +       "  value expected  " + expBudgetReport);
	     	
	 		if( actDefineBudget==expDefineBudget && actConfirmBudget==expConfirmBudget && actReviseBudget==expReviseBudget
	 				&& actApprovalBudget==expApprovalBudget && actRejectedBudget==expRejectedBudget && actReviseBudgetDetail==expReviseBudgetDetail
	 				&& actBudgetReport==expBudgetReport)
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
	 			return true;
	 		}
	 		else
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
	 			return false;
	 		}
	 	}
	 	/*catch (Exception e) 
	    {
	 		String exception =e.getMessage().substring(0, 50);
	         excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	         System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	         return false;
	    }
	 }*/
	 	
	 	
	  
	  
	        @FindBy(xpath="//*[@id='optPlanName']")
	        private static WebElement  planName; 
	  
		    @FindBy(xpath="//*[@id='txtVersion']")
		    private static WebElement  versionNo; 

			@FindBy(xpath="//*[@id='dvDefineBudget']/div/div/div[2]/div[7]/div[2]/div/span")
			private static WebElement  budgetDefineFile; 
		  
			@FindBy(xpath="//*[@id='btnNavBar']/ul/li[1]/div/div[2]")
			private static WebElement  budgetExportBtn; 
		  
			@FindBy(xpath="//*[@id='btnNavBar']/ul/li[2]/div/div[2]")
			private static WebElement  budgetImportBtn; 
		  
			@FindBy(xpath="//*[@id='btnNavBar']/ul/li[4]/div/div[2]")
			private static WebElement  budgetClearBtn; 
		  
			@FindBy(xpath="//*[@id='btnNavBar']/ul/li[5]/div/div[2]")
			private static WebElement  budgetSaveBtn; 

			@FindBy(xpath="//*[@id='btnNavBar']/ul/li[7]/div/div/div")
			private static WebElement  budgetConfirmBtn; 

			@FindBy(xpath="//*[@id='btnNavBar']/ul/li[8]/div/div/div")
			private static WebElement  budgetDeleteBtn;

			@FindBy(xpath="//*[@id='btnCancel']/div/div[2]")
			private static WebElement  budgetCloseBtn; 
			
			@FindBy(xpath="//*[@id='btnCancel']/div/div[1]/span")
			private static WebElement  budgetCloseBtnFirst; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_control_heading_2']/div[1]")
			private static WebElement  budgetTypeColHeadingTxt; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_control_heading_3']/div[1]")
			private static WebElement  budgetDepartmenteColHeadingTxt; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_control_heading_6']/div[1]")
			private static WebElement  budgetItemColHeadingTxt; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_control_heading_7']/div[1]")
			private static WebElement  budgetAccountColHeadeingTxt; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_control_heading_8']/div[1]")
			private static WebElement  budgetQuanityColHeadingTxt; 
		
			@FindBy(xpath="//*[@id='grdDefineBudget_control_heading_9']/div[1]")
			private static WebElement  budgetValueColHeadingTxt; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_control_heading_13']/div[1]")
			private static WebElement  budgetStatusColHeadingTxt; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_col_1-2']")
			private static WebElement  budgetTypeR1C2SelectCell; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_cmbbudgettype']")
			private static WebElement  budgetTypeR1C1InputCell; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_col_1-3']")
			private static WebElement  budgetDepartmentR1C3SelectCell; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_option_Tag1']")
			private static WebElement  budgetDepartmentR1C3InputCell; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_col_1-6']")
			private static WebElement  budgetItemR1C4SelectCell; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_option_product']")
			private static WebElement  budgetItemR1C4InputCell; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_col_1-7']")
			private static WebElement  budgetAccountR1C5SelectCell; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_option_account']")
			private static WebElement  budgetAccountR1C5InputCell; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_col_1-8']")
			private static WebElement  budgetQtyR1C6SelectCell; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_txtquantity']")
			private static WebElement  budgetQtyR1C57InputCell; 
				
			@FindBy(xpath="//*[@id='grdDefineBudget_col_1-9']")
			private static WebElement  budgetValueR1C3SelectCell; 
			
			@FindBy(xpath="//*[@id='grdDefineBudget_txtbudgetvalue']")
			private static WebElement  budgetValueR1C9InputCell; 
	  
			@FindBy(xpath="//*[@id='grdDefineBudget_col_1-13']")
			private static WebElement  budgetStatusR1C4SelectCell; 
			
			
			
			   
			
			
			
	  public static boolean checkDefineBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	  {
	 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	 	
	 	/*try
	 	{*/
            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetDefineBudget));
            financialsBudgetDefineBudget.click();
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));     	
	     	
	 		boolean actPlanName              = planName.isDisplayed();
	     	boolean actVersionName           = versionNo.isDisplayed();
	     	boolean actBudgetFileBtn         = budgetDefineFile.isDisplayed();
	     	boolean actExportBtn             = budgetExportBtn.isDisplayed();
	     	boolean actImportBtn             = budgetImportBtn.isDisplayed();
	     	boolean actClearBtn              = budgetClearBtn.isDisplayed();
	     	boolean actSaveBtn               = budgetSaveBtn.isDisplayed();
	     	boolean actConfirmBtn            = budgetConfirmBtn.isDisplayed();
	     	boolean actCloseBtn              = budgetCloseBtn.isDisplayed();
	     	boolean actTypeColHeading        = budgetTypeColHeadingTxt.isDisplayed();
	     	boolean actDepartmentColHeading  = budgetDepartmenteColHeadingTxt.isDisplayed();
	     	boolean actItemHeading           = budgetItemColHeadingTxt.isDisplayed();
	     	boolean actAccountHeading        = budgetAccountColHeadeingTxt.isDisplayed();
	     	boolean actQtyHeading            = budgetQuanityColHeadingTxt.isDisplayed();
	     	boolean actValHeading            = budgetValueColHeadingTxt.isDisplayed();
	     	boolean actStatusHeading         = budgetStatusColHeadingTxt.isDisplayed();
	     	
	     	
	     	boolean expPlanName              = true;
	     	boolean expVersionName           = true;
	     	boolean expBudgetFileBtn         = true;
	     	boolean expExportBtn             = true;
	     	boolean expImportBtn             = true;
	     	boolean expClearBtn              = true;
	     	boolean expSaveBtn               = true;
	     	boolean expConfirmBtn            = true;
	     	boolean expCloseBtn              = true;
	     	boolean expTypeColHeading        = true;
	     	boolean expDepartmentColHeading  = true;
	     	boolean expItemHeading           = true;
	     	boolean expAccountHeading        = true;
	     	boolean expQtyHeading            = true;
	     	boolean expValHeading            = true;
	     	boolean expStatusHeading         = true;
	     	
	    	System.out.println("Budget Menu Define Budget         : " + actPlanName +             "  value expected  " + expPlanName);
	    	System.out.println("Budget Menu Confirm Budget        : " + actVersionName +          "  value expected  " + expVersionName);
	    	System.out.println("Budget Menu Budget Budget         : " + actBudgetFileBtn +        "  value expected  " + expBudgetFileBtn);
	    	System.out.println("Budget Menu Approval Budget       : " + actExportBtn +            "  value expected  " + expExportBtn);
	    	System.out.println("Budget Menu Rejected Budget       : " + actImportBtn +            "  value expected  " + expImportBtn);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actClearBtn +             "  value expected  " + expClearBtn);
	    	System.out.println("Budget Menu Budget Report Budget  : " + actSaveBtn +              "  value expected  " + expSaveBtn);
	    	System.out.println("Budget Menu Define Budget         : " + actConfirmBtn +           "  value expected  " + expConfirmBtn);
	    	System.out.println("Budget Menu Confirm Budget        : " + actCloseBtn +             "  value expected  " + expCloseBtn);
	    	System.out.println("Budget Menu Revise Budget         : " + actTypeColHeading +       "  value expected  " + expTypeColHeading);
	    	System.out.println("Budget Menu Rejected Budget       : " + actDepartmentColHeading + "  value expected  " +expDepartmentColHeading);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actItemHeading +          "  value expected  " + expItemHeading);
	    	System.out.println("Budget Menu Budget Report Budget  : " + actAccountHeading +       "  value expected  " + expAccountHeading);
	    	System.out.println("Budget Menu Define Budget         : " + actQtyHeading +           "  value expected  " + expQtyHeading);
	    	System.out.println("Budget Menu Confirm Budget        : " + actValHeading +           "  value expected  " + expValHeading);
	    	System.out.println("Budget Menu Revise Budget         : " + actStatusHeading +        "  value expected  " + expStatusHeading);
	    	
	    	
	 		if( actPlanName==expPlanName && actVersionName==expVersionName && actBudgetFileBtn==expBudgetFileBtn
	 				&& actImportBtn==expImportBtn && actClearBtn==expClearBtn && actSaveBtn==expSaveBtn
	 				&& actConfirmBtn==expConfirmBtn && actConfirmBtn==expConfirmBtn  && actCloseBtn==expCloseBtn && actTypeColHeading==expTypeColHeading
	 				&& actItemHeading==expItemHeading && actDepartmentColHeading==expDepartmentColHeading && actItemHeading==expItemHeading
	 				&& actAccountHeading==expAccountHeading && actQtyHeading==expQtyHeading && actValHeading==expValHeading && actStatusHeading==expStatusHeading)
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
	 			return true;
	 		}
	 		else
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
	 			return false;
	 		}
	 	}
	 	/*catch (Exception e) 
	     {
	 		 String exception =e.getMessage().substring(0, 50);
	         excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	         System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	         return false;
	     }
	  }*/
  
	  
	
	 @FindBy(xpath="//div[@class='theme_color font-6']")
	 public static WebElement errorMessage;
		 
	 @FindBy(xpath="//span[@class='icon-reject2 theme_color']")
	 public static WebElement errorMessageCloseBtn;
		
	 @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[1]/td[3]")
	 public static WebElement budgetTypeGetText;
	 
	 @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[1]/td[4]")
	 public static WebElement budgetDepartmentGetText;
	 	
	 @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[1]/td[7]")
	 public static WebElement budgetItemGetText;
	 
	 @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[1]/td[8]")
	 public static WebElement budgetAccountGetText;
	 
	 @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[1]/td[9]")
	 public static WebElement budgetQtyGetText;	
	 
	 @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[1]/td[10]")
	 public static WebElement budgetValueGetText;
	
	 @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[1]/td[14]")
	 public static WebElement budgetStatusGetText;
	
	 
	 public static boolean checkSaveConfirmPlanBudgetDefineBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
	 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	 	
	 	/*try
	 	{*/ 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));     	
	 		planName.sendKeys("ConfirmPlanBudget");
	 		Thread.sleep(2000);
	 		versionNo.click();
	 		versionNo.sendKeys(Keys.TAB);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetTypeR1C2SelectCell));     	
	 		budgetTypeR1C2SelectCell.click();
	 		
	 		budgetTypeR1C1InputCell.click();
	 		
	 		budgetTypeR1C1InputCell.sendKeys("By Product By Qty");  	
	 		Thread.sleep(2000);
	 		budgetTypeR1C1InputCell.sendKeys(Keys.TAB);
	    	
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDepartmentR1C3SelectCell));     	
	 		budgetDepartmentR1C3SelectCell.click();	 	
	 		
	 		budgetDepartmentR1C3InputCell.click();
	 
	 		budgetDepartmentR1C3InputCell.sendKeys("DUBAI");
	 		Thread.sleep(2000);
	 		budgetDepartmentR1C3InputCell.sendKeys(Keys.TAB);
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetItemR1C4SelectCell));     	
	 		budgetItemR1C4SelectCell.click();	 	
	 		
	 		budgetItemR1C4InputCell.click();
	 		budgetItemR1C4InputCell.sendKeys("BR COGS ITEM");
	 		Thread.sleep(2000);
	 		budgetItemR1C4InputCell.sendKeys(Keys.TAB);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAccountR1C5SelectCell));     	
	 		budgetAccountR1C5SelectCell.click();	 	
	 		/*Thread.sleep(2000);
	 		budgetAccountR1C5InputCell.click();
	 		budgetAccountR1C5InputCell.sendKeys("BR COGS ITEM");*/
	 		Thread.sleep(2000);
	 		budgetAccountR1C5InputCell.sendKeys(Keys.TAB);
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetQtyR1C6SelectCell));     	
	 		budgetQtyR1C6SelectCell.click();	 	
	 		
	 		budgetQtyR1C57InputCell.click();
	 		budgetQtyR1C57InputCell.sendKeys("12");
	 		Thread.sleep(2000);
	 		budgetQtyR1C57InputCell.sendKeys(Keys.TAB);
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetValueR1C3SelectCell));     	
	 		budgetValueR1C3SelectCell.click();	 	
	 		Thread.sleep(2000);
	 		budgetValueR1C9InputCell.click();
	 		budgetValueR1C9InputCell.sendKeys("120");
	 		budgetValueR1C9InputCell.sendKeys(Keys.TAB);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetSaveBtn));     	
	 		budgetSaveBtn.click();	
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
	 		String actErrorMessage=errorMessage.getText();
	 		String expErrorMessage="Budget Saved succesfully";

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
	 		errorMessageCloseBtn.click();
	 		
	 		
	 		System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
	 		
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));   
	 		planName.sendKeys(Keys.HOME);
	 		planName.sendKeys(Keys.SHIFT,Keys.END);
	 		planName.sendKeys("ConfirmPlanBudget");
	 		Thread.sleep(2000);
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(versionNo));
	 		versionNo.click();
	 		
	 		System.out.println(versionNo.getAttribute("value"));	 	
	 	

	 		
	 		String actVersionNo      = versionNo.getAttribute("value");
	 		String actBudgetType     = budgetTypeGetText.getText();
	 		String actDepartment     = budgetDepartmentGetText.getText();
	 		String actAccount        = budgetAccountGetText.getText();
	 		String actItem           = budgetItemGetText.getText();
	 		String actQy             = budgetQtyGetText.getText();
	 		String actValue          = budgetValueGetText.getText();
	 		String actBudgetStatus   = budgetStatusGetText.getText();
	 		
	 		String expVersionNo      = "1";
	 		String expBudgetType     =  "By Product By Qty By Value";
	 		String expDepartment     = "DUBAI";
	 		String expAccount        = "BR COGS ACC INV";
	 		String expItem           = "BR COGS ITEM";
	 		String expQty            = "12.00";
	 		String expValue          = "120.00";
	 		String expBudgetStatus   = "Pending";
	 		
	 		
	 		System.out.println("Budget Menu Define Budget         : " + actVersionNo +        "  value expected  " + expVersionNo);
	    	System.out.println("Budget Menu Confirm Budget        : " + actBudgetType +       "  value expected  " + expBudgetType);
	    	System.out.println("Budget Menu Revise Budget         : " + actDepartment +       "  value expected  " + expDepartment);
	    	System.out.println("Budget Menu Approval Budget       : " + actAccount +          "  value expected  " + expAccount);
	    	System.out.println("Budget Menu Rejected Budget       : " + actItem +             "  value expected  " + expItem);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actQy +               "  value expected  " + expQty);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actValue +            "  value expected  " + expValue);
	    	System.out.println("Budget Menu Budget Report Budget  : " + actBudgetStatus +     "  value expected  " + actBudgetStatus);
	     
	    	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetConfirmBtn));     	
	    	budgetConfirmBtn.click();
	    	
	    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
	 		String actErrorMessages=errorMessage.getText();
	 		String expErrorMessages="Budget Confirmed successfully";

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
	 		errorMessageCloseBtn.click();
	 		
	 		
	 		System.out.println("ValidationMessage  :  "+actErrorMessages +" Value Expected : "+expErrorMessages);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));     	
	 		planName.sendKeys("ConfirmPlanBudget");
	 		
	 		Thread.sleep(2000);
	 		
	 		versionNo.click();
	 		
	 		System.out.println(versionNo.getAttribute("value"));	 	

	 		
	 		String actVersionNos      = versionNo.getAttribute("value");
	 		String actBudgetTypes     = budgetTypeGetText.getText();
	 		String actDepartments     = budgetDepartmentGetText.getText();
	 		String actAccounts        = budgetAccountGetText.getText();
	 		String actItems           = budgetItemGetText.getText();
	 		String actQys             = budgetQtyGetText.getText();
	 		String actValues          = budgetValueGetText.getText();
	 		String actBudgetStatuss   = budgetStatusGetText.getText();
	 		
	 		String expVersionNos      = "1";
	 		String expBudgetTypes     =  "By Product By Qty By Value";
	 		String expDepartments     = "DUBAI";
	 		String expAccounts        = "BR COGS ACC INV";
	 		String expItems           = "BR COGS ITEM";
	 		String expQtys            = "12.00";
	 		String expValues          = "120.00";
	 		String expBudgetStatuss   = "Authorized";
	 		
	 		
	 		System.out.println("Budget Menu Define Budget         : " + actVersionNos +        "  value expected  " + expVersionNos);
	    	System.out.println("Budget Menu Confirm Budget        : " + actBudgetTypes +       "  value expected  " + expBudgetTypes);
	    	System.out.println("Budget Menu Revise Budget         : " + actDepartments +       "  value expected  " + expDepartments);
	    	System.out.println("Budget Menu Approval Budget       : " + actAccounts +          "  value expected  " + expAccounts);
	    	System.out.println("Budget Menu Rejected Budget       : " + actItems +             "  value expected  " + expItems);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actQys +               "  value expected  " + expQtys);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actValues +            "  value expected  " + expValues);
	    	System.out.println("Budget Menu Budget Report Budget  : " + actBudgetStatuss +     "  value expected  " + expBudgetStatuss);
	    	
	    	
	    	Thread.sleep(5000);
	    	
	    	budgetCloseBtnFirst.click();
	    	
	 		
	 		if( /*actErrorMessage.equalsIgnoreCase(expErrorMessage) &&*/  actVersionNo.equalsIgnoreCase(expVersionNo) && actBudgetType.equalsIgnoreCase(expBudgetType)
	 				&& actDepartment.equalsIgnoreCase(expDepartment) && actAccount.equalsIgnoreCase(expAccount)
	 				&& actAccount.equalsIgnoreCase(expAccount) && actItem.equalsIgnoreCase(expItem)
	 				&& actQy.equalsIgnoreCase(expQty) && actValue.equalsIgnoreCase(expValue) && actBudgetStatus.equalsIgnoreCase(expBudgetStatus)
	 				
	 			    &&  actErrorMessages.equalsIgnoreCase(expErrorMessages) &&  actVersionNos.equalsIgnoreCase(expVersionNos) && actBudgetTypes.equalsIgnoreCase(expBudgetTypes)
	 				&& actDepartments.equalsIgnoreCase(expDepartments) && actAccounts.equalsIgnoreCase(expAccounts)
	 				&& actAccounts.equalsIgnoreCase(expAccounts) && actItems.equalsIgnoreCase(expItems)
	 				&& actQys.equalsIgnoreCase(expQtys) && actValues.equalsIgnoreCase(expValues) && actBudgetStatuss.equalsIgnoreCase(expBudgetStatuss))
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
	 			return true;
	 		}
	 		else
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
	 			return false;
	 		}
	 	}
	 	/*catch (Exception e) 
	     {
	 		 String exception =e.getMessage().substring(0, 50);
	         excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	         System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	         return false;
	     }
	  }*/


	 
	 public static boolean checkSaveEditPlanBudgetDefineBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
	 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	 	
	 	/*try
	 	{*/ 	
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	 		financialsMenu.click();
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetMenu));
	 		financialsBudgetMenu.click();

            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetDefineBudget));
            financialsBudgetDefineBudget.click();
            
            Thread.sleep(3000);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName)); 
	 		planName.sendKeys(Keys.HOME);
	 		planName.sendKeys(Keys.SHIFT,Keys.END);
	 		Thread.sleep(2000);
	 		planName.sendKeys("NonConfirmPlanBudget");
	 		Thread.sleep(2000);
	 		versionNo.click();
	 		versionNo.sendKeys(Keys.TAB);
	 		Thread.sleep(2000);
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetTypeR1C2SelectCell));     	
	 		budgetTypeR1C2SelectCell.click();
	 		
	 		budgetTypeR1C1InputCell.click();
	 		
	 		budgetTypeR1C1InputCell.sendKeys("By Product By Qty By Value");  	
	 		Thread.sleep(2000);
	 		budgetTypeR1C1InputCell.sendKeys(Keys.TAB);
	    	
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDepartmentR1C3SelectCell));     	
	 		budgetDepartmentR1C3SelectCell.click();	 	
	 		
	 		budgetDepartmentR1C3InputCell.click();
	 
	 		budgetDepartmentR1C3InputCell.sendKeys("DUBAI");
	 		Thread.sleep(2000);
	 		budgetDepartmentR1C3InputCell.sendKeys(Keys.TAB);
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetItemR1C4SelectCell));     	
	 		budgetItemR1C4SelectCell.click();	 	
	 		
	 		budgetItemR1C4InputCell.click();
	 		budgetItemR1C4InputCell.sendKeys("FIFO COGS ITEM");
	 		Thread.sleep(2000);
	 		budgetItemR1C4InputCell.sendKeys(Keys.TAB);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAccountR1C5SelectCell));     	
	 		budgetAccountR1C5SelectCell.click();	 	
	 		Thread.sleep(2000);
	 		budgetAccountR1C5InputCell.click();
	 		//budgetAccountR1C5InputCell.sendKeys("BR COGS ITEM");
	 		Thread.sleep(2000);
	 		budgetAccountR1C5InputCell.sendKeys(Keys.TAB);
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetQtyR1C6SelectCell));     	
	 		budgetQtyR1C6SelectCell.click();	 	
	 		
	 		budgetQtyR1C57InputCell.click();
	 		budgetQtyR1C57InputCell.sendKeys("5");
	 		Thread.sleep(2000);
	 		budgetQtyR1C57InputCell.sendKeys(Keys.TAB);
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetValueR1C3SelectCell));     	
	 		budgetValueR1C3SelectCell.click();	 	
	 		Thread.sleep(2000);
	 		budgetValueR1C9InputCell.click();
	 		budgetValueR1C9InputCell.sendKeys("4");
	 		budgetValueR1C9InputCell.sendKeys(Keys.TAB);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetSaveBtn));     	
	 		budgetSaveBtn.click();	
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
	 		String actErrorMessage=errorMessage.getText();
	 		String expErrorMessage="Budget Saved succesfully";

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
	 		errorMessageCloseBtn.click();
	 		
            System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
	 		
            
            Thread.sleep(2000);
            
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));     	
	 		planName.click();
	 		planName.sendKeys(Keys.END);
	 		planName.sendKeys(Keys.SHIFT,Keys.HOME);
	 		planName.sendKeys("NonConfirmPlanBudget");
	 		Thread.sleep(2000);
	 		
	 		versionNo.click();
	 		
	 		System.out.println(versionNo.getAttribute("value"));	 	

	 		
	 		String actVersionNos      = versionNo.getAttribute("value");
	 		String actBudgetTypes     = budgetTypeGetText.getText();
	 		String actDepartments     = budgetDepartmentGetText.getText();
	 		String actAccounts        = budgetAccountGetText.getText();
	 		String actItems           = budgetItemGetText.getText();
	 		String actQys             = budgetQtyGetText.getText();
	 		String actValues          = budgetValueGetText.getText();
	 		String actBudgetStatuss   = budgetStatusGetText.getText();
	 		
	 		String expVersionNos      = "1";
	 		String expBudgetTypes     =  "By Product By Qty By Value";
	 		String expDepartments     = "DUBAI";
	 		String expAccounts        = "FIFO COGS ACC INV";
	 		String expItems           = "FIFO COGS ITEM";
	 		String expQtys            = "12.00";
	 		String expValues          = "4.00";
	 		String expBudgetStatuss   = "Pending";
	 		
	 		
	 		System.out.println("Budget Menu Define Budget         : " + actVersionNos +        "  value expected  " + expVersionNos);
	    	System.out.println("Budget Menu Confirm Budget        : " + actBudgetTypes +       "  value expected  " + expBudgetTypes);
	    	System.out.println("Budget Menu Revise Budget         : " + actDepartments +       "  value expected  " + expDepartments);
	    	System.out.println("Budget Menu Approval Budget       : " + actAccounts +          "  value expected  " + expAccounts);
	    	System.out.println("Budget Menu Rejected Budget       : " + actItems +             "  value expected  " + expItems);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actQys +               "  value expected  " + expQtys);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actValues +            "  value expected  " + expValues);
	    	System.out.println("Budget Menu Budget Report Budget  : " + actBudgetStatuss +     "  value expected  " + expBudgetStatuss);
	 		
	 		
	 		Thread.sleep(2000);
	    	budgetValueGetText.click();
	 		budgetValueR1C9InputCell.click();
	 		budgetValueR1C9InputCell.sendKeys("10");
	 		budgetValueR1C9InputCell.sendKeys(Keys.TAB);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetSaveBtn));     	
	 		budgetSaveBtn.click();	
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
	 		String actErrorMessages=errorMessage.getText();
	 		String expErrorMessages="Budget Saved succesfully";

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
	 		errorMessageCloseBtn.click();
	 		


            System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));     	
	 		planName.sendKeys("NonConfirmPlanBudget");
	 		
	 		Thread.sleep(2000);
	 		
	 		versionNo.click();
	 		
	 		System.out.println(versionNo.getAttribute("value"));	 	

	 		
	 		String actVersionNo      = versionNo.getAttribute("value");
	 		String actBudgetType     = budgetTypeGetText.getText();
	 		String actDepartment     = budgetDepartmentGetText.getText();
	 		String actAccount        = budgetAccountGetText.getText();
	 		String actItem           = budgetItemGetText.getText();
	 		//String actQy             = budgetQtyGetText.getText();
	 		String actValus          = budgetValueGetText.getText();
	 		String actBudgetStatus   = budgetStatusGetText.getText();
	 		
	 		String expVersionNo      = "1";
	 		String expBudgetTypess     =  "By Product By Value";
	 		String expDepartmentsss     = "DUBAI";
	 		String expAccount       = "FIFo COGS ACC INV";
	 		String expItemm          = "FIFO COGS ITEM";
	 		//String expQtys            = "12.00";
	 		String expValuess          = "104";
	 		String expBudgetStatusss  = "Pending";
	 		
	 		
	 		System.out.println("Budget Menu Define Budget         : " + actVersionNos +        "  value expected  " + expVersionNos);
	    	System.out.println("Budget Menu Confirm Budget        : " + actBudgetTypes +       "  value expected  " + expBudgetTypes);
	    	System.out.println("Budget Menu Revise Budget         : " + actDepartments +       "  value expected  " + expDepartments);
	    	System.out.println("Budget Menu Approval Budget       : " + actAccounts +          "  value expected  " + expAccounts);
	    	System.out.println("Budget Menu Rejected Budget       : " + actItems +             "  value expected  " + expItems);
	    	//System.out.println("Budget Menu Revise Budget Budget  : " + actQys +               "  value expected  " + expQtys);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actValues +            "  value expected  " + expValues);
	    	System.out.println("Budget Menu Budget Report Budget  : " + actBudgetStatuss +     "  value expected  " + expBudgetStatuss);
	 		System.out.println("ValidationMessage                 :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
	 		
	 		
	 		Thread.sleep(2000);
	 		budgetCloseBtnFirst.click();
	 		
	 		
	 		if( actErrorMessage.equalsIgnoreCase(expErrorMessage) 
	 				
	 			    &&  actVersionNos.equalsIgnoreCase(expVersionNos) && actBudgetTypes.equalsIgnoreCase(expBudgetTypes)
	 				&& actDepartments.equalsIgnoreCase(expDepartments) && actAccounts.equalsIgnoreCase(expAccounts)
	 				&& actAccounts.equalsIgnoreCase(expAccounts) && actItems.equalsIgnoreCase(expItems)
	 				&& actValues.equalsIgnoreCase(expValues) && actBudgetStatuss.equalsIgnoreCase(expBudgetStatuss)


	 				
	 			    &&  actVersionNos.equalsIgnoreCase(expVersionNos) && actBudgetTypes.equalsIgnoreCase(expBudgetTypes)
	 				&& actDepartments.equalsIgnoreCase(expDepartments) && actAccounts.equalsIgnoreCase(expAccounts)
	 				&& actAccounts.equalsIgnoreCase(expAccounts) && actItems.equalsIgnoreCase(expItems)
	 				&& actValues.equalsIgnoreCase(expValues) && actBudgetStatuss.equalsIgnoreCase(expBudgetStatuss)
	 				)
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
	 			return true;
	 		}
	 		else
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
	 			return false;
	 		}
	 	}
	 	/*catch (Exception e) 
	     {
	 		 String exception =e.getMessage().substring(0, 50);
	         excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	         System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	         return false;
	     }
	 		
	 }*/
	  
	 
	 
	 public static boolean checkSaveToDeletePlanBudgetDefineBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
	 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	 	
	 	/*try
	 	{ */		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	 		financialsMenu.click();
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetMenu));
	 		financialsBudgetMenu.click();

            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetDefineBudget));
            financialsBudgetDefineBudget.click();
            
	 		Thread.sleep(3000);
            
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));     	
	 		planName.sendKeys("ToDeleteNonConfirmPlanBudget");
	 		Thread.sleep(2000);
	 		versionNo.click();
	 		versionNo.sendKeys(Keys.TAB);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetTypeR1C2SelectCell));     	
	 		budgetTypeR1C2SelectCell.click();
	 		
	 		budgetTypeR1C1InputCell.click();
	 		
	 		budgetTypeR1C1InputCell.sendKeys("By Account By Value");  	
	 		Thread.sleep(2000);
	 		budgetTypeR1C1InputCell.sendKeys(Keys.TAB);
	    	
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDepartmentR1C3SelectCell));     	
	 		budgetDepartmentR1C3SelectCell.click();	 	
	 		
	 		budgetDepartmentR1C3InputCell.click();
	 
	 		budgetDepartmentR1C3InputCell.sendKeys("DUBAI");
	 		Thread.sleep(2000);
	 		budgetDepartmentR1C3InputCell.sendKeys(Keys.TAB);
	 		
	 		
	 		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetItemR1C4SelectCell));     	
	 		budgetItemR1C4SelectCell.click();	 	
	 		
	 		budgetItemR1C4InputCell.click();
	 		budgetItemR1C4InputCell.sendKeys("FIFO COGS ITEM");
	 		Thread.sleep(2000);
	 		budgetItemR1C4InputCell.sendKeys(Keys.TAB);*/
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAccountR1C5SelectCell));     	
	 		budgetAccountR1C5SelectCell.click();	 	
	 		Thread.sleep(2000);
	 		budgetAccountR1C5InputCell.click();
	 		budgetAccountR1C5InputCell.sendKeys("STD RATE COGS ACC INV");
	 		Thread.sleep(2000);
	 		budgetAccountR1C5InputCell.sendKeys(Keys.TAB);
	 		
	 		
	 	/*	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetQtyR1C6SelectCell));     	
	 		budgetQtyR1C6SelectCell.click();	 	
	 		*/
	 		/*budgetQtyR1C57InputCell.click();
	 		budgetQtyR1C57InputCell.sendKeys("5");
	 		Thread.sleep(2000);
	 		budgetQtyR1C57InputCell.sendKeys(Keys.TAB);*/
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetValueR1C3SelectCell));     	
	 		budgetValueR1C3SelectCell.click();	 	
	 		Thread.sleep(2000);
	 		budgetValueR1C9InputCell.click();
	 		budgetValueR1C9InputCell.sendKeys("4");
	 		budgetValueR1C9InputCell.sendKeys(Keys.TAB);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetSaveBtn));     	
	 		budgetSaveBtn.click();	
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
	 		String actErrorMessage=errorMessage.getText();
	 		String expErrorMessage="Budget Saved succesfully";

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
	 		errorMessageCloseBtn.click();
	 		


            System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));     	
	 		planName.sendKeys("ToDeleteNonConfirmPlanBudget");
	 		
	 		Thread.sleep(2000);
	 		
	 		versionNo.click();
	 		
	 		System.out.println(versionNo.getAttribute("value"));	 	
	 		
	 		String actVersionNos      = versionNo.getAttribute("value");
	 		String actBudgetTypes     = budgetTypeGetText.getText();
	 		String actDepartments     = budgetDepartmentGetText.getText();
	 		String actAccounts        = budgetAccountGetText.getText();
	 		//String actItems           = budgetItemGetText.getText();
	 		//String actQy             = budgetQtyGetText.getText();
	 		String actValues          = budgetValueGetText.getText();
	 		String actBudgetStatuss   = budgetStatusGetText.getText();

	 		
	 		
	 		String expVersionNos      = "1";
	 		String expBudgetTypes     =  "By Account By Value";
	 		String expDepartments     = "DUBAI";
	 		String expAccounts        = "STD RATE COGS ACC INV";
	 		//String expItems           = "FIFO COGS ITEM";
	 		//String expQtys            = "12.00";
	 		String expValues          = "4.00";
	 		String expBudgetStatuss   = "Pending";
	 		
	 		
	 		System.out.println("Budget Menu Define Budget         : " + actVersionNos +        "  value expected  " + expVersionNos);
	    	System.out.println("Budget Menu Confirm Budget        : " + actBudgetTypes +       "  value expected  " + expBudgetTypes);
	    	System.out.println("Budget Menu Revise Budget         : " + actDepartments +       "  value expected  " + expDepartments);
	    	System.out.println("Budget Menu Approval Budget       : " + actAccounts +          "  value expected  " + expAccounts);
	    	//System.out.println("Budget Menu Rejected Budget       : " + actItems +             "  value expected  " + expItems);
	    	//System.out.println("Budget Menu Revise Budget Budget  : " + actQys +               "  value expected  " + expQtys);
	    	System.out.println("Budget Menu Revise Budget Budget  : " + actValues +            "  value expected  " + expValues);
	    	System.out.println("Budget Menu Budget Report Budget  : " + actBudgetStatuss +     "  value expected  " + expBudgetStatuss);
	 		
	 		
	 		
	 		System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
	 		
	 		
	 		
	 		/*budgetValueR1C9InputCell.click();
	 		budgetValueR1C9InputCell.sendKeys("5");
	 		budgetValueR1C9InputCell.sendKeys(Keys.TAB);*/
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDeleteBtn));     	
	 		budgetDeleteBtn.click();	
	 			 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
	 		String actErrorMessages=errorMessage.getText();
	 		String expErrorMessages="Budget Saved succesfully";

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
	 		errorMessageCloseBtn.click();
	 	
            System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));     	
	 		planName.sendKeys("NonConfirmPlanBudget");
	 		
	 		Thread.sleep(2000);
	 		
	 		versionNo.click();
	 		
	 		System.out.println(versionNo.getAttribute("value"));	 	
	 		
	 		boolean actBudgetType     = budgetTypeGetText.getText().isEmpty();
	 		boolean actDepartment     = budgetDepartmentGetText.getText().isEmpty();
	 		boolean actAccount        = budgetAccountGetText.getText().isEmpty();
	 		boolean actItem           = budgetItemGetText.getText().isEmpty();
	 
	 		boolean actValue          = budgetValueGetText.getText().isEmpty();
	 		boolean actBudgetStatus   = budgetStatusGetText.getText().isEmpty();
	 		
	
	 		boolean expBudgetType     = true;
	 		boolean expDepartment     = true;
	 		boolean expAccount        = true;
	 		boolean expItem           = true;

	 		boolean expValue        = true;
	 		boolean expBudgetStatus = true;
	 		
	 		
	 		
	    	System.out.println("Budget Menu Confirm Budget        : " + actBudgetType +       "  value expected  " + expBudgetType);
	    	System.out.println("Budget Menu Revise Budget         : " + actDepartment +       "  value expected  " + expDepartment);
	    	System.out.println("Budget Menu Approval Budget       : " + actAccount +          "  value expected  " + expAccount);
	    	System.out.println("Budget Menu Rejected Budget       : " + actItem +             "  value expected  " + expItem);

	    	System.out.println("Budget Menu Revise Budget Budget  : " + actValue +            "  value expected  " + expValue);
	    	System.out.println("Budget Menu Budget Report Budget  : " + actBudgetStatus +     "  value expected  " + expBudgetStatus);
	 		
	 		
	 		
	 		System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
	 		if( actErrorMessage.equalsIgnoreCase(expErrorMessage) 
	 				
	 			    &&  actVersionNos.equalsIgnoreCase(expVersionNos) && actBudgetTypes.equalsIgnoreCase(expBudgetTypes)
	 				&& actDepartments.equalsIgnoreCase(expDepartments) && actAccounts.equalsIgnoreCase(expAccounts)
	 				&& actAccounts.equalsIgnoreCase(expAccounts) 
	 				&& actValues.equalsIgnoreCase(expValues) && actBudgetStatuss.equalsIgnoreCase(expBudgetStatuss)
	 				&& actBudgetStatus==expBudgetStatus && actValue==expValue && actItem==expItem && actAccount==expAccount
	 				&& actDepartment==expDepartment && actBudgetType==expBudgetType)
	 		
	 		/*if( actErrorMessage.equalsIgnoreCase(expErrorMessage) 
	 				
	 			    &&  actVersionNos.equalsIgnoreCase(expVersionNos) && actBudgetTypes.equalsIgnoreCase(expBudgetTypes)
	 				&& actDepartments.equalsIgnoreCase(expDepartments) && actAccounts.equalsIgnoreCase(expAccounts)
	 				&& actAccounts.equalsIgnoreCase(expAccounts) && actItems.equalsIgnoreCase(expItems)
	 				&& actValues.equalsIgnoreCase(expValues) && actBudgetStatuss.equalsIgnoreCase(expBudgetStatuss)
	 				)*/
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
	 			return true;
	 		}
	 		else
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
	 			return false;
	 		}
	 	}
	 	/*catch (Exception e) 
	     {
	 		 String exception =e.getMessage().substring(0, 50);
	         excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	         System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	         return false;
	     }		
	 }*/
	 
	 

	 @FindBy(xpath="//*[@id='dashName']")
	 private static WebElement dashboard;
	 
	 public boolean checkCloseButtonInDefineBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	 {
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetCloseBtn));
			budgetCloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
			
			boolean actdashboard=dashboard.isDisplayed();
			boolean expdashboard=true;
				
			System.out.println("dashboard  :"+actdashboard +"  Value Expected  : "+expdashboard);
			
			if(actdashboard==expdashboard)	
			{
				System.out.println("Test Pass : Cancel Option Working ");
				excelReader.setCellData(xlfile, xlSheetName, 53, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Cancel Option Not Working");
				excelReader.setCellData(xlfile, xlSheetName, 53, 9, resFail);
				return false;
			}
		}

	  
	 
	 
	 
	  // CONFIRM BUDGET
	 
	 
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[1]/div[1]/div[2]/table/tbody/tr[2]/td[1]/input[2]")
      private static WebElement  planNameInConfirmBudget; 
	  	  
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[1]/div[2]/div[2]")
      private static WebElement  versionTxtBox; 
	  
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[1]/nav/div/div[2]/ul/li[4]/div/div[1]/span")
	  private static WebElement  exportBtn; 

      @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[1]/nav/div/div[2]/ul/li[5]/div/div[2]")
      private static WebElement  clearBtn; 
	  
      @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[1]/nav/div/div[2]/ul/li[6]/div/div[2]")
	  private static WebElement  closeBtn; 
	  
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[4]/fieldset/div[1]/div[2]/table/tbody/tr[2]/td[1]/input[2]")
	  private static WebElement  accountCmbx; 
	  
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[4]/fieldset/div[2]/div[2]/table/tbody/tr[2]/td[1]/input[2]")
	  private static WebElement  itemCmbx; 
	  
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[4]/fieldset/div[3]/div[2]/table/tbody/tr[2]/td[1]/input[2]")
	  private static WebElement departmentCmbx; 

	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[6]/div/table/tbody/tr/td[4]")
	  private static WebElement budgetTxtBox; 

	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[6]/div/table/tbody/tr/td[5]")
	  private static WebElement departmentTxtBox;

	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[6]/div/table/tbody/tr/td[8]")
	  private static WebElement itemTxtBox; 
		
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[6]/div/table/tbody/tr/td[9]")
	  private static WebElement accountTxtBox; 
	
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[6]/div/table/tbody/tr/td[10]")
	  private static WebElement qtyTxtBox; 
		
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[6]/div/table/tbody/tr/td[11]")
	  private static WebElement  valueTxtBox; 
		
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[7]/div[2]")
      private static WebElement  budgetTotalValue; 
	  
	  @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[8]/div[2]")
      private static WebElement  budgetTotalQty; 
		
		   
		
		
		
	public static boolean checkConfirmBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	 excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
	 /*try
	 {*/
   	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	   financialsMenu.click();

	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetMenu));
	   financialsBudgetMenu.click();
		 
	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetConfirmBudget));
	   financialsBudgetConfirmBudget.click();
      
       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameInConfirmBudget));
	    boolean actPlanName              = planNameInConfirmBudget.isDisplayed();
	   	boolean actExportBtn             = exportBtn.isDisplayed();
	   	boolean actClearBtn              = clearBtn.isDisplayed();
	   	boolean actCloseBtn              = closeBtn.isDisplayed();
	   	boolean actAccountCmbx           = accountCmbx.isDisplayed();
	   	boolean actItemCmbx              = itemCmbx.isDisplayed();
	   	boolean actDepartmetCmbx         = departmentCmbx.isDisplayed();
	   	boolean actBudgetTxtBox          = budgetTxtBox.isDisplayed();
	   	boolean actDepartmentTxtBox      = departmentTxtBox.isDisplayed();
	   	boolean actItemTxtBox            = itemTxtBox.isDisplayed();
	   	boolean actAccountTxtBox         = accountTxtBox.isDisplayed();
	   	boolean actQtyTxtBox             = qtyTxtBox.isDisplayed();
	   	boolean actValueTxtBox           = valueTxtBox.isDisplayed();
   	
	   	boolean expPlanName              = true;
	   	boolean expExportBtn             = true;
	   	boolean expClearBtn              = true;
	   	boolean expCloseBtn              = true;
	   	boolean expAccountCmbx           = true;
	   	boolean expItemCmbx              = true;
	   	boolean expDepartmentCmbx        = true;
	   	boolean expBudgetTxtBox          = true;
	   	boolean expDepartmentTxtBox      = true;
	   	boolean expItemTxtBox            = true;
	   	boolean expAccountTxtBox         = true;
	   	boolean expQtyTxtBox             = true;
	   	boolean expValueTxtBox           = true;
	   	
   	
	  	System.out.println("Budget Menu Define Budget         : " + actPlanName +           "  value expected  " + expPlanName);
	  	System.out.println("Budget Menu Confirm Budget        : " + actExportBtn +          "  value expected  " + expExportBtn);
	  	System.out.println("Budget Menu Budget Budget         : " + actClearBtn +           "  value expected  " + expClearBtn);
	  	System.out.println("Budget Menu Approval Budget       : " + actCloseBtn +           "  value expected  " + expCloseBtn);
	  	System.out.println("Budget Menu Rejected Budget       : " + actAccountCmbx +        "  value expected  " + expAccountCmbx);
	  	System.out.println("Budget Menu Revise Budget Budget  : " + actItemCmbx +           "  value expected  " + expItemCmbx);
	  	System.out.println("Budget Menu Budget Report Budget  : " + actDepartmentTxtBox +   "  value expected  " + expDepartmentTxtBox);
	  	System.out.println("Budget Menu Define Budget         : " + actItemTxtBox +         "  value expected  " + expItemTxtBox);
	  	System.out.println("Budget Menu Confirm Budget        : " + actAccountTxtBox +      "  value expected  " + expAccountTxtBox);
	  	System.out.println("Budget Menu Revise Budget         : " + actQtyTxtBox +          "  value expected  " + expQtyTxtBox);
	  	System.out.println("Budget Menu Rejected Budget       : " + actValueTxtBox +        "  value expected  " + expValueTxtBox);	
  	
		if( actPlanName==expPlanName && actExportBtn==expExportBtn && actClearBtn==expClearBtn
				&& actCloseBtn==expCloseBtn && actAccountCmbx==expAccountCmbx && actItemCmbx==expItemCmbx
				&& actDepartmentTxtBox==expDepartmentTxtBox && actItemTxtBox==expItemTxtBox  && actAccountTxtBox==expAccountTxtBox 
				&& actQtyTxtBox==expQtyTxtBox && actValueTxtBox==expValueTxtBox)
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
			return false;
		}
	}
	/*catch (Exception e) 
    {
	   String exception =e.getMessage().substring(0, 50);
       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
       return false;
    }
  }*/



	public static boolean checkConfirmBudgetValueIsDisplay() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	 excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
	 /*try
	 {*/
       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameInConfirmBudget));     
       planNameInConfirmBudget.click();
       planNameInConfirmBudget.sendKeys("ConfirmPlanBudget");
       Thread.sleep(2000);
       planNameInConfirmBudget.sendKeys(Keys.TAB);
       
       Thread.sleep(2000);
       
	    String actPlanName              = planNameInConfirmBudget.getAttribute("value");
	    String actVersionTxt            = versionTxtBox.getText();
	   	String actBudgetTxtBox          = budgetTxtBox.getText();
	   	String actDepartmentTxtBox      = departmentTxtBox.getText();
	   	String actItemTxtBox            = itemTxtBox.getText();
	   	String actAccountTxtBox         = accountTxtBox.getText();
	   	String actQtyTxtBox             = qtyTxtBox.getText();
	   	String actValueTxtBox           = valueTxtBox.getText();
   	    String actBudgetTotalValue      = budgetTotalValue.getText();
   	    String actBudgetTotalQty        = budgetTotalQty.getText();
	   	
	   	
	   	String expPlanName              = "ConfirmPlanBudget";
	   	String expVersionTxt             = "1";
	   	String expBudgetTxtBox          = "By Product By Qty By Value";
	   	String expDepartmentTxtBox      = "DUBAI";
	   	String expItemTxtBox            = "BR COGS ITEM";
	   	String expAccountTxtBox         = "BR COGS ACC INV";
	   	String expQtyTxtBox             = "12.00";
	   	String expValueTxtBox           = "120.00";
        String expBudgetTotalValue      = "120.00";
	   	String expBudgetTotalQty        = "12.00";
   	
	  	System.out.println("Budget Menu Define Budget         : " + actPlanName +           "  value expected  " + expPlanName);
	 	System.out.println("Budget Menu Approval Budget       : " + actVersionTxt +         "  value expected  " + expVersionTxt);
	  	System.out.println("Budget Menu Budget Report Budget  : " + actDepartmentTxtBox +   "  value expected  " + expDepartmentTxtBox);
	  	System.out.println("Budget Menu Define Budget         : " + actItemTxtBox +         "  value expected  " + expItemTxtBox);
	  	System.out.println("Budget Menu Confirm Budget        : " + actAccountTxtBox +      "  value expected  " + expAccountTxtBox);
	  	System.out.println("Budget Menu Revise Budget         : " + actQtyTxtBox +          "  value expected  " + expQtyTxtBox);
	  	System.out.println("Budget Menu Rejected Budget       : " + actValueTxtBox +        "  value expected  " + expValueTxtBox);	
		System.out.println("Budget Menu Confirm Budget        : " + actBudgetTotalValue +   "  value expected  " + expBudgetTotalValue);
	  	System.out.println("Budget Menu Budget Budget         : " + actBudgetTotalQty +     "  value expected  " + expBudgetTotalQty);
	  
	  	
		if( actPlanName.equalsIgnoreCase(expPlanName) && actVersionTxt.equalsIgnoreCase(expVersionTxt) 
				&& actBudgetTotalQty.equalsIgnoreCase(expBudgetTotalQty) && actVersionTxt.equalsIgnoreCase(expVersionTxt) 
				&& actDepartmentTxtBox.equalsIgnoreCase(expDepartmentTxtBox) && actItemTxtBox.equalsIgnoreCase(expItemTxtBox)
				&& actAccountTxtBox.equalsIgnoreCase(expAccountTxtBox) && actQtyTxtBox.equalsIgnoreCase(expQtyTxtBox) 
				&& actValueTxtBox.equalsIgnoreCase(expValueTxtBox))
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
			return false;
		}
	}
	/*catch (Exception e) 
    {
	   String exception =e.getMessage().substring(0, 50);
       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
       return false;
    }
  }*/
	 
	 
	
	
	public static boolean checkNonConfirmBudgetValueIsDisplay() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	 excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
	 /*try
	 {*/
       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameInConfirmBudget));     
       planNameInConfirmBudget.click();
       planNameInConfirmBudget.sendKeys(Keys.HOME);
       planNameInConfirmBudget.sendKeys(Keys.SHIFT,Keys.END);
       planNameInConfirmBudget.sendKeys("NonConfirmPlanBudget");
       Thread.sleep(2000);
       planNameInConfirmBudget.sendKeys(Keys.TAB);
       
       Thread.sleep(2000);
       
	    String actPlanName              = planNameInConfirmBudget.getAttribute("value");
	    String actVersionTxt            = versionTxtBox.getText();
	   	boolean actBudgetTxtBox         = budgetTxtBox.getText().isEmpty();
	   	boolean actDepartmentTxtBox     = departmentTxtBox.getText().isEmpty();
	    boolean actItemTxtBox           = itemTxtBox.getText().isEmpty();
	   	boolean actAccountTxtBox        = accountTxtBox.getText().isEmpty();
	   	boolean actQtyTxtBox            = qtyTxtBox.getText().isEmpty();
	   	boolean actValueTxtBox          = valueTxtBox.getText().isEmpty();
   	    String actBudgetTotalValue      = budgetTotalValue.getText();
   	    String actBudgetTotalQty        = budgetTotalQty.getText();
	   	
   	   
	   	
	   	String expPlanName              = "NonConfirmPlanBudget";
	   	String expVersionTxt            = "1";
	   	boolean expBudgetTxtBox         = true;
	   	boolean expDepartmentTxtBox     = true;
	   	boolean expItemTxtBox           = true;
	   	boolean expAccountTxtBox        = true;
	   	boolean expQtyTxtBox            = true;
	   	boolean expValueTxtBox          = true;
        String expBudgetTotalValue      = "0.00";
	   	String expBudgetTotalQty        = "0";
   	
	  	System.out.println("Budget Menu Define Budget         : " + actPlanName +           "  value expected  " + expPlanName);
	 	System.out.println("Budget Menu Approval Budget       : " + actVersionTxt +         "  value expected  " + expVersionTxt);
	  	System.out.println("Budget Menu Budget Report Budget  : " + actDepartmentTxtBox +   "  value expected  " + expDepartmentTxtBox);
	  	System.out.println("Budget Menu Define Budget         : " + actItemTxtBox +         "  value expected  " + expItemTxtBox);
	  	System.out.println("Budget Menu Confirm Budget        : " + actAccountTxtBox +      "  value expected  " + expAccountTxtBox);
	  	System.out.println("Budget Menu Revise Budget         : " + actQtyTxtBox +          "  value expected  " + expQtyTxtBox);
	  	System.out.println("Budget Menu Rejected Budget       : " + actValueTxtBox +        "  value expected  " + expValueTxtBox);	
		System.out.println("Budget Menu Confirm Budget        : " + actBudgetTotalValue +   "  value expected  " + expBudgetTotalValue);
	  	System.out.println("Budget Menu Budget Budget         : " + actBudgetTotalQty +     "  value expected  " + expBudgetTotalQty);
	  
	  	
		if( actPlanName.equalsIgnoreCase(expPlanName) && actVersionTxt.equalsIgnoreCase(expVersionTxt) 
			    && actDepartmentTxtBox==expDepartmentTxtBox 
				&& actItemTxtBox==expItemTxtBox && actAccountTxtBox==expAccountTxtBox 
				&& actQtyTxtBox==expQtyTxtBox && actValueTxtBox==expValueTxtBox 
				&& actBudgetTotalValue.equalsIgnoreCase(expBudgetTotalValue) && actBudgetTotalQty.equalsIgnoreCase(expBudgetTotalQty))
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
			return false;
		}
	}
	/*catch (Exception e) 
    {
	   String exception =e.getMessage().substring(0, 50);
       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
       return false;
    }
  }*/
	
  public boolean checkCloseButtonInConfirmBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
  {
	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeBtn));
	closeBtn.click();
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
	
	boolean actdashboard=dashboard.isDisplayed();
	boolean expdashboard=true;
		
	System.out.println("dashboard  :"+actdashboard +"  Value Expected  : "+expdashboard);
	
	if(actdashboard==expdashboard)	
	{
		System.out.println("Test Pass : Cancel Option Working ");
		excelReader.setCellData(xlfile, xlSheetName, 53, 9, resPass);
		return true;
	}
	else
	{
		System.out.println("Test Fail : Cancel Option Not Working");
		excelReader.setCellData(xlfile, xlSheetName, 53, 9, resFail);
		return false;
	}
   }

	
	

	
	    //APPEND BUDGET
	
		@FindBy(xpath="//input[@id='optPlanName']")
		private static WebElement planNameTxt;
		
		@FindBy(xpath="//label[@id='dpDate']")
		private static WebElement dateTxt;
		
		@FindBy(xpath="//label[@id='dpValidFromDate']")
		private static WebElement validFromDateTxtArea;
		
		@FindBy(xpath="//label[@id='dpValidToDate']")
		private static WebElement valdToDateTxtArea;
		
		@FindBy(xpath="//label[@id='dpValidToDate']")
		private static WebElement dateTxtArea;
		
		@FindBy(xpath="//input[@id='uploadFile']")
		private static WebElement selectFiledTxt;
		
		@FindBy(xpath="//input[@id='uploadBtn']")
		private static WebElement uploadBtn;
		
		@FindBy(xpath="//div[contains(text(),'Split')]")
		private static WebElement bodySplitTxt;
		
		@FindBy(xpath="//div[contains(text(),'Budget Type')]")
		private static WebElement bodyBudgetTypeTxt;
		
		@FindBy(xpath="//div[contains(text(),'Account')]")
		private static WebElement bodyAccountTxt;
		
		@FindBy(xpath="//div[contains(text(),'Quantity')]")
		private static WebElement bodyQuantityTxt;
		
		@FindBy(xpath="//div[contains(text(),'Budget Value')]")
		private static WebElement bodyBudgetValueTxt;
		
		@FindBy(xpath="//div[contains(text(),'Narration')]")
		private static WebElement bodyNarrationTxt;
		
		@FindBy(xpath="//div[contains(text(),'Add Quantity')]")
		private static WebElement bodyAddQuantityTxt;
		
		@FindBy(xpath="//div[contains(text(),'Reduce Quantity')]")
		private static WebElement bodyReduceQtyTxt;
		
		@FindBy(xpath="//div[contains(text(),'Add Value')]")
		private static WebElement bodyAddValueTxt;
		
		@FindBy(xpath="//div[contains(text(),'Reduce Value')]")
		private static WebElement bodyReduceValueTxt;	
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_1-1']//i[@class='icon-edit icon-font9 editField']")
		private static WebElement grid1stRowEditBtn;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_2-1']//i[@class='icon-edit icon-font9 editField']")
		private static WebElement grid2ndRowEditBtn;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_3-1']//i[@class='icon-edit icon-font9 editField']")
		private static WebElement grid3rdRowEditBtn;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_4-1']//i[@class='icon-edit icon-font9 editField']")
		private static WebElement grid4thRowEditBtn;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_5-1']//i[@class='icon-edit icon-font9 editField']")
		private static WebElement grid5thRowEditBtn;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_6-1']//i[@class='icon-edit icon-font9 editField']")
		private static WebElement grid6thRowEditBtn;
				
		@FindBy(xpath="//tbody[@id='grdAddEditBudget_body']/tr/td[4]")
		private static WebElement gridBudgetTypeRow1Click;
		
		@FindBy(xpath="//tbody[@id='grdAddEditBudget_body']/tr/td[4]")
		private static WebElement gridBudgetTypeRow1;
		
		@FindBy(xpath="//tbody[@id='grdAddEditBudget_body']/tr/td[4]/select")
		private static WebElement gridBudgetTypeInputRow1;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_2-3']")
		private static WebElement gridBudgetTypeRow2;

		@FindBy(xpath="//td[@id='grdAddEditBudget_col_3-3']")
		private static WebElement gridBudgetTypeRow3;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_4-3']")
		private static WebElement gridBudgetTypeRow4;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_5-3']")
		private static WebElement gridBudgetTypeRow5;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_6-3']")
		private static WebElement gridBudgetTypeRow6;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[1]/div[2]/div[2]/table/tbody/tr[1]/td[5]")
		private static WebElement gridBudgetDepartmentFirstRow;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[1]/div[2]/div[2]/table/tbody/tr[1]/td[5]")
		private static WebElement gridBudgetDepartmentSecondRow;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[1]/div[2]/div[2]/table/tbody/tr[1]/td[8]")
		private static WebElement gridBudgetItemFirstRow;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[1]/div[2]/div[2]/table/tbody/tr[2]/td[8]")
		private static WebElement gridBudgetItemSecondRow;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[1]/div[2]/div[2]/table/tbody/tr[1]/td[9]")
		private static WebElement gridBudgetAccountFirstRow;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[1]/div[2]/div[2]/table/tbody/tr[2]/td[9]")
		private static WebElement gridBudgetAccountSecondRow;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[1]/div[2]/div[2]/table/tbody/tr[1]/td[10]")
		private static WebElement gridBudgetQtyFirstRow;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[1]/div[2]/div[2]/table/tbody/tr[2]/td[10]")
		private static WebElement gridBudgetQtySecondRow;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[1]/div[2]/div[2]/table/tbody/tr[1]/td[13]")
		private static WebElement gridBudgetValueFirstRow;
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div[1]/div[2]/div[2]/table/tbody/tr[2]/td[13]")
		private static WebElement gridBudgetValueSecondRow;
		
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_1-8']")
		private static WebElement gridAccountRow1;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_2-8']")
		private static WebElement gridAccountRow2;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_3-8']")
		private static WebElement gridAccountRow3;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_4-8']")
		private static WebElement gridAccountRow4;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_5-8']")
		private static WebElement gridAccountRow5;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_6-8']")
		private static WebElement gridAccountRow6;
				
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_1-9']")
		private static WebElement gridQtyRow1;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_2-9']")
		private static WebElement gridQtyRow2;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_3-9']")
		private static WebElement gridQtyRow3;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_4-9']")
		private static WebElement gridQtyRow4;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_5-9']")
		private static WebElement gridQtyRow5;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_6-9']")
		private static WebElement gridQtyRow6;
			
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_1-12']")
		private static WebElement gridBudgetValueRow1;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_2-12']")
		private static WebElement gridBudgetValueRow2;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_3-12']")
		private static WebElement gridBudgetValueRow3;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_4-12']")
		private static WebElement gridBudgetValueRow4;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_5-12']")
		private static WebElement gridBudgetValueRow5;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_6-12']")
		private static WebElement gridBudgetValueRow6;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_1-15']")
		private static WebElement gridNarrationRow1;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_2-15']")
		private static WebElement gridNarrationRow2;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_3-15']")
		private static WebElement gridNarrationRow3;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_4-15']")
		private static WebElement gridNarrationRow4;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_5-15']")
		private static WebElement gridNarrationRow5;
		
		@FindBy(xpath="//td[@id='grdAddEditBudget_col_6-15']")
		private static WebElement gridNarrationRow6;
				
		
		@FindBy(xpath="//input[@id='txtGridNarration']")
		private static WebElement gridEnterNarrationTxt;
		
		@FindBy(xpath="//select[@id='grdAddEditBudget_cmbbudgettype']")
		private static WebElement gridEnterBudgetTypeDrpdwn;
		
		@FindBy(xpath="//input[@id='grdAddEditBudget_option_account']")
		private static WebElement gridEnterAccountDrpdwn;
		
		@FindBy(xpath="//input[@id='grdBudgetValue']")
		private static WebElement ap_gridEnterBudgetValue;
		
		@FindBy(xpath="//span[@id='btnSave']")
		private static WebElement budgetAppendSaveBtn;
		
		@FindBy(xpath="//span[@class='icon-close icon-font7']")
		private static WebElement cancelBtn;
		
	
		
		public static boolean checkReviseAppendBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		 /*try
		 {*/
	   	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		   financialsMenu.click();

		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetMenu));
		   financialsBudgetMenu.click();
			 
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetReviseBudget));
		   financialsBudgetReviseBudget.click();
		   
		   Thread.sleep(3000);
	      		   
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(appendBudgetMenu));
		   boolean actReviseMenuAppend      = appendBudgetMenu.isDisplayed();
           boolean actReviseMenuAddOrReduce = addORReduceBudgetMenu.isDisplayed();
		   boolean actReviseMenuTransfer    = transferBudgetMenu.isDisplayed(); 
		 		   
		   boolean expReviseMenuAppend      = true;
		   boolean expReviseMenuAddOrReduce = true;
		   boolean expReviseMenuTransfer    = true;
		   
		   System.out.println("Append Budget          : " + actReviseMenuAppend +       "  value expected  " + expReviseMenuAppend);
		   System.out.println("Add Or Reduce Budget   : " + actReviseMenuAddOrReduce +  "  value expected  " + expReviseMenuAddOrReduce);
		   System.out.println("Transfer Budget        : " + actReviseMenuTransfer +     "  value expected  " + expReviseMenuTransfer);
		  	
		   boolean actMenusOfReviseBudget = actReviseMenuAppend && actReviseMenuAddOrReduce && actReviseMenuTransfer;
		   boolean expMenusOfReviseBudget = true;
		   
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(appendBudgetMenu));
		   appendBudgetMenu.click();
		   
		   Thread.sleep(2000);
		   
	       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));
		   boolean actPlanName              = planNameTxt.isDisplayed();
		   boolean actSelectFieldTxt        = selectFiledTxt.isDisplayed();
		   boolean actUploadBtn             = uploadBtn.isDisplayed();
		   boolean actBodyBudgetType        = bodyBudgetTypeTxt.isDisplayed();		 
		   boolean actDepartmentTxtBox      = gridBudgetDepartmentFirstRow.isDisplayed();
		   boolean actItemTxtBox            = gridBudgetItemFirstRow.isDisplayed();
		   boolean actAccountTxtBox         = bodyAccountTxt.isDisplayed();
		   boolean actQtyTxtBox             = bodyQuantityTxt.isDisplayed(); 
		   boolean actValueTxtBox           = bodyBudgetValueTxt.isDisplayed();
		 
		   boolean expPlanName              = true;
		   boolean expSelectFieldTxt        = true;
		   boolean expUploadBtn             = true;
		   boolean expBodyBudgetType        = true;	   
		   boolean expDepartmentTxtBox      = true;
		   boolean expItemTxtBox            = true;
		   boolean expAccountTxtBox         = true;
		   boolean expQtyTxtBox             = true;
		   boolean expValueTxtBox           = true;
		   	   	
		  	System.out.println("Budget Menu Define Budget         : " + actPlanName +            "  value expected  " + expPlanName);
		  	System.out.println("Budget Menu Confirm Budget        : " + actSelectFieldTxt +      "  value expected  " + expSelectFieldTxt);
		  	System.out.println("Budget Menu Budget Budget         : " + actUploadBtn +           "  value expected  " + expUploadBtn);
		  	System.out.println("Budget Menu Approval Budget       : " + actBodyBudgetType +      "  value expected  " + expBodyBudgetType);
		  	System.out.println("Budget Menu Budget Report Budget  : " + actDepartmentTxtBox +    "  value expected  " + expDepartmentTxtBox);
		  	System.out.println("Budget Menu Define Budget         : " + actItemTxtBox +          "  value expected  " + expItemTxtBox);
		  	System.out.println("Budget Menu Confirm Budget        : " + actAccountTxtBox +       "  value expected  " + expAccountTxtBox);
		  	System.out.println("Budget Menu Revise Budget         : " + actQtyTxtBox +           "  value expected  " + expQtyTxtBox);
		  	System.out.println("Budget Menu Rejected Budget       : " + actValueTxtBox +         "  value expected  " + expValueTxtBox);	
	  	
			if( actPlanName==expPlanName && actDepartmentTxtBox==expDepartmentTxtBox 
					&& actItemTxtBox==expItemTxtBox  && actAccountTxtBox==expAccountTxtBox 
					&& actQtyTxtBox==expQtyTxtBox && actValueTxtBox==expValueTxtBox
					&& actMenusOfReviseBudget==expMenusOfReviseBudget)				
			{
				excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
				return true;
			}
			else
			{
				excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
		   String exception =e.getMessage().substring(0, 50);
	       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	       return false;
	    }
	  }*/
		
		
		@FindBy(xpath="//tbody[@id='grdAddEditBudget_body']/tr/td[5]")
		private static WebElement  budgetAppendDepartmentR1C3SelectCell; 
	
		@FindBy(xpath="//table[@id='grdAddEditBudget_option_Tag1_input_container']/tbody/tr[2]/td/input[2]")
		private static WebElement  budgetAppendDepartmentR1C3InputCell; 
				
		@FindBy(xpath="//*[@id='grdAddEditBudget_option_product']")
		private static WebElement  budgetAppendItemR1C3SelectCell; 
	
		@FindBy(xpath="//*[@id='grdAddEditBudget_option_account']")
		private static WebElement  budgetAppendAccountR1C3SelectCell; 
	
		@FindBy(xpath="//*[@id='grdQuantity']")
		private static WebElement  budgetAppendQtyR1C3SelectCell;
		
		@FindBy(xpath="//*[@id='grdBudgetValue']")
		private static WebElement  budgetAppendValueR1C3SelectCell;
		
		@FindBy(xpath="//*[@id='txtGridNarration']")
		private static WebElement  budgetAppendNarrationR1C3SelectCell;
		
		public static boolean checkSaveConfirmPlanBudgetAppendBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		 	
		 	/*try
		 	{*/ 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));     	
		 		planNameTxt.sendKeys("ConfirmPlanBudget");
		 		Thread.sleep(2000);
		 		
		 		planNameTxt.sendKeys(Keys.TAB);		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(gridBudgetTypeRow1));     	
		 		gridBudgetTypeRow1.click();
		 			 		
		 		gridBudgetTypeInputRow1.click();
		 		gridBudgetTypeInputRow1.sendKeys("By Product By Qty", Keys.ENTER);
		 
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAppendDepartmentR1C3SelectCell));     	
		 		budgetAppendDepartmentR1C3SelectCell.click();	 	
		 		
		 		budgetAppendDepartmentR1C3InputCell.sendKeys("DUBAI");
		 		Thread.sleep(2000);
		 		budgetAppendDepartmentR1C3InputCell.sendKeys(Keys.TAB);
		 			
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAppendItemR1C3SelectCell));     	
		 		budgetAppendItemR1C3SelectCell.click();	 	
		 	
		 		budgetAppendItemR1C3SelectCell.sendKeys("STD RATE COGS ITEM");
		 		Thread.sleep(2000);
		 		budgetAppendItemR1C3SelectCell.sendKeys(Keys.TAB);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAppendAccountR1C3SelectCell));     	
		 		budgetAppendAccountR1C3SelectCell.click();	 	
		 		
		 		Thread.sleep(5000);
		
		 		System.out.println( budgetAppendAccountR1C3SelectCell.getText());
		 		System.out.println( budgetAppendAccountR1C3SelectCell.getAttribute("value"));
		 		
		 		String accGetAccountName = budgetAppendAccountR1C3SelectCell.getAttribute("value");
		 		String expGetAccountName = "STD RATE COGS ACC INV";
		 		
		 		System.out.println("Account Name Value Actual : " + accGetAccountName + "Value Expected : " + expGetAccountName);
		 				 		
		 		//budgetAppendAccountR1C3SelectCell.sendKeys("STD RATE COGS ACC INV");
		 		Thread.sleep(2000);
		 		budgetAppendAccountR1C3SelectCell.sendKeys(Keys.TAB);
		 		
		 		System.out.println("q"+budgetAppendAccountR1C3SelectCell.getText());
		 		System.out.println( budgetAppendAccountR1C3SelectCell.getAttribute("value"));
		 	
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAppendQtyR1C3SelectCell));     	
		 		budgetAppendQtyR1C3SelectCell.click();	 	
		 				 	
		 		budgetAppendQtyR1C3SelectCell.sendKeys("24");
		 		Thread.sleep(2000);
		 		budgetAppendQtyR1C3SelectCell.sendKeys(Keys.TAB);
		 				 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAppendValueR1C3SelectCell));     	
		 		budgetAppendValueR1C3SelectCell.click();	 	
		 		Thread.sleep(2000);
		 		
		 		budgetAppendValueR1C3SelectCell.sendKeys("240");
		 		budgetAppendValueR1C3SelectCell.sendKeys(Keys.TAB);	 		
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAppendNarrationR1C3SelectCell));     	
		 		budgetAppendNarrationR1C3SelectCell.click();	 	
		 		Thread.sleep(2000);
		 		
		 		budgetAppendNarrationR1C3SelectCell.sendKeys("Test");
		 				 	 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAppendSaveBtn));     	
		 		budgetAppendSaveBtn.click();			 		
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
		 		String actErrorMessage=errorMessage.getText();
		 		String expErrorMessage="Add Budget Save Successfully";

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
		 		errorMessageCloseBtn.click();
			 		
		 		System.out.println("ValidationMessage  :  " + actErrorMessage +" Value Expected : " + expErrorMessage);
		 			 			 		
		 		if( accGetAccountName.equalsIgnoreCase(expGetAccountName) && actErrorMessage.equalsIgnoreCase(expErrorMessage))
		 		{
		 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
		 			return true;
		 		}
		 		else
		 		{
		 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
		 			return false;
		 		}
		 	}
		 	/*catch (Exception e) 
		    {
	          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	          return false;
		    }
		  }*/
		
		 
		public boolean checkCloseButtonInBudgetAppend() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
				excelReader=new ExcelReader(POJOUtility.getExcelPath());
				xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
				cancelBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
				
				boolean actdashboard=dashboard.isDisplayed();
				boolean expdashboard=true;
					
				System.out.println("dashboard  :"+actdashboard +"  Value Expected  : "+expdashboard);
				
				if(actdashboard==expdashboard)	
				{
					System.out.println("Test Pass : Cancel Option Working ");
					excelReader.setCellData(xlfile, xlSheetName, 53, 9, resPass);
					return true;
				}
				else
				{
					System.out.println("Test Fail : Cancel Option Not Working");
					excelReader.setCellData(xlfile, xlSheetName, 53, 9, resFail);
					return false;
				}
			}
 
		
		
		//ADD OR REDUCE THE BUDGET
		public static boolean checkReviseADDOrREDUCEBudgetMenu() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		 /*try
		 {*/
	   	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		   financialsMenu.click();

		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetMenu));
		   financialsBudgetMenu.click();
			 
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetReviseBudget));
		   financialsBudgetReviseBudget.click();
	      		   
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addORReduceBudgetMenu));
		   addORReduceBudgetMenu.click();
		   
		   Thread.sleep(3000);
				   
	       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));
		   boolean actPlanName              = planNameTxt.isDisplayed();
		   boolean actSelectFieldTxt        = selectFiledTxt.isDisplayed();
		   boolean actUploadBtn             = uploadBtn.isDisplayed();
		   boolean actBodyBudgetType        = bodyBudgetTypeTxt.isDisplayed();		 
		   boolean actDepartmentTxtBox      = gridBudgetDepartmentFirstRow.isDisplayed();
		   boolean actItemTxtBox            = gridBudgetItemFirstRow.isDisplayed();
		   boolean actAccountTxtBox         = bodyAccountTxt.isDisplayed();
		   boolean actQtyTxtBox             = bodyQuantityTxt.isDisplayed(); 
		   boolean actValueTxtBox           = bodyBudgetValueTxt.isDisplayed();
		 
		   boolean expPlanName              = true;
		   boolean expSelectFieldTxt        = true;
		   boolean expUploadBtn             = true;
		   boolean expBodyBudgetType        = true;	   
		   boolean expDepartmentTxtBox      = true;
		   boolean expItemTxtBox            = true;
		   boolean expAccountTxtBox         = true;
		   boolean expQtyTxtBox             = true;
		   boolean expValueTxtBox           = true;
		   	   	
		   System.out.println("Budget Menu Define Budget         : " + actPlanName +            "  value expected  " + expPlanName);
		   System.out.println("Budget Menu Confirm Budget        : " + actSelectFieldTxt +      "  value expected  " + expSelectFieldTxt);
		   System.out.println("Budget Menu Budget Budget         : " + actUploadBtn +           "  value expected  " + expUploadBtn);
		   System.out.println("Budget Menu Approval Budget       : " + actBodyBudgetType +      "  value expected  " + expBodyBudgetType);
		   System.out.println("Budget Menu Budget Report Budget  : " + actDepartmentTxtBox +    "  value expected  " + expDepartmentTxtBox);
		   System.out.println("Budget Menu Define Budget         : " + actItemTxtBox +          "  value expected  " + expItemTxtBox);
		   System.out.println("Budget Menu Confirm Budget        : " + actAccountTxtBox +       "  value expected  " + expAccountTxtBox);
		   System.out.println("Budget Menu Revise Budget         : " + actQtyTxtBox +           "  value expected  " + expQtyTxtBox);
		   System.out.println("Budget Menu Rejected Budget       : " + actValueTxtBox +         "  value expected  " + expValueTxtBox);	
	  	
			if( actPlanName==expPlanName && actDepartmentTxtBox==expDepartmentTxtBox 
					&& actItemTxtBox==expItemTxtBox  && actAccountTxtBox==expAccountTxtBox 
					&& actQtyTxtBox==expQtyTxtBox && actValueTxtBox==expValueTxtBox)				
			{
				excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
				return true;
			}
			else
			{
				excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
		   String exception =e.getMessage().substring(0, 50);
	       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	       return false;
	    }
	  }*/

		
		public static boolean checkConfirmPlanBudgetAddOrReduceBudgetLoadingValues() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		 	
		 	/*try
		 	{*/ 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));     	
		 		planNameTxt.sendKeys("ConfirmPlanBudget");
		 		Thread.sleep(2000);
		 		
		 		planNameTxt.sendKeys(Keys.TAB);		
		 		Thread.sleep(2000);		 	
		 			 		
		 		System.out.println(budgetAppendDepartmentR1C3SelectCell.getText());
		 		
		 		String actDepartmentValueR1   = gridBudgetDepartmentFirstRow.getText();
		 		String actDepartmentValueR2   = gridBudgetDepartmentSecondRow.getText();
		 		String actItemValueR1         = gridBudgetItemFirstRow.getText();
		 		String actItemValueR2         = gridBudgetItemSecondRow.getText();
		 		String actAccountValueR1      = gridBudgetAccountFirstRow.getText();
		 		String actAccountValueR2      = gridBudgetAccountSecondRow.getText();
		 		String actQtyValueR1          = gridBudgetQtyFirstRow.getText();
		 		String actQtyValueR2          = gridBudgetQtySecondRow.getText();
		 		String actValueR1             = gridBudgetValueFirstRow.getText();
		 		String actValueR2             = gridBudgetValueSecondRow.getText();
		 		
		 		String expDepartmentValueR1   = "DUBAI";
		 		String expDepartmentValueR2   = "DUBAI";
		 		String expItemValueR1         = "BR COGS ITEM";
		 		String expItemValueR2         = "STD RATE COGS ITEM";
		 		String expAccountValueR1      = "BR COGS ACC INV";
		 		String expAccountValueR2      = "STD RATE COGS ACC INV";
		 		String expQtyValueR1          = "12.00";
		 		String expQtyValueR2          = "24.00";
		 		String expValueR1             = "120.00";
		 		String expValueR2             = "240.00";
		 		
		 		
		 		System.out.println("Budget Menu Define Budget    : " + actDepartmentValueR1 +   "  value expected  " + expDepartmentValueR1);
			    System.out.println("Budget Menu Confirm Budget   : " + actDepartmentValueR2 +   "  value expected  " + expDepartmentValueR2);
			   
			    System.out.println("Budget Menu Define Budget    : " + actItemValueR1 +         "  value expected  " + expItemValueR1);
			    System.out.println("Budget Menu Confirm Budget   : " + actItemValueR2 +         "  value expected  " + expItemValueR2);
		
			    System.out.println("Budget Menu Define Budget    : " + actAccountValueR1 +      "  value expected  " + expAccountValueR1);
			    System.out.println("Budget Menu Confirm Budget   : " + actAccountValueR2 +      "  value expected  " + expAccountValueR2);
			   
			    System.out.println("Budget Menu Define Budget    : " + actQtyValueR1 +          "  value expected  " + expQtyValueR1);
			    System.out.println("Budget Menu Confirm Budget   : " + actQtyValueR2 +          "  value expected  " + expQtyValueR2);
			   
			    System.out.println("Budget Menu Define Budget    : " + actValueR1 +             "  value expected  " + expValueR1);
			    System.out.println("Budget Menu Confirm Budget   : " + actValueR2 +             "  value expected  " + expValueR2);
			   
		 		
		 		if(actDepartmentValueR1.equalsIgnoreCase(expDepartmentValueR1) && actDepartmentValueR2.equalsIgnoreCase(expDepartmentValueR2)
		 				&& actItemValueR1.equalsIgnoreCase(expItemValueR1) && actItemValueR2.equalsIgnoreCase(expItemValueR2)
		 				&& actAccountValueR1.equalsIgnoreCase(expAccountValueR1) && actAccountValueR2.equalsIgnoreCase(expAccountValueR2)
		 				&& actQtyValueR1.equalsIgnoreCase(expQtyValueR1) && actQtyValueR2.equalsIgnoreCase(expQtyValueR2)
		 				&& actValueR1.equalsIgnoreCase(expValueR1) && actValueR2.equalsIgnoreCase(expValueR2))
		 		{
		 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
		 			return true;
		 		}
		 		else
		 		{
		 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
		 			return false;
		 		}
		 	}
		 	/*catch (Exception e) 
		    {
	          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	          return false;
		    }
		  }*/
		
		

	
	
	@FindBy(xpath="//*[@id='grdAddEditBudget_body']/tr[1]/td[12]")
	private static WebElement  budgetAddOrReduceReduceQtySelect; 

	@FindBy(xpath="//*[@id='grdAddEditBudget_txtReduceQuantity']")
	private static WebElement  budgetAddOrReduceReduceQtyInput; 
			
	@FindBy(xpath="//*[@id='grdAddEditBudget_body']/tr[2]/td[11]")
	private static WebElement  budgetAddOrReduceAddQtySelect; 

	@FindBy(xpath="//*[@id='grdAddEditBudget_txtAddQuantity']")
	private static WebElement  budgetAddOrReduceAddQtyInput; 

	@FindBy(xpath="//*[@id='grdAddEditBudget_body']/tr[1]/td[15]")
	private static WebElement  budgetAddReduceReduceValueSelect;
	
	@FindBy(xpath="//*[@id='grdAddEditBudget_txtReduceValue']")
	private static WebElement  budgetAddReduceReduceValueInput;
	
	@FindBy(xpath="//*[@id='grdAddEditBudget_body']/tr[2]/td[14]")
	private static WebElement  budgetAddReduceAddValueSelect;
	
	@FindBy(xpath="//*[@id='grdAddEditBudget_txtAddValue']")
	private static WebElement  budgetAddReduceAddValueInput;
	
	
	public static boolean checkSaveConfirmPlanBudgetAddOrReduce() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	 	
	 	try
	 	{ 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAddOrReduceReduceQtySelect));     	
	 		budgetAddOrReduceReduceQtySelect.click();	 	
	 		
	 		budgetAddOrReduceReduceQtyInput.sendKeys("6");
	 		Thread.sleep(2000);
	 		budgetAddOrReduceReduceQtyInput.sendKeys(Keys.TAB);
	 			
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAddOrReduceAddQtySelect));     	
	 		budgetAddOrReduceAddQtySelect.click();	 	
	 	
	 		budgetAddOrReduceAddQtyInput.sendKeys("6");
	 		Thread.sleep(2000);
	 		budgetAddOrReduceAddQtyInput.sendKeys(Keys.TAB);
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAddReduceReduceValueSelect));     	
	 		budgetAddReduceReduceValueSelect.click();	 	
	 		
	 		budgetAddReduceReduceValueInput.sendKeys("60");
	 		Thread.sleep(2000);
	 		budgetAddReduceReduceValueInput.sendKeys(Keys.TAB);
	 			
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAddReduceAddValueSelect));     	
	 		budgetAddReduceAddValueSelect.click();	 	
	 	
	 		budgetAddReduceAddValueInput.sendKeys("120");
	 		Thread.sleep(2000);
	 		budgetAddReduceAddValueInput.sendKeys(Keys.TAB); 	 		
	 		
	 		
	 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAppendSaveBtn));     	
	 		budgetAppendSaveBtn.click();			 		
	 		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
	 		String actErrorMessage=errorMessage.getText();
	 		String expErrorMessage="Modify Budget Save Successfully";

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
	 		errorMessageCloseBtn.click();
		 		
	 		System.out.println("ValidationMessage  :  " + actErrorMessage +" Value Expected : " + expErrorMessage);
	 			 			 		
	 		if(actErrorMessage.equalsIgnoreCase(expErrorMessage))
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
	 			return true;
	 		}
	 		else
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
	 			return false;
	 		}
	 	}
	 	catch (Exception e) 
	    {
        excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
        System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
        return false;
	    }
	  }
		
	
	
	public static boolean checkConfirmPlanBudgetAddOrReduceBudgetUpdateValues() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	 	
	 	/*try
	 	{ */		
	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));     	
	 		planNameTxt.sendKeys("ConfirmPlanBudget");
	 		Thread.sleep(2000);
	 		
	 		planNameTxt.sendKeys(Keys.TAB);		
	 		Thread.sleep(2000);		 	
	 			 		
	 		System.out.println(budgetAppendDepartmentR1C3SelectCell.getText());
	 		
	 		String actDepartmentValueR1   = gridBudgetDepartmentFirstRow.getText();
	 		String actDepartmentValueR2   = gridBudgetDepartmentSecondRow.getText();
	 		String actItemValueR1         = gridBudgetItemFirstRow.getText();
	 		String actItemValueR2         = gridBudgetItemSecondRow.getText();
	 		String actAccountValueR1      = gridBudgetAccountFirstRow.getText();
	 		String actAccountValueR2      = gridBudgetAccountSecondRow.getText();
	 		String actQtyValueR1          = gridBudgetQtyFirstRow.getText();
	 		String actQtyValueR2          = gridBudgetQtySecondRow.getText();
	 		String actValueR1             = gridBudgetValueFirstRow.getText();
	 		String actValueR2             = gridBudgetValueSecondRow.getText();
	 		
	 		String expDepartmentValueR1   = "DUBAI";
	 		String expDepartmentValueR2   = "DUBAI";
	 		String expItemValueR1         = "BR COGS ITEM";
	 		String expItemValueR2         = "STD RATE COGS ITEM";
	 		String expAccountValueR1      = "BR COGS ACC INV";
	 		String expAccountValueR2      = "STD RATE COGS ACC INV";
	 		String expQtyValueR1          = "6.00";
	 		String expQtyValueR2          = "30.00";
	 		String expValueR1             = "60.00";
	 		String expValueR2             = "360.00";
	 		
	 		
	 		System.out.println("Budget Menu Define Budget    : " + actDepartmentValueR1 +   "  value expected  " + expDepartmentValueR1);
		    System.out.println("Budget Menu Confirm Budget   : " + actDepartmentValueR2 +   "  value expected  " + expDepartmentValueR2);
		   
		    System.out.println("Budget Menu Define Budget    : " + actItemValueR1 +         "  value expected  " + expItemValueR1);
		    System.out.println("Budget Menu Confirm Budget   : " + actItemValueR2 +         "  value expected  " + expItemValueR2);
	
		    System.out.println("Budget Menu Define Budget    : " + actAccountValueR1 +      "  value expected  " + expAccountValueR1);
		    System.out.println("Budget Menu Confirm Budget   : " + actAccountValueR2 +      "  value expected  " + expAccountValueR2);
		   
		    System.out.println("Budget Menu Define Budget    : " + actQtyValueR1 +          "  value expected  " + expQtyValueR1);
		    System.out.println("Budget Menu Confirm Budget   : " + actQtyValueR2 +          "  value expected  " + expQtyValueR2);
		   
		    System.out.println("Budget Menu Define Budget    : " + actValueR1 +             "  value expected  " + expValueR1);
		    System.out.println("Budget Menu Confirm Budget   : " + actValueR2 +             "  value expected  " + expValueR2);
		   
	 		
	 		if(actDepartmentValueR1.equalsIgnoreCase(expDepartmentValueR1) && actDepartmentValueR2.equalsIgnoreCase(expDepartmentValueR2)
	 				&& actItemValueR1.equalsIgnoreCase(expItemValueR1) && actItemValueR2.equalsIgnoreCase(expItemValueR2)
	 				&& actAccountValueR1.equalsIgnoreCase(expAccountValueR1) && actAccountValueR2.equalsIgnoreCase(expAccountValueR2)
	 				&& actQtyValueR1.equalsIgnoreCase(expQtyValueR1) && actQtyValueR2.equalsIgnoreCase(expQtyValueR2)
	 				&& actValueR1.equalsIgnoreCase(expValueR1) && actValueR2.equalsIgnoreCase(expValueR2))
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
	 			return true;
	 		}
	 		else
	 		{
	 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
	 			return false;
	 		}
	 	}
	 	/*catch (Exception e) 
	    {
          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
          return false;
	    }
	  }*/
		
	
	
	
	
	
	public boolean checkCloseButtonInBudgetADDORREDUCEValue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
			cancelBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
			
			boolean actdashboard=dashboard.isDisplayed();
			boolean expdashboard=true;
				
			System.out.println("dashboard  :"+actdashboard +"  Value Expected  : "+expdashboard);
			
			if(actdashboard==expdashboard)	
			{
				System.out.println("Test Pass : Cancel Option Working ");
				excelReader.setCellData(xlfile, xlSheetName, 53, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Cancel Option Not Working");
				excelReader.setCellData(xlfile, xlSheetName, 53, 9, resFail);
				return false;
			}
	}

	
       // TRANSFER BUDGETS
	
		@FindBy(xpath="//label[@id='txtVersion']")
		private static WebElement budgetVersionTxt;
		
		@FindBy(xpath="//label[@id='lblValidToDate']")
		private static WebElement validToDateTxt;
		
		@FindBy(xpath="//label[@id='lblLastUpdateDate']")
		private static WebElement lastUpdateTxt;
		
		@FindBy(xpath="//label[@id='txtReviseNo']")
		private static WebElement reviseTxt;
		
		
		@FindBy(xpath="//*[@id='myNavbar']/ul/li[1]/div/div[2]")
		private static WebElement budgetsTransferSaveBtn;
		
		
		// TRANSFER FROM

		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div[1]/div/div[2]/div[5]/div[1]/fieldset/div/div[2]/table/tbody/tr[2]/td[1]/input[2]")
		private static WebElement TransFrom_DepartmnentTxt;

		@FindBy(xpath="//input[@id='optFromProduct']")
		private static WebElement TransFrom_itemTxt;
		
		@FindBy(xpath="//input[@id='optFromAccount']")
		private static WebElement TransFrom_accountTxt;
		
		@FindBy(xpath="//input[@id='txtFromQuantity']")
		private static WebElement TransFrom_quantityTxt;
		
		@FindBy(xpath="//input[@id='txtFromSpentQuantity']")
		private static WebElement TransFrom_spentQtyTxt;
		
		@FindBy(xpath="//input[@id='txtFromPreCommittedQuantity']")
		private static WebElement TransFrom_preCommitedQtyTxt;
		
		@FindBy(xpath="//input[@id='txtFromCommittedQuantity']")
		private static WebElement TransFrom_committedQtyTxt;
		
		@FindBy(xpath="//input[@id='txtFromBalanceQuantity']")
		private static WebElement TransFrom_balQtyTxt;
		
		@FindBy(xpath="//input[@id='txtFromBudgetValue']")
		private static WebElement TransFrom_budgetValueTxt;
		
		@FindBy(xpath="//input[@id='txtFromSpentBudget']")
		private static WebElement TransFrom_spentbudgetTxt;
		
		@FindBy(xpath="//input[@id='txtFromPreCommittedValue']")
		private static WebElement TransFrom_preCommittedTxt;
		
		@FindBy(xpath="//input[@id='txtFromCommittedBudget']")
		private static WebElement TransFrom_committedBudget ;
		
		@FindBy(xpath="//input[@id='txtFromBalanceBudget']")
		private static WebElement TransFrom_balBudgetTxt;
		
		@FindBy(xpath="//input[@id='txtFromReduceQuantity']")
		private static WebElement TransFrom_reduceQtyTxt;
		
		@FindBy(xpath="//input[@id='txtFromReduceValue']")
		private static WebElement TransFrom_reduceValueTxt;
		
		
		
		   // TRANSFER TO
		
		
			@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div[1]/div/div[2]/div[5]/div[2]/fieldset/div/div[2]/table/tbody/tr[2]/td[1]/input[2]")
			private static WebElement TransTo_DepartmentTxt;
			
			@FindBy(xpath="//input[@id='optToProduct']")
			private static WebElement TransTo_itemTxt;
			
			@FindBy(xpath="//input[@id='optToAccount']")
			private static WebElement TransTo_accountTxt;
			
			@FindBy(xpath="//input[@id='txtToQuantity']")
			private static WebElement TransTo_quantityTxt;
			
			@FindBy(xpath="//input[@id='txtToSpentQuantity']")
			private static WebElement TransTo_spentQtyTxt;
			
			@FindBy(xpath="//input[@id='txtToPreCommittedQuantity']")
			private static WebElement TransTo_preCommitedQtyTxt;
			
			@FindBy(xpath="//input[@id='txtToCommittedQuantity']")
			private static WebElement TransTo_committedQtyTxt;
			
			@FindBy(xpath="//input[@id='txtToBalanceQuantity']")
			private static WebElement TransTo_balQtyTxt;
			
			@FindBy(xpath="//input[@id='txtToBudgetValue']")
			private static WebElement TransTo_budgetValueTxt;
			
			@FindBy(xpath="//input[@id='txtToSpentBudget']")
			private static WebElement TransTo_spentbudgetTxt;
			
			@FindBy(xpath="//input[@id='txtToPreCommittedValue']")
			private static WebElement TransTo_preCommittedTxt;
			
			@FindBy(xpath="//input[@id='txtToCommittedBudget']")
			private static WebElement TransTo_committedBudget ;
			
			@FindBy(xpath="//input[@id='txtToBalanceBudget']")
			private static WebElement TransTo_balBudgetTxt;
			
			@FindBy(xpath="//input[@id='txtToReduceQuantity']")
			private static WebElement TransTo_reduceQtyTxt;
			
			@FindBy(xpath="//input[@id='txtToReduceValue']")
			private static WebElement TransTo_reduceValueTxt;
	
		@FindBy(xpath="//span[@class='icon-clear icon-font7']")
		private static WebElement headerClearBtn ;
		
		@FindBy(xpath="//button[@id='btnTBSplitBugdet']")
		private static WebElement splitBtn;
		
		@FindBy(xpath="//button[@id='btnAddToGrid']")
		private static WebElement addToGridBtn;
		
		@FindBy(xpath="//button[@id='btnClear']")
		private static WebElement clearbtn;
	
		@FindBy(xpath="//*[@id='txtToAddQuantity']")
		private static WebElement toAddQty;
		
		@FindBy(xpath="//*[@id='txtToAddValue']")
		private static WebElement toAddValue;
		
		@FindBy(xpath="//*[@id='grdTransferBudget_body']/tr/td[3]")
		private static WebElement gridOfDepartment;
	
		@FindBy(xpath="//*[@id='grdTransferBudget_body']/tr/td[6]")
		private static WebElement gridOfItem;
	
		@FindBy(xpath="//*[@id='grdTransferBudget_body']/tr/td[7]")
		private static WebElement gridOfAccount;
		
		@FindBy(xpath="//*[@id='grdTransferBudget_body']/tr/td[8]")
		private static WebElement gridOfQty;
		
		@FindBy(xpath="//*[@id='grdTransferBudget_body']/tr/td[9]")
		private static WebElement gridOfValue;
		
		@FindBy(xpath="//*[@id='txtGridNarration']")
		private static WebElement gridOfNarration;
		
		@FindBy(xpath="//*[@id='ucCancel']/div/div[2]")
		private static WebElement transferBudgetCloseBtn;
		
		
		public static boolean checkReviseMenuTransferBudgets() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		 /*try
		 {*/
	   	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
		   financialsMenu.click();

		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetMenu));
		   financialsBudgetMenu.click();
			 
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetReviseBudget));
		   financialsBudgetReviseBudget.click();
	      		   
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transferBudgetMenu));
		   transferBudgetMenu.click();
			
		   Thread.sleep(3000);
		   
	       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetVersionTxt));
	       
		   boolean actBudgetVersionTxt              = budgetVersionTxt.isDisplayed();
		   boolean actReviseTxt                     = reviseTxt.isDisplayed();
		   boolean actTransFrom_itemTxt             = TransFrom_itemTxt.isDisplayed();
		   boolean actTransFrom_accountTxt          = TransFrom_accountTxt.isDisplayed();		 
		   boolean actTransFrom_quantityTxt         = TransFrom_quantityTxt.isDisplayed();
		   boolean actTransFrom_spentQtyTxt         = TransFrom_spentQtyTxt.isDisplayed();
		   boolean actTransFrom_preCommitedQtyTxt   = TransFrom_preCommitedQtyTxt.isDisplayed();
		   boolean actTransFrom_committedQtyTxt     = TransFrom_committedQtyTxt.isDisplayed(); 
		   boolean actTransFrom_balQtyTxt           = TransFrom_balQtyTxt.isDisplayed();
		   boolean actTransFrom_budgetValueTxt      = TransFrom_budgetValueTxt.isDisplayed(); 
		   boolean actTransFrom_spentbudgetTxt      = TransFrom_spentbudgetTxt.isDisplayed();
		   boolean actTransFrom_preCommittedTxt     = TransFrom_preCommittedTxt.isDisplayed(); 
		   boolean actTransFrom_committedBudget     = TransFrom_committedBudget.isDisplayed();
		   boolean actTransFrom_balBudgetTxt        = TransFrom_balBudgetTxt.isDisplayed(); 
		   boolean actTransFrom_reduceQtyTxt        = TransFrom_reduceQtyTxt.isDisplayed();
		   boolean actTransFrom_reduceValueTxt      = TransFrom_reduceValueTxt.isDisplayed(); 
		  
		   boolean actTransTo_itemTxt               = TransTo_itemTxt.isDisplayed();
		   boolean actTransTo_accountTxt            = TransTo_accountTxt.isDisplayed();		 
		   boolean actTransTo_quantityTxt           = TransTo_quantityTxt.isDisplayed();
		   boolean actTransTo_spentQtyTxt           = TransTo_spentQtyTxt.isDisplayed();
		   boolean actTransTo_preCommitedQtyTxt     = TransTo_preCommitedQtyTxt.isDisplayed();
		   boolean actTransTo_committedQtyTxt       = TransTo_committedQtyTxt.isDisplayed(); 
		   boolean actTransTo_balQtyTxt             = TransTo_balQtyTxt.isDisplayed();
		   boolean actTransTo_budgetValueTxt        = TransTo_budgetValueTxt.isDisplayed(); 
		   boolean actTransTo_spentbudgetTxt        = TransTo_spentbudgetTxt.isDisplayed();
		   boolean actTransTo_preCommittedTxt       = TransTo_preCommittedTxt.isDisplayed(); 
		   boolean actTransTo_committedBudget       = TransTo_committedBudget.isDisplayed();
		   boolean actTransTo_balBudgetTxt          = TransTo_balBudgetTxt.isDisplayed(); 
		   /*boolean actTransTo_reduceQtyTxt          = TransTo_reduceQtyTxt.isDisplayed();
		   boolean actTransTo_reduceValueTxt        = TransTo_reduceValueTxt.isDisplayed();*/ 		   
		   
		   boolean actHeaderClearBtn                = headerClearBtn.isDisplayed(); 
		   boolean actSplitBtn                      = splitBtn.isDisplayed();
		   boolean actAddToGridBtn                  = addToGridBtn.isDisplayed(); 
		   boolean actClearBtn                      = clearbtn.isDisplayed();
		    	
		   
		   boolean expBudgetVersionTxt              = true;
		   boolean expReviseTxt                     = true;
		   boolean expTransFrom_itemTxt             = true;
		   boolean expTransFrom_accountTxt          = true;		 
		   boolean expTransFrom_quantityTxt         = true;
		   boolean expTransFrom_spentQtyTxt         = true;
		   boolean expTransFrom_preCommitedQtyTxt   = true;
		   boolean expTransFrom_committedQtyTxt     = true; 
		   boolean expTransFrom_balQtyTxt           = true;
		   boolean expTransFrom_budgetValueTxt      = true; 
		   boolean expTransFrom_spentbudgetTxt      = true;
		   boolean expTransFrom_preCommittedTxt     = true; 
		   boolean expTransFrom_committedBudget     = true;
		   boolean expTransFrom_balBudgetTxt        = true; 
		   boolean expTransFrom_reduceQtyTxt        = true;
		   boolean expTransFrom_reduceValueTxt      = true; 
		  
		   boolean expTransTo_itemTxt               = true;
		   boolean expTransTo_accountTxt            = true;		 
		   boolean expTransTo_quantityTxt           = true;
		   boolean expTransTo_spentQtyTxt           = true;
		   boolean expTransTo_preCommitedQtyTxt     = true;
		   boolean expTransTo_committedQtyTxt       = true; 
		   boolean expTransTo_balQtyTxt             = true;
		   boolean expTransTo_budgetValueTxt        = true; 
		   boolean expTransTo_spentbudgetTxt        = true;
		   boolean expTransTo_preCommittedTxt       = true; 
		   boolean expTransTo_committedBudget       = true;
		   boolean expTransTo_balBudgetTxt          = true; 
		   boolean expTransTo_reduceQtyTxt          = true;
		   boolean expTransTo_reduceValueTxt        = true; 		   
		   
		   boolean expHeaderClearBtn                = true; 
		   boolean expSplitBtn                      = true;
		   boolean expAddToGridBtn                  = true; 
		   boolean expClearBtn                      = true;
		    
		   	   	
		   System.out.println("Budget Menu Define Budget         : " + actBudgetVersionTxt +            "  value expected  " + expBudgetVersionTxt);
		   System.out.println("Budget Menu Confirm Budget        : " + actReviseTxt +                   "  value expected  " + expReviseTxt);
		   System.out.println("Budget Menu Budget Budget         : " + actTransFrom_itemTxt +           "  value expected  " + expTransFrom_itemTxt);
		   System.out.println("Budget Menu Approval Budget       : " + actTransFrom_accountTxt +        "  value expected  " + expTransFrom_accountTxt);
		   System.out.println("Budget Menu Budget Report Budget  : " + actTransFrom_quantityTxt +       "  value expected  " + expTransFrom_quantityTxt);
		   System.out.println("Budget Menu Define Budget         : " + actTransFrom_spentQtyTxt +       "  value expected  " + expTransFrom_spentQtyTxt);
		   System.out.println("Budget Menu Confirm Budget        : " + actTransFrom_preCommitedQtyTxt + "  value expected  " + expTransFrom_preCommitedQtyTxt);
		   System.out.println("Budget Menu Revise Budget         : " + actTransFrom_committedQtyTxt +   "  value expected  " + expTransFrom_committedQtyTxt);
		   System.out.println("Budget Menu Rejected Budget       : " + actTransFrom_balQtyTxt +         "  value expected  " + expTransFrom_balQtyTxt);	
		   System.out.println("Budget Menu Confirm Budget        : " + actTransFrom_budgetValueTxt +    "  value expected  " + expTransFrom_budgetValueTxt);
		   System.out.println("Budget Menu Revise Budget         : " + actTransFrom_spentbudgetTxt +    "  value expected  " + expTransFrom_spentbudgetTxt);
		   System.out.println("Budget Menu Rejected Budget       : " + actTransFrom_preCommittedTxt +   "  value expected  " + expTransFrom_preCommittedTxt);		   
		   System.out.println("Budget Menu Confirm Budget        : " + actTransFrom_balBudgetTxt +      "  value expected  " + expTransFrom_balBudgetTxt);
		   System.out.println("Budget Menu Revise Budget         : " + actTransFrom_reduceQtyTxt +      "  value expected  " + expTransFrom_reduceQtyTxt);
		   System.out.println("Budget Menu Rejected Budget       : " + actTransFrom_reduceValueTxt +    "  value expected  " + expTransFrom_reduceValueTxt);	
		   
		   System.out.println("Budget Menu Define Budget         : " + actTransTo_itemTxt +            "  value expected  " + expTransTo_itemTxt);
		   System.out.println("Budget Menu Confirm Budget        : " + actTransTo_accountTxt +         "  value expected  " + expTransTo_accountTxt);
		   System.out.println("Budget Menu Budget Budget         : " + actTransTo_quantityTxt +        "  value expected  " + expTransTo_quantityTxt);
		   System.out.println("Budget Menu Approval Budget       : " + actTransTo_spentQtyTxt +        "  value expected  " + expTransTo_spentQtyTxt);
		   System.out.println("Budget Menu Budget Report Budget  : " + actTransTo_preCommitedQtyTxt +  "  value expected  " + expTransTo_preCommitedQtyTxt);
		   System.out.println("Budget Menu Define Budget         : " + actTransTo_committedQtyTxt +    "  value expected  " + expTransTo_committedQtyTxt);
		   System.out.println("Budget Menu Confirm Budget        : " + actTransTo_balQtyTxt +          "  value expected  " + expTransTo_budgetValueTxt);
		   System.out.println("Budget Menu Revise Budget         : " + actTransTo_budgetValueTxt +     "  value expected  " + expTransFrom_committedQtyTxt);
		   System.out.println("Budget Menu Rejected Budget       : " + actTransTo_spentbudgetTxt +     "  value expected  " + expTransTo_spentbudgetTxt);	
		   System.out.println("Budget Menu Confirm Budget        : " + actTransTo_preCommittedTxt +    "  value expected  " + expTransTo_preCommittedTxt);
		   System.out.println("Budget Menu Revise Budget         : " + actTransTo_committedBudget +    "  value expected  " + expTransTo_committedBudget);
		   System.out.println("Budget Menu Rejected Budget       : " + actTransTo_balBudgetTxt +       "  value expected  " + expTransTo_balBudgetTxt);		   
		   /*System.out.println("Budget Menu Revise Budget         : " + actTransTo_reduceQtyTxt +       "  value expected  " + expTransTo_reduceQtyTxt);
		   System.out.println("Budget Menu Rejected Budget       : " + actTransTo_reduceValueTxt +     "  value expected  " + expTransTo_reduceValueTxt);	
*/
		   System.out.println("Budget Menu Revise Budget         : " + actHeaderClearBtn +    "  value expected  " + expHeaderClearBtn);
		   System.out.println("Budget Menu Rejected Budget       : " + actSplitBtn +          "  value expected  " + expSplitBtn);		   
		   System.out.println("Budget Menu Revise Budget         : " + actAddToGridBtn +      "  value expected  " + expAddToGridBtn);
		   System.out.println("Budget Menu Rejected Budget       : " + actClearBtn +          "  value expected  " + expClearBtn);	

		   
		   
		   if( actBudgetVersionTxt==expBudgetVersionTxt && actReviseTxt==expReviseTxt 
					&& actHeaderClearBtn==expHeaderClearBtn  && actSplitBtn==expSplitBtn 
					&& actAddToGridBtn==expAddToGridBtn && actClearBtn==expClearBtn)				
			{
				excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
				return true;
			}
			else
			{
				excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
		   String exception =e.getMessage().substring(0, 50);
	       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	       return false;
	    }
	}*/
		

		
		
		
		public static boolean checkTransferBudgerFromOptionsValues() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		 /*try
		 { */    	
			 
		   boolean actHeaderClearBtn                = headerClearBtn.isDisplayed(); 
		   boolean actSplitBtn                      = splitBtn.isDisplayed();
		   boolean actAddToGridBtn                  = addToGridBtn.isDisplayed(); 
		   boolean actClearBtn                      = clearbtn.isDisplayed(); 	
			  			   
			   boolean expHeaderClearBtn                = true; 
			   boolean expSplitBtn                      = true;
			   boolean expAddToGridBtn                  = true; 
			   boolean expClearBtn                      = true;
			 
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));     	
	 	   planNameTxt.sendKeys("ConfirmPlanBudget");
	 	   Thread.sleep(2000);
	 	   planNameTxt.sendKeys(Keys.TAB,Keys.TAB);
				   
	       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetVersionTxt));
	       
		   String actBudgetVersionTxt              = budgetVersionTxt.getText();
		   String actReviseTxt                     = reviseTxt.getText();
		   
		   String expBudgetVersionTxt              = "1";
		   String expReviseTxt                     = "3";
		   
		   System.out.println("Budget Menu Define Budget     : " + actBudgetVersionTxt +  "  value expected  " + expBudgetVersionTxt);
		   System.out.println("Budget Menu Confirm Budget    : " + actReviseTxt +         "  value expected  " + expReviseTxt);
		    
		   TransFrom_DepartmnentTxt.click();
		   TransFrom_DepartmnentTxt.sendKeys("DUBAI");
		   Thread.sleep(2000);
		   TransFrom_DepartmnentTxt.sendKeys(Keys.TAB);
		   TransFrom_itemTxt.sendKeys("BR COGS ITEM");
		   Thread.sleep(2000);
		   TransFrom_itemTxt.sendKeys(Keys.TAB);
		   TransFrom_accountTxt.sendKeys("BR COGS ACC INV"); 
		   Thread.sleep(2000);
		   TransFrom_accountTxt.sendKeys(Keys.TAB);
		  
		   Thread.sleep(2000);
		   
		   String actGetQty        = TransFrom_quantityTxt.getAttribute("value");
		   
		   String actBalanceQty    = TransFrom_balQtyTxt.getAttribute("value");
		   
		   String actBalanceValue  = TransFrom_budgetValueTxt.getAttribute("value");
		   
		   String actBudgetBalance = TransFrom_balBudgetTxt.getAttribute("value");
		   
		   String expGetQty        = "6.00";
		   
		   String expBalanceQty    = "6.00";
		 
		   String expBalanceValue  = "60.00";
		   
		   String expBudgetBalance = "60.00";
	   	   	   	
		   System.out.println("Budget Menu Define Budget     : " + actGetQty +         "  value expected  " + expGetQty);
		   System.out.println("Budget Menu Confirm Budget    : " + actBalanceQty +     "  value expected  " + expBalanceQty);
		   System.out.println("Budget Menu Budget Budget     : " + actBalanceValue +   "  value expected  " + expBalanceValue);
		   System.out.println("Budget Menu Approval Budget   : " + actBudgetBalance +  "  value expected  " + expBudgetBalance);
		   
		   
		   
		   TransFrom_reduceQtyTxt.sendKeys("2");
		   TransFrom_reduceQtyTxt.sendKeys(Keys.TAB);
		   TransFrom_reduceValueTxt.sendKeys("20");
		   TransFrom_reduceValueTxt.sendKeys(Keys.TAB);
		   
		   TransTo_DepartmentTxt.click();
		   TransTo_DepartmentTxt.sendKeys("DUBAI");
		   Thread.sleep(2000);
		   TransTo_DepartmentTxt.sendKeys(Keys.TAB);
		   TransTo_itemTxt.sendKeys("STD RATE COGS ITEM");
		   Thread.sleep(2000);
		   TransTo_itemTxt.sendKeys(Keys.TAB);
		   TransTo_accountTxt.sendKeys("STD RATE COGS ACC INV"); 
		   Thread.sleep(2000);
		   TransTo_accountTxt.sendKeys(Keys.TAB);
		   
		   
           String actTOGetQty        = TransTo_quantityTxt.getAttribute("value");
		   
		   String actTOBalanceQty    = TransTo_balQtyTxt.getAttribute("value");
		   
		   String actTOBalanceValue  = TransTo_budgetValueTxt.getAttribute("value");
		   
		   String actTOBudgetBalance = TransTo_balBudgetTxt.getAttribute("value");
		   
		   String expTOGetQty        = "30.00";
		   
		   String expTOBalanceQty    = "32.00";
		 
		   String expTOBalanceValue  = "360.00";
		   
		   String expTOBudgetBalance = "380.00";
	   	   	   	
		   System.out.println("Budget Menu Define Budget     : " + actTOGetQty +         "  value expected  " + expTOGetQty);
		   System.out.println("Budget Menu Confirm Budget    : " + actTOBalanceQty +     "  value expected  " + expTOBalanceQty);
		   System.out.println("Budget Menu Budget Budget     : " + actTOBalanceValue +   "  value expected  " + expTOBalanceValue);
		   System.out.println("Budget Menu Approval Budget   : " + actTOBudgetBalance +  "  value expected  " + expTOBudgetBalance);
		   
		   String actAddToQty     = toAddQty.getAttribute("value");
		   String actAddToValue   = toAddValue.getAttribute("value");
		   
		   String expAddToQty    = "2.00";
		   String expAddToValue  = "20.00";
		   
		   System.out.println("Budget Menu Define Budget     : " + actAddToQty +       "  value expected  " + expAddToQty);
		   System.out.println("Budget Menu Confirm Budget    : " + actAddToValue +     "  value expected  "   + expAddToValue);
		   
		   
		   
		   if(true)				
			{
				excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
				return true;
			}
			else
			{
				excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
		   String exception =e.getMessage().substring(0, 50);
	       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	       return false;
	    }
	}*/
		
		
		
		public static boolean checkAddToGridOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		/* try
		 { */    	
			 
		   System.out.println(toAddQty.getAttribute("value"));
		   
		   
	
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addToGridBtn));  
		   addToGridBtn.click();
		   	   
		   Thread.sleep(2000);
		   
		   
		   System.out.println(gridOfDepartment.getText());
		   System.out.println(gridOfItem.getText());
		   System.out.println(gridOfAccount.getText());
		   System.out.println(gridOfQty.getText());
		   System.out.println(gridOfValue.getText());
		   
           String actDepartment        = gridOfDepartment.getAttribute("value");
		   
		   String actItem              = gridOfItem.getAttribute("value");
		   
		   String actAccount           = gridOfAccount.getAttribute("value");
		   
		   String actQty               = gridOfQty.getAttribute("value");
		   
		   String actValue             = gridOfValue.getAttribute("value");
		   
		   String expDepartment        = "DUBAI to DUBAI";
		   
		   String expItem              = "BR COGS ITEM to STD RATE COGS ITEM";
	
		   String expAccount           = "BR COGS ACC INV to STD RATE COGS ACC INV";
		   
		   String expQty               = "2.00";
		   
		   String expValue             = "20.00";
		   
	   	   	   	
		   System.out.println("Budget Menu Define Budget     : " + actDepartment +     "  value expected  " + expDepartment);
		   System.out.println("Budget Menu Confirm Budget    : " + actItem +           "  value expected  " + expItem);
		   System.out.println("Budget Menu Budget Budget     : " + actAccount +        "  value expected  " + expAccount);
		   System.out.println("Budget Menu Approval Budget   : " + actQty +            "  value expected  " + expQty);
		   System.out.println("Budget Menu Approval Budget   : " + actValue +          "  value expected  " + expValue);
			   
		   
		   
		   budgetsTransferSaveBtn.click();
		   
		   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
	 		String actErrorMessage=errorMessage.getText();
	 		String expErrorMessage="Data Saved Successfully.";

	 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
	 		errorMessageCloseBtn.click();
		 		
	 		System.out.println("ValidationMessage  :  " + actErrorMessage +" Value Expected : " + expErrorMessage);

		   
		   if(true)				
			{
				excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
				return true;
			}
			else
			{
				excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
				return false;
			}
		}
		/*catch (Exception e) 
	    {
	       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
	       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
	       return false;
	    }
	}*/
		
		
		public boolean checkCloseButtonInBudgetTransfer() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
				excelReader=new ExcelReader(POJOUtility.getExcelPath());
				xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transferBudgetCloseBtn));
				transferBudgetCloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
				
				boolean actdashboard=dashboard.isDisplayed();
				boolean expdashboard=true;
					
				System.out.println("dashboard  :"+actdashboard +"  Value Expected  : "+expdashboard);
				
				if(actdashboard==expdashboard)	
				{
					System.out.println("Test Pass : Cancel Option Working ");
					excelReader.setCellData(xlfile, xlSheetName, 53, 9, resPass);
					return true;
				}
				else
				{
					System.out.println("Test Fail : Cancel Option Not Working");
					excelReader.setCellData(xlfile, xlSheetName, 53, 9, resFail);
					return false;
				}
		}
		
		
		
		
		// BUDGET AUTHORIZATION 
		

		//Setting Master Authorization	[Master Authorization---MasterAuthor]
			
		@FindBy(xpath="//*[@id='id_Authorization_container']/div/div/nav/div[1]/a/div/span[2]")
		private static WebElement masterAuthorLabel; 
			
		@FindBy(xpath="//input[@id='txtDoc']")
		private static WebElement budgetDropdown; 	
		
			@FindBy(xpath="//li[contains(text(),'Account')]")
			private static WebElement account; 	
				
			@FindBy(xpath="//li[contains(text(),'Item')]")
			private static WebElement item; 	
				
			@FindBy(xpath="//li[contains(text(),'Department')]")
			private static WebElement department; 	
				
			@FindBy(xpath="//li[contains(text(),'Warehouse')]")
			private static WebElement warehouse; 	
				
			@FindBy(xpath="//li[contains(text(),'Cost Center')]")
			private static WebElement costCenter; 	
				
			@FindBy(xpath="//li[contains(text(),'Location')]")
			private static WebElement location; 	
				
			@FindBy(xpath="//li[contains(text(),'Region')]")
			private static WebElement region; 	
				
			@FindBy(xpath="//li[contains(text(),'Country')]")
			private static WebElement country; 	
				
			@FindBy(xpath="//li[contains(text(),'State')]")
			private static WebElement state; 	
				
			@FindBy(xpath="//li[contains(text(),'City')]")
			private static WebElement city; 	
				
			@FindBy(xpath="//li[contains(text(),'Units')]")
			private static WebElement units; 	
				
			@FindBy(xpath="//li[contains(text(),'Bins')]")
			private static WebElement bins; 	
				
			@FindBy(xpath="//li[contains(text(),'Tax Code')]")
			private static WebElement taxCode; 	
				
			@FindBy(xpath="//li[contains(text(),'Plant')]")
			private static WebElement plant; 	
				
			@FindBy(xpath="//li[contains(text(),'Supply Area')]")
			private static WebElement supplyArea; 	
				
			@FindBy(xpath="//li[contains(text(),'Maintenance Parameter')]")
			private static WebElement maintenanceParameter; 	
				
			@FindBy(xpath="//li[contains(text(),'Safety Instructions')]")
			private static WebElement safetyInstructions; 	
				
			@FindBy(xpath="//li[contains(text(),'Work Center')]")
			private static WebElement workCenter; 	
				
			@FindBy(xpath="//li[contains(text(),'Capacity')]")
			private static WebElement capacity; 	
				
			@FindBy(xpath="//li[contains(text(),'Holiday')]")
			private static WebElement holiday; 	
				
			@FindBy(xpath="//li[contains(text(),'Process')]")
			private static WebElement process; 	
				
			@FindBy(xpath="//li[contains(text(),'Qc failure reason')]")
			private static WebElement qcFailureReason; 	
				
			@FindBy(xpath="//li[contains(text(),'QC Parameters')]")
			private static WebElement qcParameters; 	
				
			@FindBy(xpath="//li[contains(text(),'Insurance')]")
			private static WebElement insurance; 	
				
			@FindBy(xpath="//li[contains(text(),'Fixed Asset')]")
			private static WebElement fixedAsset; 	
				
			@FindBy(xpath="//li[contains(text(),'Break down standard reason')]")
			private static WebElement breakDownStandardReason; 	
				
			@FindBy(xpath="//*[@id='MasterMenu']/li[27]")
			private static WebElement employee; 	
				
			@FindBy(xpath="//li[contains(text(),'Designation')]")
			private static WebElement designation; 	
				
			@FindBy(xpath="//li[contains(text(),'Position')]")
			private static WebElement position; 	
				
			@FindBy(xpath="//li[contains(text(),'Qualification')]")
			private static WebElement qualification; 	
				
			@FindBy(xpath="//li[contains(text(),'Specialization')]")
			private static WebElement specialization; 	
				
			@FindBy(xpath="//li[contains(text(),'Nationality')]")
			private static WebElement nationality; 	
				
			@FindBy(xpath="//li[contains(text(),'Skill Type')]")
			private static WebElement skillType; 	
				
			@FindBy(xpath="//*[@id='MasterMenu']/li[34]")
			private static WebElement skill; 	
				
			@FindBy(xpath="//li[contains(text(),'SourceType')]")
			private static WebElement sourceType; 	
				
			@FindBy(xpath="//*[@id='MasterMenu']/li[36]")
			private static WebElement source; 	
				
			@FindBy(xpath="//li[contains(text(),'RoundType')]")
			private static WebElement roundType; 	
				
			@FindBy(xpath="//*[@id='MasterMenu']/li[38]")
			private static WebElement grade; 	
				
			@FindBy(xpath="//li[contains(text(),'Scale')]")
			private static WebElement scale; 	
				
			@FindBy(xpath="//li[contains(text(),'Course Type')]")
			private static WebElement courseType; 	
				
			@FindBy(xpath="//*[@id='MasterMenu']/li[41]")
			private static WebElement course; 	
				
			@FindBy(xpath="//li[contains(text(),'Trainer')]")
			private static WebElement trainer; 	
				
			@FindBy(xpath="//li[contains(text(),'Airline Sector')]")
			private static WebElement airlineSector; 	
				
			@FindBy(xpath="//li[contains(text(),'Venue')]")
			private static WebElement venue; 	
				
			@FindBy(xpath="//li[contains(text(),'Request Types')]")
			private static WebElement requestTypes; 	
				
			@FindBy(xpath="//li[contains(text(),'Expense Claims')]")
			private static WebElement expenseClaims; 	
				
			@FindBy(xpath="//li[contains(text(),'Employee Bank')]")
			private static WebElement employeeBank; 	
				
			@FindBy(xpath="//li[contains(text(),'Travel Agent')]")
			private static WebElement travelAgent; 	
				
			@FindBy(xpath="//li[contains(text(),'Job Grade')]")
			private static WebElement jobGrade; 	
				
			@FindBy(xpath="//li[contains(text(),'Outlet')]")
			private static WebElement outlet; 	
				
			@FindBy(xpath="//li[contains(text(),'Counter')]")
			private static WebElement counter; 	
				
			@FindBy(xpath="//li[contains(text(),'Member Type')]")
			private static WebElement memberType; 	
				
			@FindBy(xpath="//li[contains(text(),'Gift Voucher Definition')]")
			private static WebElement giftVoucherDefinition; 	
				
			@FindBy(xpath="//li[contains(text(),'Category')]")
			private static WebElement category; 	
				
			@FindBy(xpath="//li[contains(text(),'Bank Card Type')]")
			private static WebElement bankCardType; 	
				
			@FindBy(xpath="//*[@id='MasterMenu']/li[56]")
			private static WebElement member; 	
				
			@FindBy(xpath="//li[contains(text(),'Discount Voucher Definition')]")
			private static WebElement discountVoucherDefinition; 	
				
			@FindBy(xpath="//li[contains(text(),'Floor')]")
			private static WebElement floor; 	
				
			@FindBy(xpath="//li[contains(text(),'Section')]")
			private static WebElement section; 	
				
			@FindBy(xpath="//li[contains(text(),'Table')]")
			private static WebElement table; 	
				
			@FindBy(xpath="//li[contains(text(),'Guest')]")
			private static WebElement guest; 	
				
			@FindBy(xpath="//li[contains(text(),'Void Remarks')]")
			private static WebElement voidRemarks; 	
				
			@FindBy(xpath="//li[contains(text(),'Member Card Definition')]")
			private static WebElement memberCardDefinition; 	
				
			@FindBy(xpath="//li[contains(text(),'Return Remarks')]")
			private static WebElement returnRemarks; 	
				
			@FindBy(xpath="//li[contains(text(),'Kitchen Display System')]")
			private static WebElement kitchenDisplaySystem; 	
				
			@FindBy(xpath="//li[contains(text(),'Delivery Time Interval')]")
			private static WebElement deliveryTimeInterval; 	
				
			@FindBy(xpath="//li[contains(text(),'E- Payment')]")
			private static WebElement ePayment; 	

			
		//Name Fields		
		@FindBy(xpath="//*[@id='ctrlAuthorizationH']")
		private static WebElement nameTxt; 	
			
		@FindBy(xpath="//*[@id='ctrlAuthorizationH_input_settings']/span")
		private static WebElement nameSettingBtn; 	

		    @FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_heading']")
		    private static WebElement nameSettingPop_CustomizeDisplayColoumnLabel;
		
		    @FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_container']")
		    private static WebElement nameSettingPop_sAuthorizationTxt;
		
		    @FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_footer']/div/div/input[1]")
		    private static WebElement nameSettingPop_StandardFieldsBtn;
		
				@FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_standardfields_list']")
			    private static WebElement nameSettingPop_StandardFields_FiledDropdown;
			
				@FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_standardfields_header']")
			    private static WebElement nameSettingPop_StandardFields_HeaderTxt;
			
				@FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_standardfields_alignment']")
			    private static WebElement nameSettingPop_StandardFields_AlignmentDropdown;
			
				@FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_standardfields_width']")
			    private static WebElement nameSettingPop_StandardFields_WidthTxt;
			
				@FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_footer']/div/div/input[3]")
				private static WebElement nameSettingPop_StandardFields_OkBtn;
				
				@FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_footer']/div/div/input[4]")
				private static WebElement nameSettingPop_StandardFields_CancelBtn;
		
		    @FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_footer']/div/div/input[2]")
		    private static WebElement nameSettingPop_DeleteColoumnBtn;
		
		    @FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_footer']/div/div/input[3]")
		    private static WebElement nameSettingPop_OkBtn;
		
		    @FindBy(xpath="//*[@id='ctrlAuthorizationH_customize_popup_footer']/div/div/input[4]")
		    private static WebElement nameSettingPop_CancelBtn;
		
		
		@FindBy(xpath="//*[@id='chkActive']")
		private static WebElement activeChkBox;  
		
		@FindBy(xpath="//*[@id='chkAlwaysauthorizationstartsfromfirstlevel']")
		private static WebElement alwaysauthorizationstartsfromfirstlevelChkBox; 
			
		
		//Master Authorization Clear,Delete,Save,Cancel
		@FindBy(xpath="//span[@class='icon-clear icon-font7']")
		private static WebElement clearbtnInAuth; 
		
		@FindBy(xpath="//span[@class='icon-delete icon-font7']")
		private static WebElement deleteBtn; 
		
		@FindBy(xpath="//span[@class='icon-save icon-font7']")
		private static WebElement saveBtn; 
		
		@FindBy(xpath="//span[@class='icon-close icon-font7']")
		private static WebElement cancelBtnInAuth; 
			
		
		//Action On Initial Rejection
		@FindBy(xpath="//*[@id='ddlAuthStatus']")
		private static WebElement statusDropdown; 	
		
		@FindBy(xpath="//*[@id='chkSuspendCreditLimit']")
		private static WebElement suspendedCreditLimitChkBox;
		
		@FindBy(xpath="//*[@id='spnAddLevel']/i")
		private static WebElement addLevelBtn; 
		
		@FindBy(xpath="//*[@id='spnAddParallel']/i")
		private static WebElement addParallelBtn;    
		
		
		@FindBy(xpath="//i[@id='id_StartNode']")
		private static WebElement startPointBtn; 
		
		//After Clicking On AddLevel/Add Parallel Button The Authorization Definition POp Rises--Related Fields 
		@FindBy(xpath="//*[@id='id_report_popup_heading']")
		private static WebElement authorizationDefinitionLabel; 
			
		@FindBy(xpath="//span[contains(text(),'Condition')]")
		private static WebElement conditionTab; 
			
		//Description Fields
		@FindBy(xpath="//*[@id='txtDescription']")
		private static WebElement descriptionTxt; 
			
		@FindBy(xpath="//*[@id='chkConditionNotRequired']")
		private static WebElement conditionNotRequiredChkBox; 
		
		@FindBy(xpath="//*[@id='chkCreditLimitExceeded']")
		private static WebElement creditLimitExceededChkBox; 
			
		@FindBy(xpath="//*[@id='chkBudgetExceeded']")
		private static WebElement budgetExceededChkBox; 
			
		@FindBy(xpath="//*[@id='chkNegativeCashCheck']")
		private static WebElement negativeCashCheckChkBox; 
			
		@FindBy(xpath="//*[@id='chkLineWiseAuthorization']")
		private static WebElement lineWiseAuthorizationChkBox; 
			
		@FindBy(xpath="//*[@id='chkCreditDaysExceeded']")
		private static WebElement creditDaysExceededChkBox; 
			
		@FindBy(xpath="//*[@id='chkNegativeStockCheck']")
		private static WebElement negativeStockCheckChkBox; 
			
		@FindBy(xpath="//*[@id='chkDateRange']")
		private static WebElement dateRangeChkBox; 
			
		@FindBy(xpath="//*[@id='dpFromDate']")
		private static WebElement fromDateTxt; 
			
		@FindBy(xpath="//*[@id='dpToDate']")
		private static WebElement toDateTxt; 
		
		@FindBy(xpath="//*[@id='chkTimeRange']")
		private static WebElement timeRangeChkBox;
			
		@FindBy(xpath="//*[@id='chkTimeRange']")
		private static WebElement fromTimeTxt; 
			
		@FindBy(xpath="//*[@id='tpToTime']")
		private static WebElement totimeTxt; 
		
		
		//Condition Condition Fields
		@FindBy(xpath="//body[@class='custom_scrollbar']/section[@id='mainDiv']/div[@id='id_focus8_wrapper_default']/div[@class='content-wrapper']/section[@id='page_Content']/div[@id='div_budgetPlanAuthorization']/div[@id='div_bugetPlanAuthorizationBody']/div[@id='div_bodyDisplay']/div[@id='div_authbody']/div[@id='id_report_popup_body']/div[@id='AllContentsDiv']/div[@id='tab_Condition']/div[@class='divConditionFormula']/div[@class='col-lg-11 col-md-11 col-sm-12 col-xs-12']/div/div[@id='divCondition']/div[@id='3308_0_AdvanceFilter_']/table/tbody/tr/td[1]/select[1]")
		private static WebElement conjunctionDropdown; 
			
		@FindBy(xpath="//tbody//input[@placeholder='Select Field']")
		private static WebElement selectFieldTxt; 

	    //Selection Field  Inner Fields
		
		@FindBy(xpath="//div[@id='filterTree_3308_0_AdvanceFilter_']//a[@id='1']")
		private static WebElement seletFieldAccount;
		
		@FindBy(xpath="//div[@id='filterTree_3308_0_AdvanceFilter_']//a[@id='2']")
		private static WebElement seletFieldItem;
		
		@FindBy(xpath="//div[@id='filterTree_3308_0_AdvanceFilter_']//a[@id='-1']")
		private static WebElement seletFieldAmount;
		
		
		@FindBy(xpath="//*[@id='5002']")
		private static WebElement sName;
		
		@FindBy(xpath="//*[@id='5003']")
		private static WebElement sCode;
		
		@FindBy(xpath="//*[@id='5004']")
		private static WebElement iAccountType;
		
		@FindBy(xpath="//*[@id='5005']")
		private static WebElement fCreditLimit;
		
		@FindBy(xpath="//*[@id='5006']")
		private static WebElement iCreditDays;
		
		@FindBy(xpath="//*[@id='300012']")
		private static WebElement iDebitCreditProposal;
		
		@FindBy(xpath="//*[@id='300013']")
		private static WebElement iDebitCreditRequired;
		
		@FindBy(xpath="//*[@id='300014']")
		private static WebElement iExchangeAdjustmentGainAC;
		
		@FindBy(xpath="//*[@id='300015']")
		private static WebElement iExchangeAdjustmentLossAC;
		
		@FindBy(xpath="//*[@id='300016']")
		private static WebElement iPrimaryAccount;
		
		@FindBy(xpath="//*[@id='300018']")
		private static WebElement iConsolidationMethod;
		
		@FindBy(xpath="//*[@id='300023']")
		private static WebElement sAddress;
		
		@FindBy(xpath="//*[@id='300024']")
		private static WebElement sDeliveryAddress;
		
		@FindBy(xpath="//*[@id='300027']")
		private static WebElement sPin;
		
		@FindBy(xpath="//*[@id='300028']")
		private static WebElement sDeliveryPin;
		
		@FindBy(xpath="//*[@id='300083']")
		private static WebElement iCity;
		
		@FindBy(xpath="//*[@id='300084']")
		private static WebElement iDeliveryCity;
		
		@FindBy(xpath="//*[@id='300249']")
		private static WebElement sEmail;
		
		@FindBy(xpath="//*[@id='300250']")
		private static WebElement bSendEmailtocustomer;
		
		@FindBy(xpath="//*[@id='302774']")
		private static WebElement fChequeDiscountLimit;
		
		@FindBy(xpath="//*[@id='302870']")
		private static WebElement fRateofinterest;
		
		@FindBy(xpath="//*[@id='302968']")
		private static WebElement iBankAC;
		
		@FindBy(xpath="//*[@id='303573']")
		private static WebElement sTelNo;
		
		@FindBy(xpath="//*[@id='303574']")
		private static WebElement sFaxNo;
		
		@FindBy(xpath="//*[@id='303785']")
		private static WebElement sPassword;
		
		@FindBy(xpath="//*[@id='303786']")
		private static WebElement bAllowCustomerPortal;
		
		@FindBy(xpath="//*[@id='303935']")
		private static WebElement sBankAccountName;
		
		@FindBy(xpath="//*[@id='303936']")
		private static WebElement sBankAccountNumber;
		
		@FindBy(xpath="//*[@id='303937']")
		private static WebElement sIFSCCode;
		
		@FindBy(xpath="//*[@id='304047']")
		private static WebElement iPaymentType;
		
		@FindBy(xpath="//*[@id='304157']")
		private static WebElement iPDCDiscountedAC;
		
		@FindBy(xpath="//*[@id='304269']")
		private static WebElement sFinanceEmail;
		
		@FindBy(xpath="//*[@id='304379']")
		private static WebElement sPortalEmail;
			
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter_']/table/tbody/tr/td[3]/select")
		private static WebElement selectOperatorDropdown; 
			
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter_']/table/tbody/tr/td[4]/select")
		private static WebElement compareWithDropdown; 
			
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter_']/table/tbody/tr/td[5]/input")
		private static WebElement valueTxt;
		
		@FindBy(xpath="//input[@id='advancefilter_master_3308_0_']")
		private static WebElement enter_valueTxt;
		
		@FindBy(xpath="//select[@id='advancefilter_SelectOptionControl_3308_0']")
		private static WebElement valueDropdown;
			
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter_']/table/tbody/tr/td[6]/span")
		private static WebElement deleteRow1Btn; 
			
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter_']/table/tbody/tr/td[7]/span")
		private static WebElement addRowBtn; 
			
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter']/table/tbody/tr[2]/td[6]/span")
		private static WebElement deleteSecondRowBtn; 
		
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter_']/table/tbody/tr[2]/td[1]/select")
		private static WebElement conjunctionDropdown2; 
		
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter_']/table/tbody/tr[2]/td[2]")
		private static WebElement selectFieldTxt2; 
		
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter_']/table/tbody/tr[2]/td[3]/select")
		private static WebElement selectOperatorDropdown2; 
		
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter_']/table/tbody/tr[2]/td[4]/select")
		private static WebElement comparewithDropdown2; 
		
		@FindBy(xpath="//*[@id='3308_0_AdvanceFilter_']/table/tbody/tr[2]/td[5]/input")
		private static WebElement valueTxt2; 
		
		@FindBy(xpath="//*[@id='advancefilter_SelectOptionControl_67_0']")
		private static WebElement valueDropdown2; 
		
		//On Entry
		@FindBy(xpath="//select[@id='ddlStatus']")
		private static WebElement actionOnEntryStatusDropdown; 	
			
		//Authorization Mode
		@FindBy(xpath="//*[@id='chkFullAuthorization']")
		private static WebElement fullAuthorizationChkBox; 	
		
		//Undo AuthorizationOn Editing
		@FindBy(xpath="//*[@id='ddlUndoAuthEditing']")
		private static WebElement undoAuthEditingDropdown; 	
		
		
		//Master Authorization Add Field User Selection
		@FindBy(xpath="//*[@id='id_report_popup_body']/div[2]/div/div[1]/ul/li[2]/a/span")
		private static WebElement userSelectionTab;
		
		@FindBy(xpath="//*[@id='ddlUserSelection']")
		private static WebElement userselectionDropdown;
		
		@FindBy(xpath="//*[@id='ddlCriteria']")
		private static WebElement criteriaDropdown;
		
		// Grid
		@FindBy(xpath="//*[@id='id_tblUserSelection_User']")
		private static WebElement userdropdown; 	
		
		@FindBy(xpath="//*[@id='id_tblUserSelection_col_1-1']")
		private static WebElement userrow1;
			
		@FindBy(xpath="//*[@id='id_tblUserSelection_col_2-1']")
		private static WebElement userrow2; 
		
		//Alert Tab Fields
		@FindBy(xpath="//*[@id='id_report_popup_body']/div[2]/div/div[1]/ul/li[3]/a/span")
		private static WebElement alertsTab;	
			
		//Grid
		@FindBy(xpath="//*[@id='chkReceived']")
		private static WebElement receivedChkBox;	
			
		@FindBy(xpath="//*[@id='chkApproved']")
		private static WebElement approvedChkBox;	
			
		@FindBy(xpath="//*[@id='chkRejected']")
		private static WebElement rejectedChkBox;	
			
		@FindBy(xpath="//*[@id='chkStopped']")
		private static WebElement stoppedChkBox;	
		
		//EmailHeaderTemplate
		@FindBy(xpath="//*[@id='id_tblAlerts_control_heading_2']/div/u")
		private static WebElement emailHeaderTemplateBtn;	
		
			@FindBy(xpath="//*[@id='id_SMSTemplate']/div/nav/div/div[1]/a/div/span[2]")
			private static WebElement emailHeaderTemplate_CustomizeTemplateLabel;	
			
		//Header Fields Of Customize Template Pop
		@FindBy(xpath="//*[@id='btnNewSmsTemplate']")
		private static WebElement customizeTemplateNewBtn;	
		
		@FindBy(xpath="//*[@id='btnSaveSmsTemplate']")
		private static WebElement customizeTemplateSaveBtn;	
			
		@FindBy(xpath="//*[@id='btnCloseSMSModal']")
		private static WebElement customizeTemplateCloseBtn;	
			
		@FindBy(xpath="//*[@id='smsTemplateOptCtrl']")
		private static WebElement customizeTemplateTemplatestxt;	
			
		@FindBy(xpath="//*[@id='smsTemplateOptCtrl_input_settings']/span")
		private static WebElement customizeTemplate_TemplatesSettingBtn;	
		
		
			@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_heading']")
			private static WebElement templateSettingPop_CustomizeDisplayColoumnLabel;
			         
			@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_container']")
			private static WebElement templateSettingPop_sTemplateNameTxt;

			@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_footer']/div/div/input[1]")
			private static WebElement templateSettingPop_StandardFieldsBtn;

				@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_standardfields_list']")
			    private static WebElement templateSettingPop_StandardFields_FiledDropdown;
			
				@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_standardfields_header']")
			    private static WebElement templateSettingPop_StandardFields_HeaderTxt;
			
				@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_standardfields_alignment']")
			    private static WebElement templateSettingPop_StandardFields_AlignmentDropdown;
			
				@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_standardfields_width']")
			    private static WebElement templateSettingPop_StandardFields_WidthTxt;
			
				@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_footer']/div/div/input[3]")
				private static WebElement templateSettingPop_StandardFields_OkBtn;
				
				@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_footer']/div/div/input[4]")
				private static WebElement templateSettingPop_StandardFields_CancelBtn;

			@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_footer']/div/div/input[2]")
			private static WebElement templateSettingPop_DeleteColoumnBtn;

			@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_footer']/div/div/input[3]")
			private static WebElement templateSettingPop_OkBtn;

			@FindBy(xpath="//*[@id='smsTemplateOptCtrl_customize_popup_footer']/div/div/input[4]")
			private static WebElement templateSettingPop_CancelBtn;
			
		//Email Template Tab			
		//Header Fields
		@FindBy(xpath="//u[contains(text(),'Email Template')]")
		private static WebElement emailTemplateBtn;				
					
		@FindBy(xpath="//*[@id='forHf']/span[1]/span/i[1]")
		private static WebElement emailTemplateNewBtn;	
		
		@FindBy(xpath="//i[@class='icon-font6 icon-new']")
		private static WebElement  newdropdown;
		
		@FindBy(xpath="//*[@id='forHf']/span[1]/ul/li/a[1]")
		private static WebElement  new_LayoutBtn;
		
		@FindBy(xpath="//*[@id='forHf']/span[1]/ul/li/a[2]")
		private static WebElement  new_PageBtn;
		
		@FindBy(xpath="//i[@class='icon-font6 icon-open']")
		private static WebElement  opendropdown;
		
		@FindBy(xpath="//*[@id='forHf']/span[2]/ul/li/a[1]")
		private static WebElement  openBtn;

		   //Pop
		    @FindBy(xpath="//*[@id='1']")
		    private static WebElement  open_Selct1Option;
		
		    @FindBy(xpath="//button[@id='btn_PopUpOk']")
			private static WebElement  openBtn_OkBtn;
			
			@FindBy(xpath="//div[@id='id_PopUp']//button[@class='Fbutton'][contains(text(),'Cancel')]")
			private static WebElement  openBtn_CancelBtn;
		
		@FindBy(xpath="//*[@id='forHf']/span[2]/ul/li/a[2]")
		private static WebElement fromXMLBtn;
		
		@FindBy(xpath="//i[@class='icon-font6 icon-save']")
		private static WebElement  savedropdown;

			@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[1]")
			private static WebElement  emailPrintSaveBtn;
			
			@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[2]")
			private static WebElement  saveAsBtn; 
			
			@FindBy(xpath="//input[@id='id_PopuplayoutName']")
			private static WebElement  saveTxtField;
			
			@FindBy(xpath="//input[@id='id_PopuplayoutName']")
			private static WebElement  save_OkBtn;
			
			@FindBy(xpath="//*[@id='id_PopUp']/div/div/div[3]/button[2]")
			private static WebElement  save_CancekBtn;
			
			@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[3]")
			private static WebElement  saveAsImageBtn; //No Response
			
			@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[4]")
			private static WebElement  saveAsXMLBtn; // Downloads xml file
			
			@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[5]")
			private static WebElement  saveAsHTMLBtn; 

		@FindBy(xpath="//*[@id='forHf']/span[4]")
		private static WebElement  removedropdown;

			@FindBy(xpath="//*[@id='forHf']/span[4]/ul/li/a[1]")
			private static WebElement  removeBtn;  
			
			@FindBy(xpath="//*[@id='forHf']/span[4]/ul/li/a[2]")
			private static WebElement  removeCurrentPageBtn; 
	                          
			/*//popup Error MEssage 
			@FindBy(xpath="//*[@id='idGlobalError']/div/table/tbody/tr/td[2]/div[1]")
			private static WebElement  newerrorMessage;//This Is only Page.You Not Delete this Page --Message while using RemovecurrentPAgeBtn 
	*/		
			@FindBy(xpath="//*[@id='forHf']/span[4]/ul/li/a[3]")
			private static WebElement  removeControlBtn;				

		@FindBy(xpath="//*[@id='forHf']/span[5]")
		private static WebElement  pageNo;
		
		@FindBy(xpath="//*[@id='forHf']/span[6]")
		private static WebElement  controldrpdown;
			
			@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[1]")
			private static WebElement  staticTextBtn;
			
			@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[2]")
			private static WebElement  bitmapBtn;
			
			@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[3]")
			private static WebElement  bodyGridBtn;	
			
			@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[4]")
			private static WebElement  rectangleBtn;	
			
			@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[5]")
			private static WebElement  ecllipseBtn;	
			
			@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[6]")
			private static WebElement  lineBtn;	
			
			@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[7]")
			private static WebElement  areaBtn;
			
			@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[8]")
			private static WebElement  tableBtn;
			
			@FindBy(xpath="//*[@id='forHf']/span[6]/ul/li/a[9]")
			private static WebElement  approvalHistoryBtn;
		
		@FindBy(xpath="//*[@id='forHf']/span[7]/span/i[1]")
		private static WebElement  alignmentdropdown;
			
			@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[1]")
			private static WebElement  leftAlignmentBtn;
						
			@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[2]")
			private static WebElement  rightAlignmentBtn;
			
			@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[3]")
			private static WebElement  topAlignmentBtn;	
			
			@FindBy(xpath="//*[@id='forHf']/span[7]/ul/li/a[4]")
			private static WebElement  bottomAlignmentBtn;
		
		@FindBy(xpath="//*[@id='forHf']/span[8]")
		private static WebElement  sizedropdown;

			@FindBy(xpath="//*[@id='forHf']/span[8]/ul/li/a[1]")
			private static WebElement  sameSizeBtn;
						
			@FindBy(xpath="//*[@id='forHf']/span[8]/ul/li/a[2]")
			private static WebElement  sameWidthBtn;
			
			@FindBy(xpath="//*[@id='forHf']/span[8]/ul/li/a[3]")
			private static WebElement  sameHeightBtn;
		
		@FindBy(xpath="//*[@id='forHf']/span[9]")
		private static WebElement  printdropdown;
		
		@FindBy(xpath="//*[@id='menu-print']")
		private static WebElement  print_PrintBtn;
						
			@FindBy(xpath="//*[@id='menu-preview']	")
			private static WebElement  print_PreviewBtn;
			
			@FindBy(xpath="//*[@id='fit-button']")
			private static WebElement  preview_FitToWidthBtn;
			
			@FindBy(xpath="//*[@id='zoom-in-button']")
			private static WebElement  previewZ_ZoomInBtn;
			
			@FindBy(xpath="//*[@id='zoom-out-button']")
			private static WebElement  preview_ZoomOutBtn;
			
			@FindBy(xpath="//*[@id='forHf']/span[9]/ul/li/a[3]")
			private static WebElement  pageSetUpBtn;
			
			@FindBy(xpath="//*[@id='id_PageSetUpPreview']")
			private static WebElement  pageSetupPagePreviewField;
			
			@FindBy(xpath="//*[@id='id_PageType']")
			private static WebElement  pageSetupPaperSizedropdown;
			
			@FindBy(xpath="//*[@id='id_PaperOrientation'][@value='0']")
			private static WebElement  pageSetupPotraitRadio;
			
			@FindBy(xpath="//*[@id='id_PaperOrientation'][@value='1']")
			private static WebElement  pageSetupLandScapeRadio;
			
			@FindBy(xpath="//*[@id='id_PaperUnit'][@value='0']")
			private static WebElement  pageSetupCentimetersRadio;
			
			@FindBy(xpath="//*[@id='id_PaperUnit'][@value='1']")
			private static WebElement  pageSetupInchesRadio;
			
			@FindBy(xpath="//*[@id='id_PageSizeLeftMarginDisplay']")
			private static WebElement  pageSetupLeftTxt;
			
			@FindBy(xpath="//*[@id='id_PageSizeRightMarginDisplay']")
			private static WebElement  pageSetupRightTxt;
			
			@FindBy(xpath="//*[@id='id_PageSizeTopMarginDisplay']")
			private static WebElement  pageSetupTopTxt;
			
			@FindBy(xpath="//*[@id='id_PageSizeTopMarginDisplay']")
			private static WebElement  pageSetupBottomTxt;
			
			@FindBy(xpath="//*[@id='id_PageSetupOkBtn']")
			private static WebElement  pageSetupOkBtn;
			
			@FindBy(xpath="//*[@id='id_PageSetupCancelBtn']")
			private static WebElement  pageSetUpCancelBtn;
	                  
		@FindBy(xpath="//*[@id='id_InvoiceDesingCancel']/i")
		private static WebElement  exitBtn;

	    //Header Elements

		
		@FindBy(xpath="//*[@id='li_InvoicePageSetup']/a")
		private static WebElement  headerPageSetUpTab;
		
		@FindBy(xpath="//*[@id='li_InvoiceDesigner']/a")
		private static WebElement headerDesignerTab;
		
		@FindBy(xpath="//*[@id='li_InvoicePreview']/a")
		private static WebElement  headerPreviewTab;

		//Fields  
		
		@FindBy(xpath="//*[@id='LiDesignField']/a")
		private static WebElement  headerFieldsTab;
		
		@FindBy(xpath="//*[@id='LiProp']/a")
		private static WebElement  headerPropertiesTab;

		@FindBy(xpath="//span[contains(text(),'Company')]//i[@class='icon-expand icon-font8']")
		private static WebElement  fields_CompanyExpandBtn;

			@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Company Name')]")
			private static WebElement  fields_Company_CompanyName;
			
			@FindBy(xpath="//span[contains(text(),'Company Code')]")
			private static WebElement  fields_Company_CompanyCode;
			
			@FindBy(xpath="//span[contains(text(),'Accounting Date')]")
			private static WebElement  fields_Company_AccountingDate;
			
			@FindBy(xpath="//span[contains(text(),'Address1')]")
			private static WebElement  fields_Company_Address1;
			
			@FindBy(xpath="//span[contains(text(),'Address2')]")
			private static WebElement  fields_Company_Address2;
			
			@FindBy(xpath="//li[@class='treeview active']//span[contains(text(),'City')]")
			private static WebElement  fields_Company_City;
			
			@FindBy(xpath="//span[contains(text(),'Zip Code')]")
			private static WebElement  fields_Company_Zipcode;
			
			@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Country')]")
			private static WebElement  fields_Company_Country;
			
			@FindBy(xpath="//span[contains(text(),'Phone No1')]")
			private static WebElement  fields_Company_PhoneNumber1;
			
			@FindBy(xpath="//span[contains(text(),'Phone No2')]")
			private static WebElement  fields_Company_PhoneNumber2;
			
			@FindBy(xpath="//li[@class='treeview active']//span[contains(text(),'Fax')]")
			private static WebElement  fields_Company_Fax;
			
			@FindBy(xpath="//span[contains(text(),'Tax Registration Number')]")
			private static WebElement  fields_Company_TaxRegestrationNumber;
			
			@FindBy(xpath="//span[contains(text(),'Excise Registration Number')]")
			private static WebElement  fields_Company_ExciseRegistrationNumber;
			
		@FindBy(xpath="//span[contains(text(),'Master Fields')]//i[@class='icon-expand icon-font8']")
		private static WebElement fields_MasterFieldsExpandBtn; 	
		                       
			@FindBy(xpath="//span[contains(text(),'sName')]")
			private static WebElement fields_sName; 	
				
			@FindBy(xpath="//span[contains(text(),'sCode')]")
			private static WebElement fields_sCode; 	
				
			@FindBy(xpath="//span[contains(text(),'iAccountType')]")
			private static WebElement fields_iAccountType; 	
				
			@FindBy(xpath="//span[contains(text(),'fCreditLimit')]")
			private static WebElement fields_fCreditLimit; 	
				
			@FindBy(xpath="//span[contains(text(),'iCreditDays')]")
			private static WebElement fields_iCreditDays; 	
				
			@FindBy(xpath="//span[contains(text(),'iDebitCreditProposal')]")
			private static WebElement fields_iDebitCreditProposal; 	
				
			@FindBy(xpath="//span[contains(text(),'iDebitCreditRequired')]")
			private static WebElement fields_iDebitCreditRequired; 	
				
			@FindBy(xpath="//span[contains(text(),'iExchangeAdjustmentGainAC')]")
			private static WebElement fields_iExchangeAdjustmentGainAC; 	
				
			@FindBy(xpath="//span[contains(text(),'iExchangeAdjustmentLossAC')]")
			private static WebElement fields_iExchangeAdjustmentLossAC; 	
				
			@FindBy(xpath="//span[contains(text(),'iPrimaryAccount')]")
			private static WebElement fields_iPrimaryAccount; 	
				
			@FindBy(xpath="//span[contains(text(),'iDefaultCurrency')]")
			private static WebElement fields_iDefaultCurrency; 	
				
			@FindBy(xpath="//span[contains(text(),'iConsolidationMethod')]")
			private static WebElement fields_iConsolidationMethod; 	
				
			@FindBy(xpath="//span[contains(text(),'iPaymentTerms')]")
			private static WebElement fields_iPaymentTerms; 	
				
			@FindBy(xpath="//span[contains(text(),'iReminderTerms')]")
			private static WebElement fields_iReminderTerms; 	
				
			@FindBy(xpath="//span[contains(text(),'iFinanceChargeTerms')]")
			private static WebElement fields_iFinanceChargeTerms; 	
				
			@FindBy(xpath="//span[contains(text(),'sAddress')]")
			private static WebElement fields_sAddress; 	
				
			@FindBy(xpath="//span[contains(text(),'sDeliveryAddress')]")
			private static WebElement fields_sDeliveryAddress; 	
				
			@FindBy(xpath="//span[contains(text(),'sPin')]")
			private static WebElement fields_sPin; 	
				
			@FindBy(xpath="//span[contains(text(),'sDeliveryPin')]")
			private static WebElement fields_sDeliveryPin; 	
				
			@FindBy(xpath="//span[contains(text(),'iCity')]")
			private static WebElement fields_iCity; 	
				
			@FindBy(xpath="//span[contains(text(),'iDeliveryCity')]")
			private static WebElement fields_iDeliveryCity; 	
				
			@FindBy(xpath="//span[contains(text(),'sEMail')]")
			private static WebElement fields_sEmail; 	
				
			@FindBy(xpath="//span[contains(text(),'bSendEmailtocustomer')]")
			private static WebElement fields_bSendEmailtocustomer; 	
				
			@FindBy(xpath="//span[contains(text(),'fChequeDiscountLimit')]")
			private static WebElement fields_fChequeDiscountLimit; 	
				
			@FindBy(xpath="//span[contains(text(),'fRateofinterest')]")
			private static WebElement fields_fRateofinterest; 	
				
			@FindBy(xpath="//span[contains(text(),'iBankAc')]")
			private static WebElement fields_iBankAc; 	
				
			@FindBy(xpath="//span[contains(text(),'sTelNo')]")
			private static WebElement fields_sTelNo; 	
				
			@FindBy(xpath="//span[contains(text(),'sFaxNo')]")
			private static WebElement fields_sFaxNo; 	
				
			@FindBy(xpath="//span[contains(text(),'sPassword')]")
			private static WebElement fields_sPassword; 	
				
			@FindBy(xpath="//span[contains(text(),'bAllowCustomerPortal')]")
			private static WebElement fields_bAllowCustomerPortal; 	
				
			@FindBy(xpath="//span[contains(text(),'sBankAccountName')]")
			private static WebElement fields_sBankAccountName; 	
				
			@FindBy(xpath="//span[contains(text(),'sBankAccountNumber')]")
			private static WebElement fields_sBankAccountNumber; 	
				
			@FindBy(xpath="//span[contains(text(),'sIFSCCode')]")
			private static WebElement fields_sIFSCCode; 	
				
			@FindBy(xpath="//span[contains(text(),'iPaymentType')]")
			private static WebElement fields_iPaymentType; 	
				
			@FindBy(xpath="//span[contains(text(),'iPDCDiscountedAC')]")
			private static WebElement fields_iPDCDiscountedAC; 	
				
			@FindBy(xpath="//span[contains(text(),'sFinanceEmail')]")
			private static WebElement fields_sFinanceEmail; 	
				
			@FindBy(xpath="//span[contains(text(),'sPortalEmail')]")
			private static WebElement fields_sPortalEmail; 	
			
			
		//UserDetails
		@FindBy(xpath="//span[contains(text(),'User details')]//i[@class='icon-expand icon-font8']")
		private static WebElement  fields_UserDetailsExpandBtn;	

			@FindBy(xpath="//li[@class='treeview active']//li[1]//a[1]//span[1]")
			private static WebElement userDetails_UserName;
			
			@FindBy(xpath="//li[@class='treeview active']//li[2]//a[1]//span[1]")
			private static WebElement userDetails_MobileNumber;
			
			@FindBy(xpath="//li[@class='treeview active']//li[3]//a[1]//span[1]")
			private static WebElement userDetails_PhoneNumber;
			
			@FindBy(xpath="//li[@class='treeview active']//li[4]//a[1]//span[1]")
			private static WebElement userDetails_EmailId;
			
			@FindBy(xpath="//li[@class='treeview active']//li[5]//a[1]//span[1]")
			private static WebElement userDetails_Signature;
			
			@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Employee')]")
			private static WebElement userDetails_Employee;
			
			@FindBy(xpath="//li[@class='treeview active']//li[7]//a[1]//span[1]")
			private static WebElement userDetails_AuthoriseUserName;
			
			@FindBy(xpath="//span[contains(text(),'Authorise User Mobile number')]")
			private static WebElement userDetails_AuthoriseUserMobileNumber;
			
			@FindBy(xpath="//span[contains(text(),'Authorise User Phone number')]")
			private static WebElement userDetails_AuthoriseUserPhoneNumber;
			
			@FindBy(xpath="//span[contains(text(),'Authorise User E-Mail Id')]")
			private static WebElement userDetails_AuthoriseUserEmailId;
			
			@FindBy(xpath="//span[contains(text(),'Authorise User Signature')]")
			private static WebElement userDetails_AuthoriseUserSignature;
			
			@FindBy(xpath="//span[contains(text(),'Authorise by')]")
			private static WebElement userDetails_AuthoriseBy;
			
			@FindBy(xpath="//span[contains(text(),'Next Authorise User name')]")
			private static WebElement userDetails_NextAuthoriseUserName;
			
			@FindBy(xpath="//span[contains(text(),'Next User Mobile number')]")
			private static WebElement userDetails_NextUserMobileNumber;
			
			@FindBy(xpath="//span[contains(text(),'Next User Phone number')]")
			private static WebElement userDetails_NextUserPhoneNumber;
			
			@FindBy(xpath="//span[contains(text(),'Next User E-Mail Id')]")
			private static WebElement userDetails_NextUserEmailId;
			
			@FindBy(xpath="//span[contains(text(),'Next User Signature')]")
			private static WebElement userDetails_NextUserSignature;
			
			@FindBy(xpath="//span[contains(text(),'Created by signature')]")
			private static WebElement userDetails_CreatedBySignature;
			               
			@FindBy(xpath="//span[contains(text(),'Modified by signature')]")
			private static WebElement userDetails_ModifiedBySignature;


	  //Miscellaneous 
	  @FindBy(xpath="//span[contains(text(),'Miscellaneous')]//i[@class='icon-expand icon-font8']")
	  private static WebElement  fields_MiscellaneousExpandBtn;

		@FindBy(xpath="//span[contains(text(),'Login date')]")
		private static WebElement miscellaneous_Logindate;
		
		@FindBy(xpath="//span[contains(text(),'Page Number')]")
		private static WebElement miscellaneous_PageNumber;
		
		@FindBy(xpath="//span[contains(text(),'Print Date')]")
		private static WebElement miscellaneous_Printdate;
		
		@FindBy(xpath="//span[contains(text(),'Print Time')]")
		private static WebElement miscellaneous_PrintTime;
		
		@FindBy(xpath="//span[contains(text(),'Total Page')]")
		private static WebElement miscellaneous_TotalPage;
		
		@FindBy(xpath="//span[contains(text(),'S.No')]")
		private static WebElement miscellaneous_Sno;
		
		@FindBy(xpath="//span[contains(text(),'Page Type')]")
		private static WebElement miscellaneous_PageType;
		
		@FindBy(xpath="//span[contains(text(),'Ageing details')]")
		private static WebElement miscellaneous_AgeingDetails;
		
		@FindBy(xpath="//span[contains(text(),'Ageing amounts')]")
		private static WebElement miscellaneous_AgeingAmounts;
		
		@FindBy(xpath="//span[contains(text(),'Ageing headings')]")
		private static WebElement miscellaneous_AgeingHeadings;
		
		@FindBy(xpath="//span[contains(text(),'Ageing balance')]")
		private static WebElement miscellaneous_Ageingbalance;
		
		@FindBy(xpath="//span[contains(text(),'Ageing slab1')]")
		private static WebElement miscellaneous_AgeingSlab1;
		
		@FindBy(xpath="//span[contains(text(),'Ageing slab2')]")
		private static WebElement miscellaneous_AgeingSlab2;
		
		@FindBy(xpath="//span[contains(text(),'Ageing slab3')]")
		private static WebElement miscellaneous_AgeingSlab3;
		
		@FindBy(xpath="//span[contains(text(),'Ageing slab4')]")
		private static WebElement miscellaneous_AgeingSlab4;
		
		@FindBy(xpath="//span[contains(text(),'Ageing slab5')]")
		private static WebElement miscellaneous_AgeingSlab5;
		
		@FindBy(xpath="//span[contains(text(),'Ageing slab6')]")
		private static WebElement miscellaneous_AgeingSlab6;
		
		@FindBy(xpath="//span[contains(text(),'Ageing month-wise')]")
		private static WebElement miscellaneous_AgeingMonthWise;


		//'SMS template 
		@FindBy(xpath="//u[contains(text(),'SMS template')]")
		private static WebElement headerSMSTemplateBtn;					
		
		//Pop As Same Above EmailHeaderTemplate
		//Alerts
		@FindBy(xpath="//u[contains(text(),'Alerts')]")
		private static WebElement headerAlert;

		//Pop As Same Above EmailHeaderTemplate
		//Recipient
		@FindBy(xpath="//div[contains(text(),'Recipient')]")
		private static WebElement headerRecipient;					
		
		//Pop As Same Above EmailHeaderTemplate
		//Escalation
		@FindBy(xpath="//span[@class='font-5'][contains(text(),'Escalation')]")
		private static WebElement escalationTab;					
				
		
		@FindBy(xpath="//*[@id='spnAddEscalation']/i")
		private static WebElement escalationLevelAddBtn;					
			

		//First Escalation
		@FindBy(xpath="//*[@id='txtTimeLimit1']")
		private static WebElement timelimitTxt1;					
				
		@FindBy(xpath="//*[@id='ddlTimeLimit1']")
		private static WebElement dropdown1;					
				
		@FindBy(xpath="//*[@id='ddlEscalationType1']")
		private static WebElement escalationTypeDropdown1;	
		
		@FindBy(xpath="//*[@id='divEscalation1']/div[1]/div[2]/i[1]")
		private static WebElement minimize1;					
		
		@FindBy(xpath="//*[@id='divEscalation1']/div[1]/div[2]/i[2]")
		private static WebElement close1;					
			
		//Second Escalation
		@FindBy(xpath="//*[@id='txtTimeLimit2']")
		private static WebElement timelimitTxt2;					
				
		@FindBy(xpath="//*[@id='ddlTimeLimit2']")
		private static WebElement dropdown2;					
				
		@FindBy(xpath="//*[@id='ddlEscalationType2']")
		private static WebElement escalationTypeDropdown2;	
		
		@FindBy(xpath="//*[@id='divEscalation2']/div[1]/div[2]/i[1]")
		private static WebElement minimize2;					
		
		@FindBy(xpath="//*[@id='divEscalation2']/div[1]/div[2]/i[2]")
		private static WebElement close2;				


		//Ok And Cancel for Authorization Definition
		@FindBy(xpath="//*[@id='id_report_popup_ok']/i")
		private static WebElement definitionOkBtn;					
					
		@FindBy(xpath="//*[@id='id_report_popup_cancel']/i	")
		private static WebElement definitionCancelBtn;					
			
		@FindBy(xpath="//*[@id='16']/div/span")
		private static WebElement  settingsmenuBtn;
		
		@FindBy(xpath="//span[contains(text(),'Configure Transactions')]")
		private static WebElement  configureTransactionBtn;

		@FindBy(xpath="//*[@id='68']/span")
		private static WebElement  transactionAuthorizationBtn;
		
		@FindBy(xpath="//*[@id='235']/span")
		private static WebElement  pricebookAuthorizationBtn;
		
		@FindBy(xpath="//*[@id='19']/span")
		private static WebElement  configureMastersBtn;
		
		@FindBy(xpath="//*[@id='67']/span")
		private static WebElement  masterAuthorizationBtn;
		
		@FindBy(xpath="//*[@id='69']/span")
		private static WebElement  creditLimitAuthorizationBtn;
		
		@FindBy(xpath="//*[@id='74']/span")
		private static WebElement  designWorkflowBtn;
		
		@FindBy(xpath="//*[@id='105']/span")
		private static WebElement  settingsWizardBtn;
		
		@FindBy(xpath="//*[@id='232']/span")
		private static WebElement  batchCodeGenerationBtn;
		
		@FindBy(xpath="//*[@id='34']/span")
		private static WebElement  mapCreditLimitBtn;
		
		@FindBy(xpath="//*[@id='257']/span")
		private static WebElement  vatTaxCodeBtn;
		
		@FindBy(xpath="//*[@id='258']/span")
		private static WebElement  customerPortalAdminBtn;
		
		@FindBy(xpath="//*[@id='3308']/span")
		private static WebElement  budgetAuthorizationBtn;
		
		@FindBy(xpath="//*[@id='3320']/span")
		private static WebElement  paymentGatewayIntegrationBtn;

		@FindBy(xpath="//*[@id='dashName']")
		private static WebElement labelDashboard ;

		
		
		public boolean 	checkOpenBudgetAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAuthorizationBtn));
			budgetAuthorizationBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			
		    boolean actbudgetDropdown                       = budgetDropdown.isDisplayed();
			boolean actnameTxt                              = nameTxt.isDisplayed();
			boolean actactiveChkBox                         = activeChkBox.isDisplayed();
			boolean actalwaysauthorstartsfirstlevelChkBox   = alwaysauthorizationstartsfromfirstlevelChkBox.isDisplayed();
			boolean actaddLevelBtn                          = addLevelBtn.isDisplayed();
			boolean actaddParallelBtn                       = addParallelBtn.isDisplayed();
			boolean actclearbtn               				= clearbtnInAuth.isDisplayed();
			boolean actsaveBtn  							= saveBtn.isDisplayed();
			boolean actcancelBtn 							= cancelBtnInAuth.isDisplayed();
			boolean actdeleteBtn			                = deleteBtn.isDisplayed();
			
			boolean expbudgetDropdown                       = true;
			boolean expnameTxt            					= true;
			boolean expactiveChkBox                			= true;
			boolean expalwaysauthorstartsfirstlevelChkBox   = true;
			boolean expaddLevelBtn          				= true;
			boolean expaddParallelBtn              			= true;
			boolean expclearbtn               				= true;
			boolean expsaveBtn  							= true;
			boolean expcancelBtn							= true;
			boolean expdeleteBtn							= true;
			
			System.out.println("********************* checkTransactionAuthorizationOptions ******************");
		
			System.out.println("Home Page budgetDropdown                     :"+actbudgetDropdown                      +"  Value Expected : "+expbudgetDropdown);
			System.out.println("Home Page nameTxt                            :"+actnameTxt                             +"  Value Expected : "+expnameTxt);
			System.out.println("Home Page activeChkBox                       :"+actactiveChkBox                        +"  Value Expected : "+expactiveChkBox);
			System.out.println("Home Page alwaysauthorstartsfirstlevelChkBox :"+actalwaysauthorstartsfirstlevelChkBox  +"  Value Expected : "+expalwaysauthorstartsfirstlevelChkBox);
	        System.out.println("Home Page addLevelBtn                        :"+actaddLevelBtn                         +"  Value Expected : "+expaddLevelBtn);
			System.out.println("Home Page addParallelBtn                     :"+actaddParallelBtn                      +"  Value Expected : "+expaddParallelBtn);
			System.out.println("Home Page clearbtn                           :"+actclearbtn                            +"  Value Expected : "+expclearbtn);
			System.out.println("Home Page saveBtn                            :"+actsaveBtn                             +"  Value Expected : "+expsaveBtn);
			System.out.println("Home Page cancelBtn 						 :"+actcancelBtn                           +"  Value Expected : "+expcancelBtn);
			System.out.println("Home Page deleteBtn                          :"+actdeleteBtn                           +"  Value Expected : "+expdeleteBtn);
			
			
			boolean res =  actbudgetDropdown==expbudgetDropdown  && actnameTxt==expnameTxt &&
					actactiveChkBox==expactiveChkBox && actalwaysauthorstartsfirstlevelChkBox==expalwaysauthorstartsfirstlevelChkBox && 
					actaddLevelBtn==expaddLevelBtn  && actaddParallelBtn==expaddParallelBtn && 
				    actclearbtn==expclearbtn && actsaveBtn==expsaveBtn && actcancelBtn==expcancelBtn && actdeleteBtn==expdeleteBtn;
			
			String actres = Boolean.toString(res);
			String expres = excelReader.getCellData("BudgetAuthorization", 11, 7);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 11, 8, actres);
			
			System.err.println("Dispaly Result : "+actres+" Value Expected : "+expres);
			
			if(actres.equalsIgnoreCase(expres))
				
			{
				System.out.println("Test Pass : Master Authorization is Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 10, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Master Authorization is Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 10, 9, resFail);
				return false;
			}
		}


		@FindBy(xpath="//li[contains(text(),'Define Budget')]")
		private static WebElement defineBudget;
		
		@FindBy(xpath="//li[contains(text(),'Revise Budget')]")
		private static WebElement reviseBudget;



		public boolean 	checkBudgetDropdownOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
				
			String actbudgetDropdown=budgetDropdown.getAttribute("placeholder");
			String expbudgetDropdown="Budget Type";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			Thread.sleep(2000);
			
		
			 boolean actdefineBudget        = defineBudget.isDisplayed();
			 boolean actreviseBudget        = reviseBudget.isDisplayed();
			 
			 boolean expdefineBudget        = true;
			 boolean expreviseBudget        = true;

				
			 System.out.println("****************************** checkBudgetDropdownOptions  ********************************************");
			 System.out.println("defineBudget        : "+actdefineBudget         +"   Value Expected  : "+expdefineBudget);
			 System.out.println("reviseBudget               :"+actreviseBudget                 +"   Value Expected : "+expreviseBudget);
			 
			 boolean res =  actdefineBudget==expdefineBudget && actreviseBudget==expreviseBudget;
				
			 String actres = Boolean.toString(res);
			 String expres = excelReader.getCellData("BudgetAuthorization", 13, 7);
				
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 13, 8, actres);
				
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Master Dropdown Options Are As expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 12, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Master Dropdown Options Are NOT As expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 12, 9, resFail);
				return false;
			}
		}



		public boolean 	checkSaveWithoutSelectionOfBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			
			System.out.println("************************************checkSaveWithoutSelectionOfBudget**********************************");
			
			String expMessage=excelReader.getCellData("BudgetAuthorization", 15, 7);
			
			String actMessage=checkValidationMessage(expMessage);	
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 15, 8, actMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 14, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Transaction Authorization is Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 14, 9, resFail);
				return false;
			}
		}
		
		

		public boolean 	checkSelectingDefineBudgetinBudgetDropdown() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineBudget));
			defineBudget.click();
			
			Thread.sleep(2000);
			
			
			String actbudgetDropdown=budgetDropdown.getAttribute("value");
			String expbudgetDropdown=excelReader.getCellData("BudgetAuthorization", 17, 7);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 17, 8, actbudgetDropdown);

			System.out.println("********************************checkSelectingDefineBudgetinBudgetDropdown*********************"); 
			
			
			System.out.println("budgetDropdown   : "+actbudgetDropdown+" Value Expected : "+expbudgetDropdown);
			
			if(actbudgetDropdown.equalsIgnoreCase(expbudgetDropdown))
			{
				System.out.println("Test Pass : Account is Selected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 16, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Account is Not Selected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 16, 9, resFail);
				return false;
			}
			
			
		}


		
		public boolean 	checkBannerTextNameTextBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			
			//Checking 
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			
			Thread.sleep(2000);
			
			String actnameTxt=nameTxt.getAttribute("placeholder");
			String expnameTxt=excelReader.getCellData("BudgetAuthorization", 19, 7);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 19, 8, actnameTxt);
			
			System.out.println("*****************************checkBannerTextNameTextBox***************");
			
			System.out.println("nameTxt   : "+actnameTxt+"  Value Expected : "+expnameTxt);
			
			if(actnameTxt.equalsIgnoreCase(expnameTxt))
			{
				System.out.println("Test Pass : Add New Authorization Flow is Displyed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 18, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Add New Authorization Flow is NOT Displyed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 18, 9, resFail);
				return false;
			}
		}
		
		
		
		public boolean 	checkSaveWithoutName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			System.out.println("********************************************checkSaveWithoutName***************************");
			
			String expMessage=excelReader.getCellData("BudgetAuthorization", 21, 7);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 21, 8, actMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 20, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Transaction Authorization is Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 20, 9, resFail);
				return false;
			}
		}
		
		
		
		
		public boolean 	checkActiveCheckBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(activeChkBox));
			
			boolean actactiveChkBox=activeChkBox.isSelected();
			boolean expactiveChkBox=true;
			
			System.out.println("*****************************************checkActiveCheckBox***********");
			
			System.out.println("activeChkBox   :"+actactiveChkBox +"  Value Expected : "+expactiveChkBox);
			
			boolean res = actactiveChkBox==expactiveChkBox;
			
			String actres = Boolean.toString(res);
			String expres = excelReader.getCellData("BudgetAuthorization", 23, 7);
				
			excelReader.setCellData(xlfile, "BudgetAuthorization", 23, 8, actres);
				
			if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Active ChkBox IS Enabled/Checked");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 22, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Active ChkBox IS NOT Enabled/Checked");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 22, 9, resFail);
				return false;
			}
		}
		
		
		public boolean 	checkInputNameAndClickOnSaveWithoutAuthorizationSettings() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.sendKeys("Define");
			nameTxt.sendKeys(Keys.TAB);
			
			
			Thread.sleep(2000);
			
			String actnameTxt=nameTxt.getAttribute("value");
			String expnameTxt=excelReader.getCellData("BudgetAuthorization", 25, 7);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 25, 8, actnameTxt);
			
			System.out.println("**********************************************checkInputNameAndClickOnSaveWithoutAuthorizationSettings**********************");
			
			System.out.println("nameTxt   :"+actnameTxt +" Value Expected : "+expnameTxt);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String expMessage=excelReader.getCellData("BudgetAuthorization", 26, 7);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 26, 8, actMessage);
			
			if(actnameTxt.equalsIgnoreCase(expnameTxt) && actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Test Pass : Account inputed Successfully");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 24, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Account NOT inputed Successfully");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 24, 9, resFail);
				return false;
			}
		}
		
		
		
		public boolean 	checkStartPoint() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			 getAction().doubleClick(startPointBtn).build().perform();
			
			 Thread.sleep(2000);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
			 boolean actconditionTab            = conditionTab.isDisplayed();
			 boolean actuserSelectionTab        = userSelectionTab.isDisplayed();
			 boolean actalertsTab               = alertsTab.isDisplayed();
			  
			 boolean expconditionTab            = true;
			 boolean expuserSelectionTab        = true;
			 boolean expalertsTab               = true;
			  
			 System.out.println("***********************************************checkStartPoint****************");
			  
			 System.out.println("conditionTab            "+actconditionTab         +"  Value Expected  :"+expconditionTab);
			 System.out.println("userSelectionTab        "+actuserSelectionTab     +"  Value Expected  :"+expuserSelectionTab);
			 System.out.println("alertsTab               "+actalertsTab            +"  Value Expected  :"+expalertsTab);
			 
			 boolean res = actconditionTab==expconditionTab && actuserSelectionTab==expuserSelectionTab && actalertsTab==expalertsTab;
				
			 String actres = Boolean.toString(res);
			 String expres = excelReader.getCellData("BudgetAuthorization", 28, 7);
					
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 28, 8, actres);
			  
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Start Point works As Expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 27, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Start Point NOT works As Expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 27, 9, resFail);
				return false;
			}
		}
		
		
		public boolean 	checkConditionTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
			 conditionTab.click();
			
			 Thread.sleep(2000);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			 boolean actdescriptionTxt                    = descriptionTxt.isDisplayed();
			 boolean actconditionNotRequiredChkBox        = conditionNotRequiredChkBox.isDisplayed();
			 boolean actdateRangeChkBox                   = dateRangeChkBox.isDisplayed();
			 boolean acttimeRangeChkBox           		  = timeRangeChkBox.isDisplayed();
			 boolean actfromDateTxt           			  = fromDateTxt.isDisplayed();
			 boolean acttoDateTxt           			  = toDateTxt.isDisplayed();
			 boolean actfromTimeTxt            			  = fromTimeTxt.isDisplayed();
			 boolean acttotimeTxt                         = totimeTxt.isDisplayed();
			 boolean actconjunctionDropdown        		  = conjunctionDropdown.isDisplayed();
			 boolean actselectFieldTxt               	  = selectFieldTxt.isDisplayed();
			 boolean actselectOperatorDropdown            = selectOperatorDropdown.isDisplayed();
			 boolean actcompareWithDropdown               = compareWithDropdown.isDisplayed();
			 boolean actvalueTxt               			  = valueTxt.isDisplayed();
			 boolean actdeleteRow1Btn                     = deleteRow1Btn.isDisplayed();
			 boolean actaddRowBtn              			  = addRowBtn.isDisplayed();
				
			 boolean expdescriptionTxt                    = true;
			 boolean expconditionNotRequiredChkBox        = true;
			 boolean expdateRangeChkBox                   = true;
			 boolean exptimeRangeChkBox           		  = true;
			 boolean expfromDateTxt           			  = true;
			 boolean exptoDateTxt                         = true;
			 boolean expfromTimeTxt                       = true;
			 boolean exptotimeTxt                         = true;
			 boolean expconjunctionDropdown               = true;
			 boolean expselectFieldTxt                    = true;
			 boolean expselectOperatorDropdown            = true;
			 boolean expcompareWithDropdown               = true;
			 boolean expvalueTxt               			  = true;
			 boolean expdeleteRow1Btn   				  = true;
			 boolean expaddRowBtn              			  = true;
			 
			 System.out.println("****************************** checkConditionTabOptions  ********************************************");
			 
			 System.out.println("descriptionTxt               :"+actdescriptionTxt             +"    Value Expected : "+expdescriptionTxt);
			 System.out.println("conditionNotRequiredChkBox   :"+actconditionNotRequiredChkBox +"    Value Expected : "+expconditionNotRequiredChkBox);
			 System.out.println("dateRangeChkBox              :"+actdateRangeChkBox            +"	 Value Expected : "+expdateRangeChkBox);
			 System.out.println("timeRangeChkBox              :"+acttimeRangeChkBox            +"	 Value Expected : "+exptimeRangeChkBox);
			 System.out.println("fromDateTxt                  :"+actfromDateTxt                +"	 Value Expected : "+expfromDateTxt);
			 System.out.println("toDateTxt                    :"+acttoDateTxt                  +"	 Value Expected : "+exptoDateTxt);
			 System.out.println("fromTimeTxt                  :"+actfromTimeTxt                +"	 Value Expected : "+expfromTimeTxt);
			 System.out.println("totimeTxt                    :"+acttotimeTxt                  +"	 Value Expected : "+exptotimeTxt);
			 System.out.println("conjunctionDropdown          :"+actconjunctionDropdown        +"	 Value Expected : "+expconjunctionDropdown);
			 System.out.println("selectFieldTxt               :"+actselectFieldTxt             +"	 Value Expected : "+expselectFieldTxt);
			 System.out.println("selectOperatorDropdown       :"+actselectOperatorDropdown     +"	 Value Expected : "+expselectOperatorDropdown);
			 System.out.println("compareWithDropdown          :"+actcompareWithDropdown        +"    Value Expected : "+expcompareWithDropdown);
			 System.out.println("valueTxt                     :"+actvalueTxt                   +"	 Value Expected : "+expvalueTxt);
			 System.out.println("deleteRow1Btn                :"+actdeleteRow1Btn              +"	 Value Expected : "+expdeleteRow1Btn);
			 System.out.println("addRowBtn                    :"+actaddRowBtn                  +"	 Value Expected : "+expaddRowBtn);
			
			 boolean res = actdescriptionTxt==expdescriptionTxt  && actconditionNotRequiredChkBox==expconditionNotRequiredChkBox && actdateRangeChkBox==expdateRangeChkBox
					 && acttimeRangeChkBox==exptimeRangeChkBox && actfromDateTxt==expfromDateTxt && acttoDateTxt==exptoDateTxt&& actfromTimeTxt==expfromTimeTxt 
					 && acttotimeTxt==exptotimeTxt && actconjunctionDropdown==expconjunctionDropdown && actselectFieldTxt==expselectFieldTxt 
					 && actselectOperatorDropdown==expselectOperatorDropdown && actcompareWithDropdown==expcompareWithDropdown && actvalueTxt==expvalueTxt 
					 && actdeleteRow1Btn==expdeleteRow1Btn && actaddRowBtn==expaddRowBtn;
				
			 String actres = Boolean.toString(res);
			 String expres = excelReader.getCellData("BudgetAuthorization", 30, 7);
					
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 30, 8, actres);
			 
			 
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Condition Tab Options Are Displayed ");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 29, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Condition Tab Options Are NOT Displayed ");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 29, 9, resFail);
				return false;
			}
			
		}
		
		
		
		
		public boolean 	checkUserSelectionTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			Thread.sleep(2000);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userselectionDropdown));
			 boolean actuserselectionDropdown             = userselectionDropdown.isDisplayed();
			 boolean actcriteriaDropdown                  = criteriaDropdown.isDisplayed();
			 boolean actuserrow1                          = userrow1.isDisplayed();
			 boolean actuserrow2                          = userrow2.isDisplayed();
			  
			 boolean expuserselectionDropdown             = true;
			 boolean expcriteriaDropdown                  = true;
			 boolean expuserrow1                          = true;
			 boolean expuserrow2                          = true;
			  
			 System.out.println("***********************************************checkUserSelectionTabOptions****************");
			  
			 System.out.println("UserselectionDropdown    :"+actuserselectionDropdown  +"  Value Expected  :"+expuserselectionDropdown);
			 System.out.println("CriteriaDropdown         :"+actcriteriaDropdown       +"  Value Expected  :"+expcriteriaDropdown);
			 System.out.println("Userrow1                 :"+actuserrow1               +"  Value Expected  :"+expuserrow1);
			 System.out.println("userrow2                 :"+actuserrow2               +"  Value Expected  :"+expuserrow2);
			 
			 boolean res = actuserselectionDropdown==expuserselectionDropdown && actcriteriaDropdown==expcriteriaDropdown && 
					  actuserrow1==expuserrow1 && actuserrow2==expuserrow2;
				
			 String actres = Boolean.toString(res);
			 String expres = excelReader.getCellData("BudgetAuthorization", 32, 7);
					
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 32, 8, actres);
			  
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : User Selection TAb Options Are Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 31, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : User Selection TAb Options Are NOT Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 31, 9, resFail);
				return false;
			}
		}
		
		
		
		
		@FindBy(xpath="//*[@id='id_tblUserSelection_User']")
		private static WebElement userDropdown;
		
		
		public boolean 	checkUserDropdownOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{		
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userselectionDropdown));
			userselectionDropdown.click();
			
			String userq=userselectionDropdown.getAttribute("value");
			
			System.out.println("userq"+userq);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(criteriaDropdown));
			criteriaDropdown.click();
			
			String userc=criteriaDropdown.getText();
			
			System.out.println("userc"+userc);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDropdown));
			userDropdown.click();
			userDropdown.sendKeys("UserALOptionsST");
			userDropdown.sendKeys(Keys.ENTER);
			
			String Strlpl= userDropdown.getText();
		
			System.out.println(Strlpl);
			 
	        String[] testString=Strlpl.split("\n");
	        
	        List<String> listg = Arrays.asList(testString);
	        
	        System.out.println("********************************************checkUserDropdownOptions*************************************");
	        
	        System.out.println("Converting to string array"+listg);
	        
	        
	        String actlist = listg.toString();
	        String explist = "[SU, UserAllOptionsST, UserRestrictionST, UserExclusionST]";
	        
	        System.out.println("The Users in User Dropdown Value Expected : " + actlist + "Value Actual : " + explist);
			
	        
	       if(actlist.equalsIgnoreCase(explist))
	       {
	    	   excelReader.setCellData(xlfile, "BudgetAuthorization", 33, 9, resPass);
	    	   return true;
	       }
	       else
	       {
	    	   excelReader.setCellData(xlfile, "BudgetAuthorization", 33, 9, resFail);
	    	   return false;
	       }
	    	
	    	
	    
			
		 
		}

		
		public boolean 	checkAlertsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
			 alertsTab.click();
			
			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedChkBox));
			 boolean actreceivedChkBox                   = receivedChkBox.isDisplayed();
		     boolean actapprovedChkBox                   = approvedChkBox.isDisplayed();
			 boolean actrejectedChkBox                   = rejectedChkBox.isDisplayed();
			 boolean actemailHeaderTemplateBtn           = emailHeaderTemplateBtn.isDisplayed();
			 boolean actemailTemplateBtn                 = emailTemplateBtn.isDisplayed();
			 boolean actheaderAlert                      = headerAlert.isDisplayed();
			 
			 boolean expreceivedChkBox                   = true;
			 boolean expapprovedChkBox                   = true;
			 boolean exprejectedChkBox                   = true;
			 boolean expemailHeaderTemplateBtn           = true;
			 boolean expemailTemplateBtn                 = true;
			 boolean expheaderAlert                      = true;
			 	
			 System.out.println("****************************** checkAlertsTabOptions  ********************************************");
			 
			 System.out.println("ReceivedChkBox                :"+actreceivedChkBox         +"   Value Expected : "+expreceivedChkBox);
			 System.out.println("ApprovedChkBox                :"+actapprovedChkBox         +"   Value Expected : "+expapprovedChkBox);
			 System.out.println("RejectedChkBox                :"+actrejectedChkBox         +"	 Value Expected : "+exprejectedChkBox);
			 System.out.println("EmailHeaderTemplateBtn        :"+actemailHeaderTemplateBtn +"	 Value Expected : "+expemailHeaderTemplateBtn);
			 System.out.println("EmailTemplateBtn      	       :"+actemailTemplateBtn       +"	 Value Expected : "+expemailTemplateBtn);
			 System.out.println("HeaderAlert                   :"+actheaderAlert            +"	 Value Expected : "+expheaderAlert);
			 
			 boolean res = actreceivedChkBox==expreceivedChkBox&& actapprovedChkBox==expapprovedChkBox  && actrejectedChkBox==exprejectedChkBox 
					 && actemailHeaderTemplateBtn==expemailHeaderTemplateBtn 
					  && actemailTemplateBtn==expemailTemplateBtn && actheaderAlert==expheaderAlert ;
				
			 String actres = Boolean.toString(res);
			 String expres = excelReader.getCellData("BudgetAuthorization", 36, 7);
					
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 36, 8, actres);
			  	
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 35, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 35, 9, resFail);
				return false;
			}
		}
		
		
		
		
		public boolean 	checkCancelButtonInDefinitionScreen() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionCancelBtn));
			definitionCancelBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			
			boolean actbudgetDropdown                       = budgetDropdown.isDisplayed();
			boolean actnameTxt                              = nameTxt.isDisplayed();
			boolean actactiveChkBox                         = activeChkBox.isDisplayed();
			boolean actalwaysauthorstartsfirstlevelChkBox   = alwaysauthorizationstartsfromfirstlevelChkBox.isDisplayed();
			boolean actaddLevelBtn                          = addLevelBtn.isDisplayed();
			boolean actaddParallelBtn                       = addParallelBtn.isDisplayed();
			boolean actclearbtn               				= clearbtnInAuth.isDisplayed();
			boolean actsaveBtn  							= saveBtn.isDisplayed();
			boolean actcancelBtn 							= cancelBtn.isDisplayed();
			boolean actdeleteBtn			                = deleteBtn.isDisplayed();
			
			boolean expbudgetDropdown                       = true;
			boolean expnameTxt            					= true;
			boolean expactiveChkBox                			= true;
			boolean expalwaysauthorstartsfirstlevelChkBox   = true;
			boolean expaddLevelBtn          				= true;
			boolean expaddParallelBtn              			= true;
			boolean expclearbtn               				= true;
			boolean expsaveBtn  							= true;
			boolean expcancelBtn							= true;
			boolean expdeleteBtn							= true;
			
			System.out.println("********************************************** checkCancelButtonInDefinitionScreen ******************");
		
			System.out.println("Home Page budgetDropdown                     :"+actbudgetDropdown                      +"  Value Expected : "+expbudgetDropdown);
			System.out.println("Home Page nameTxt                            :"+actnameTxt                             +"  Value Expected : "+expnameTxt);
			System.out.println("Home Page activeChkBox                       :"+actactiveChkBox                        +"  Value Expected : "+expactiveChkBox);
			System.out.println("Home Page alwaysauthorstartsfirstlevelChkBox :"+actalwaysauthorstartsfirstlevelChkBox  +"  Value Expected : "+expalwaysauthorstartsfirstlevelChkBox);
	        System.out.println("Home Page addLevelBtn                        :"+actaddLevelBtn                         +"  Value Expected : "+expaddLevelBtn);
			System.out.println("Home Page addParallelBtn                     :"+actaddParallelBtn                      +"  Value Expected : "+expaddParallelBtn);
			System.out.println("Home Page clearbtn                           :"+actclearbtn                            +"  Value Expected : "+expclearbtn);
			System.out.println("Home Page saveBtn                            :"+actsaveBtn                             +"  Value Expected : "+expsaveBtn);
			System.out.println("Home Page cancelBtn 						 :"+actcancelBtn                           +"  Value Expected : "+expcancelBtn);
			System.out.println("Home Page deleteBtn                          :"+actdeleteBtn                           +"  Value Expected : "+expdeleteBtn);
			
			
			 boolean res =  actbudgetDropdown==expbudgetDropdown  && actnameTxt==expnameTxt &&
						actactiveChkBox==expactiveChkBox && actalwaysauthorstartsfirstlevelChkBox==expalwaysauthorstartsfirstlevelChkBox && 
						actaddLevelBtn==expaddLevelBtn  && actaddParallelBtn==expaddParallelBtn && 
					    actclearbtn==expclearbtn && actsaveBtn==expsaveBtn && actcancelBtn==expcancelBtn && actdeleteBtn==expdeleteBtn ;
				
			 String actres = Boolean.toString(res);
			 String expres = excelReader.getCellData("BudgetAuthorization", 38, 7);
					
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 38, 8, actres);
			  	
			if(actres.equalsIgnoreCase(expres))
				
			{
				System.out.println("Test Pass : Cancel Option Working ");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 37, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Cancel Option Not Working");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 37, 9, resFail);
				return false;
			}
		}
		
		
		
		public boolean 	checkDescriptionAsMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			
			getAction().doubleClick(startPointBtn).build().perform();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			System.out.println("****************************checkDescriptionAsMandatory************************************"); 
		
			
			String expMessage=excelReader.getCellData("BudgetAuthorization", 40, 7);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 40, 8, actMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 39, 9, resPass);
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 39, 9, resFail);
				return false;
			}
		}
		
		
		
		
		
		public boolean 	checkConditionAsMandatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			
			descriptionTxt.sendKeys("Define");
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			System.out.println("*****************************************************checkConditionAsMandatory************************************");
			
			String expMessage=excelReader.getCellData("BudgetAuthorization", 42, 7);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 42, 8, actMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 41, 9, resPass);
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 41, 9, resFail);
				return false;
			}
		}
		
		
		
		public boolean 	checkUserSelectionIsManadatory() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			System.out.println("*******************************************************checkUserSelectionIsManadatory*****************************************");
			
			String expMessage=excelReader.getCellData("BudgetAuthorization", 44, 7);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 44, 8, actMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 43, 9, resPass);
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 43, 9, resFail);
				return false;
			}
		}
		
		
		@FindBy(xpath="//div[@class='divStopNode']//i[@class='icon-end icon-font4 grey']")
		private static WebElement endpoint;
		
		
		public boolean 	checkOkButtonOnSelectingUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDropdown));
			userDropdown.click();
			
			userDropdown.sendKeys("UserAllOptionsST");
			
			/*Select s=new Select(userDropdown);
			s.selectByVisibleText(excelReader.getCellData("BudgetAuthorization", 45, 6));*/
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			boolean Endpoint=endpoint.isDisplayed();


			String actendpoint = Boolean.toString(Endpoint);
			
			String expendpoint = excelReader.getCellData("BudgetAuthorization", 46, 7);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 46, 8, actendpoint);
			
			System.out.println("***************************************checkOkButtonOnSelectingUser*****************************");
			System.out.println("endpoint     : "+actendpoint +"  Value Expected : "+expendpoint);
			
			if(actendpoint.equalsIgnoreCase(expendpoint))
			{
				System.out.println("Test Pass : Image Displayed In Design Area");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 45, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 45, 9, resFail);
				return false;
			}
		}
		
		
		
		
		
		
		public boolean 	checkSavingDefineBudgetAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{		
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			System.out.println("**********************************checkSavingDefineBudgetAuthorization*****************************");
			
			String expMessage=excelReader.getCellData("BudgetAuthorization", 48, 7);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 48, 8, actMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 47, 9, resPass);
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 47, 9, resFail);
				return false;
			}
		}
		
		
		
		
		public boolean 	checkCancelButton() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
			cancelBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
			
			boolean Dashboard=dashboard.isDisplayed();
			
			String actdashboard = Boolean.toString(Dashboard);
			
			String expdashboard = excelReader.getCellData("BudgetAuthorization", 50, 7);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 50, 8, actdashboard);	
			
			System.out.println("********************************checkCancelButton*************************");
			
			System.out.println("dashboard  :"+actdashboard +"  Value Expected  : "+expdashboard);
			
			if(actdashboard.equalsIgnoreCase(expdashboard))
				
			{
				System.out.println("Test Pass : Cancel Option Working ");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 49, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Cancel Option Not Working");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 49, 9, resFail);
				return false;
			}
		}
		
		public boolean 	checkDisplayingMastersNameInListOfMasters() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAuthorizationBtn));
			budgetAuthorizationBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			boolean DefineBudget=defineBudget.isDisplayed();


			String actdefineBudget = Boolean.toString(DefineBudget);
			String expdefineBudget = excelReader.getCellData("BudgetAuthorization", 52, 7);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 52, 8, actdefineBudget);	
			
			System.out.println("***********************************************checkDisplayingMastersNameInListOfMasters*******");
			
			System.out.println("defineBudget       : "+actdefineBudget+"  Value Expected : "+expdefineBudget);
			
			if(actdefineBudget.equalsIgnoreCase(expdefineBudget))
			{
				System.out.println("Test Pass : Account(1) Is Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 51, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Account(1) Is NOT Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 51, 9, resFail);
				return false;
			}
		}
			
		
		
		public boolean 	checkSavedNameDisplayInNameTextBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(defineBudget));
			defineBudget.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.sendKeys("Define");
			
			Thread.sleep(3000);
			
			nameTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			boolean actstartPointBtn     = startPointBtn.isDisplayed();
			boolean actendpoint          = endpoint.isDisplayed();
			
			boolean expstartPointBtn     = true;
			boolean expendpoint          = true;
			
			boolean res = actstartPointBtn==expstartPointBtn && actendpoint==expendpoint;
			
			String actres = Boolean.toString(res);
			
			String expres = excelReader.getCellData("BudgetAuthorization", 54, 7);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 54, 8, actres);
			
			System.out.println("*************************************checkSavedNameDisplayInNameTextBox************************************");
			
			System.out.println("startPointBtn    :"+actstartPointBtn   +"  Value Expected : "+expstartPointBtn);
			System.out.println("endpoint         :"+actendpoint        +"  Value Expected : "+expendpoint);
			
			if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Saved Data is Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 53, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Saved Data is NOT Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 53, 9, resFail);
				return false;
			}
			
		}
		


		public boolean 	checkEditBudgetAuthorizationSettings() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1editBtn));
			Level1editBtn.click();
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
	        boolean actconditionTab                 = conditionTab.isDisplayed();
			boolean actuserSelectionTab             = userSelectionTab.isDisplayed();
			boolean actalertsTab                    = alertsTab.isDisplayed();
			
			boolean expconditionTab                 = true;
			boolean expuserSelectionTab             = true;
			boolean expalertsTab                    = true;
			
			boolean res = actconditionTab==expconditionTab  && actuserSelectionTab==expuserSelectionTab && actalertsTab==expalertsTab;
			
			String actres = Boolean.toString(res);
			
			String expres = excelReader.getCellData("BudgetAuthorization", 56, 7);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 56, 8, actres);
			
			System.out.println("********************* checkEditBudgetAuthorizationSettings ******************");
		
			System.out.println("ConditionTab             :"+actconditionTab     +"  Value Expected : "+expconditionTab);
			System.out.println("UserSelectionTab         :"+actuserSelectionTab +"  Value Expected : "+expuserSelectionTab);
			System.out.println("AlertsTab                :"+actalertsTab        +"	Value Expected : "+expalertsTab);
			
			if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Edit Btn Opened definition Screen");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 55, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Edit Btn NOT Opened definition Screen");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 55, 9, resFail);
				return false;
			}
			
		}
		
		
		
		
		
		public boolean checkConditionCheckbox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			conjunctionDropdown.click();
			conjunctionDropdown.sendKeys(excelReader.getCellData("BudgetAuthorization", 57, 6));
			conjunctionDropdown.sendKeys(Keys.TAB);
			Thread.sleep(3000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
			selectFieldTxt.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(seletFieldAccount));
			seletFieldAccount.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			selectOperatorDropdown.click();
			selectOperatorDropdown.sendKeys(excelReader.getCellData("BudgetAuthorization", 58, 6));
			selectOperatorDropdown.sendKeys(Keys.TAB);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			compareWithDropdown.click();
			compareWithDropdown.sendKeys(excelReader.getCellData("BudgetAuthorization", 59, 6));
			compareWithDropdown.sendKeys(Keys.TAB);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_valueTxt));
			enter_valueTxt.clear();
			enter_valueTxt.sendKeys(excelReader.getCellData("BudgetAuthorization", 60, 6));
			Thread.sleep(2000);
			
			
			String actconjunctionDropdown=conjunctionDropdown.getAttribute("value");
			String expconjunctionDropdown=excelReader.getCellData("BudgetAuthorization", 58, 7);
			
			
			String actselectFieldTxt=selectFieldTxt.getAttribute("value");
			String expselectFieldTxt=excelReader.getCellData("BudgetAuthorization", 59, 7);
			
			
			String actselectOperatorDropdown=selectOperatorDropdown.getAttribute("value");
			String expselectOperatorDropdown=excelReader.getCellData("BudgetAuthorization", 60, 7);
			
			
			String actcompareWithDropdown=compareWithDropdown.getAttribute("value");
			String expcompareWithDropdown=excelReader.getCellData("BudgetAuthorization", 61, 7);
			
			String actvalueTxt=enter_valueTxt.getAttribute("value");
			String expvalueTxt=excelReader.getCellData("BudgetAuthorization", 62, 7);
			
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 58, 8, actconjunctionDropdown);
			excelReader.setCellData(xlfile, "BudgetAuthorization", 59, 8, actselectFieldTxt);
			excelReader.setCellData(xlfile, "BudgetAuthorization", 60, 8, actselectOperatorDropdown);
			excelReader.setCellData(xlfile, "BudgetAuthorization", 61, 8, actcompareWithDropdown);
			excelReader.setCellData(xlfile, "BudgetAuthorization", 62, 8, actvalueTxt);
			
			System.out.println("************************************checkConditionCheckbox******************************************************************************");
			
			System.out.println("conjunctionDropdown     :"+actconjunctionDropdown        +" Value Expected  : "+expconjunctionDropdown); 
			System.out.println("selectFieldTxt          :"+actselectFieldTxt             +" Value Expected  : "+expselectFieldTxt);
			System.out.println("selectOperatorDropdown  :"+actselectOperatorDropdown     +"  Value Expected : "+expselectOperatorDropdown); 
			System.out.println("compareWithDropdown     :"+actcompareWithDropdown        +" Value Expected  : "+expcompareWithDropdown);
			System.out.println("valueTxt                :"+actvalueTxt                   +" Value Expected  : "+expvalueTxt);
			
			if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && 
					actselectOperatorDropdown.equals(expselectOperatorDropdown)
					&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt))
			{
				System.out.println("Test Pass : Data entered");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 57, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT entered");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 57, 9, resFail);
				return false;
			}
			
		}
		
		
		@FindBy(xpath="//span[contains(text(),'Alerts')]")
		private static WebElement alertTabInEdit;
		
		@FindBy(xpath="//u[contains(text(),'SMS template')]")
		private static WebElement headerSMSTemplate;	
		
		
		public boolean 	checkEditOnSelectAlertsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			 Thread.sleep(2000);
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertTabInEdit));
			 alertTabInEdit.click();
			
			 Thread.sleep(2000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedChkBox));
			 boolean actreceivedChkBox                   = receivedChkBox.isDisplayed();
		     boolean actapprovedChkBox                   = approvedChkBox.isDisplayed();
			 boolean actrejectedChkBox                   = rejectedChkBox.isDisplayed();
			 boolean actemailHeaderTemplateBtn           = emailHeaderTemplateBtn.isDisplayed();
			 boolean actemailTemplateBtn                 = emailTemplateBtn.isDisplayed();
			 boolean actheaderAlert                      = headerAlert.isDisplayed();
			 
			 boolean expreceivedChkBox                   = true;
			 boolean expapprovedChkBox                   = true;
			 boolean exprejectedChkBox                   = true;
			 boolean expemailHeaderTemplateBtn           = true;
			 boolean expemailTemplateBtn                 = true;
			 boolean expheaderAlert                      = true;
			 	
			 System.out.println("****************************** checkEditOnSelectAlertsTabOptions  ********************************************");
			 
			 System.out.println("ReceivedChkBox                :"+actreceivedChkBox         +"   Value Expected : "+expreceivedChkBox);
			 System.out.println("ApprovedChkBox                :"+actapprovedChkBox         +"   Value Expected : "+expapprovedChkBox);
			 System.out.println("RejectedChkBox                :"+actrejectedChkBox         +"	 Value Expected : "+exprejectedChkBox);
			 System.out.println("EmailHeaderTemplateBtn        :"+actemailHeaderTemplateBtn +"	 Value Expected : "+expemailHeaderTemplateBtn);
			 System.out.println("EmailTemplateBtn      	       :"+actemailTemplateBtn       +"	 Value Expected : "+expemailTemplateBtn);
			 System.out.println("HeaderAlert                   :"+actheaderAlert            +"	 Value Expected : "+expheaderAlert);
			  	
			 
			 boolean res = actreceivedChkBox==expreceivedChkBox&& actapprovedChkBox==expapprovedChkBox  && actrejectedChkBox==exprejectedChkBox 
					 && actemailHeaderTemplateBtn==expemailHeaderTemplateBtn 
					  && actemailTemplateBtn==expemailTemplateBtn && actheaderAlert==expheaderAlert;
				
			 String actres = Boolean.toString(res);
			 
			 String expres = excelReader.getCellData("BudgetAuthorization", 64, 7);
			 
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 64, 8, actres);
				
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 63, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 63, 9, resFail);
				return false;
			}
		}
		
		
		
		@FindBy(xpath="//input[@id='smsTemplateOptCtrl']")
		private static WebElement emailTemplates;
		
		@FindBy(xpath="//i[@class='icon-open icon-font7']")
		private static WebElement companyMenu;
		
		@FindBy(xpath="//span[contains(text(),'Company Name')]")
		private static WebElement companyNameSource;
		
		@FindBy(xpath="//textarea[@id='txtTemplate']")
		private static WebElement companyNameTemplateArea;
		
		@FindBy(xpath="//span[@id='btnSaveSmsTemplate']")
		private static WebElement saveBtnInEmailTemplateHeader;
		
		@FindBy(xpath="//span[@id='btnNewSmsTemplate']")
		private static WebElement newButtonInCreateTemplate;
		
		@FindBy(xpath="//span[@id='btnCloseSMSModal']")
		private static WebElement closeButtonInCreateTemplate;
		
		
		
		public boolean checkSaveWithAllMandatoryInputs() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailHeaderTemplateBtn));
			emailHeaderTemplateBtn.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyMenu));
			emailTemplates.sendKeys(excelReader.getCellData("BudgetAuthorization", 65, 6));

			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
			companyNameTemplateArea.click();
			companyNameTemplateArea.sendKeys(excelReader.getCellData("BudgetAuthorization", 66, 6));
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
			companyNameTemplateArea.click();
			companyNameTemplateArea.sendKeys(excelReader.getCellData("BudgetAuthorization", 67, 6));
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
			saveBtnInEmailTemplateHeader.click();
			
			System.out.println("**************************************checkSaveWithAllMandatoryInputs*******************************");
			
			String expMessage=excelReader.getCellData("BudgetAuthorization", 66, 7);
			 
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 66, 8, actMessage);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
			closeButtonInCreateTemplate.click();
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 65, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 65, 9, resFail);
				return false;
			}
		}
		
		@FindBy(xpath="//td[@id='id_tblAlerts_col_1-2']")
		private static WebElement recievedEmailHeaderTemplate;
		
		@FindBy(xpath="//select[@id='id_tblAlerts_EmailHeaderTemplate']")
		private static WebElement receivedEmailHeaderTemplateInput;
		
		
		public boolean checkCreatedEmailInEmailHedaerTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedChkBox));
			receivedChkBox.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recievedEmailHeaderTemplate));
			recievedEmailHeaderTemplate.click();
			Thread.sleep(2000);
			
			//receivedEmailHeaderTemplateInput.click();
			
			//Thread.sleep(2000);
			
			
			Select s=new Select(receivedEmailHeaderTemplateInput);
			s.selectByVisibleText(excelReader.getCellData("BudgetAuthorization", 68, 6));
			
			receivedEmailHeaderTemplateInput.sendKeys(Keys.TAB);
			
			//receivedEmailHeaderTemplateInput.sendKeys("createEmailTemplate");
			
			Thread.sleep(2000);
			
			//receivedEmailHeaderTemplateInput.click();
			
			//Thread.sleep(2000);
			
			
			 String actreceivedEmailHeaderTemplateInput=recievedEmailHeaderTemplate.getText();
			
			
			String expreceivedEmailHeaderTemplateInput=excelReader.getCellData("BudgetAuthorization", 69, 7);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 69, 8, actreceivedEmailHeaderTemplateInput);
			
			System.out.println("actreceivedEmailHeaderTemplateInput : "+actreceivedEmailHeaderTemplateInput);
			System.out.println("expreceivedEmailHeaderTemplateInput : "+expreceivedEmailHeaderTemplateInput);
			
			System.out.println("**************************************checkCreatedEmailInEmailHedaerTemplate*********************");
			
			System.out.println(" receivedEmailHeaderTemplateInput  :"+actreceivedEmailHeaderTemplateInput+"  Value Expected : "+expreceivedEmailHeaderTemplateInput);
			
			if(actreceivedEmailHeaderTemplateInput.equalsIgnoreCase(expreceivedEmailHeaderTemplateInput))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 68, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 68, 9, resFail);
				return false;
			}
		}
		
		


	@FindBy(xpath="//span[contains(text(),'Header')]//i[@class='icon-expand icon-font8']")
	private static WebElement  Finance_Purchases_HeaderExpandBtn;

	@FindBy(xpath="//span[contains(text(),'Header')]//i[@class='icon-collepse icon-font8']")
	private static WebElement  Finance_Purchases_HeaderMinimiseBtn;

	@FindBy(xpath="//span[contains(text(),'Miscellaneous')]//i[@class='icon-expand icon-font8']")
	private static WebElement MiscellaneousExpandBtn;

	@FindBy(xpath="//span[contains(text(),'Company')]//i[@class='icon-expand icon-font8']")
	private static WebElement CompanyExpandBtn;

	@FindBy(xpath="//span[contains(text(),'Company')]//i[@class='icon-collepse icon-font8']")
	private static WebElement CompanyMinimiseBtn;

	@FindBy(xpath="//span[contains(text(),'User details')]//i[@class='icon-expand icon-font8']")
	private static WebElement  UserDetailsExpandbtn;

	@FindBy(xpath="//span[contains(text(),'Body')]//i[@class='icon-expand icon-font8']")
	private static WebElement  PSS_BodyExpandBtn;

	@FindBy(xpath="//span[contains(text(),'Body')]//i[@class='icon-collepse icon-font8']")
	private static WebElement  PSS_BodyMinimizeBtn;

								
		
				
	@FindBy(xpath="//span[contains(text(),'Company Name')]")
	private static WebElement  Company_CompanyNamedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Company Code')]")
	private static WebElement  Company_CompanyCodedraganddrop;

	@FindBy(xpath="//span[contains(text(),'Accounting Date')]")
	private static WebElement  Comapny_AccountingDatedrapandDrop;

	@FindBy(xpath="//span[contains(text(),'Address1')]")
	private static WebElement  Comapny_Address1drapandDrop;

	@FindBy(xpath="//span[contains(text(),'Address2')]")
	private static WebElement  Comapny_Address2drapandDrop;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'City')]")
	private static WebElement  Comapny_CitydrapandDrop;

	@FindBy(xpath="//span[contains(text(),'Zip Code')]")
	private static WebElement  Comapny_ZipCodedrapandDrop;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Country')]")
	private static WebElement  Comapny_CountrydrapandDrop;

	@FindBy(xpath="//span[contains(text(),'Phone No1')]")
	private static WebElement  Comapny_PhoneNo1yrapandDrop;

	@FindBy(xpath="//span[contains(text(),'Phone No2')]")
	private static WebElement  Comapny_PhoneNo2rapandDrop;

	@FindBy(xpath="//span[contains(text(),'Fax')]")
	private static WebElement  Comapny_FaxrapandDrop;
						
	@FindBy(xpath="//span[contains(text(),'Tax Registration Number')]")
	private static WebElement  Company_TaxRegistrationNumber;

	@FindBy(xpath="//span[contains(text(),'Excise Registration Number')]")
	private static WebElement  Company_ExciseRegistrationNumber;


	@FindBy(xpath="//i[@class='icon-font6 icon-new']")
	private static WebElement  Newdropdown;


	@FindBy(xpath="//span[contains(text(),'Users')]//i[@class='icon-expand icon-font8']")
	private static WebElement userExpandBtn;


	@FindBy(xpath="//span[contains(text(),'Users')]//i[@class='icon-collepse icon-font8']")
	private static WebElement userMinimizeBtn;


	@FindBy(xpath="//span[contains(text(),'Logged in Username')]")
	private static WebElement  user_LoggedInUserNameDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Logged in Login Abbreviation')]")
	private static WebElement  user_LoggedInLoginAbbreviationDragandDrop;

	@FindBy(xpath="//span[contains(text(),'End User')]")
	private static WebElement  user_EndUserDragandDrop;




		public boolean checkEmailTemplateUserExpandInFieldsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplateBtn));
		 	 emailTemplateBtn.click();
		
		     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Newdropdown));
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userExpandBtn));
			 userExpandBtn.click();
			
		
			 Thread.sleep(3000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(user_LoggedInUserNameDragandDrop));
			 boolean actuser_LoggedInUserNameDragandDrop             = user_LoggedInUserNameDragandDrop.isDisplayed();
		     boolean actuser_LoggedInLoginAbbreviationDragandDrop    = user_LoggedInLoginAbbreviationDragandDrop.isDisplayed();
			 boolean actuser_EndUserDragandDrop                      = user_EndUserDragandDrop.isDisplayed();
			 
			 boolean expuser_LoggedInUserNameDragandDrop             = true;
			 boolean expuser_LoggedInLoginAbbreviationDragandDrop    = true;
			 boolean expuser_EndUserDragandDrop                      = true;
			 
				
			 System.out.println("****************************** checkEmailTemplateUserExpandInFieldsTabOptions  ********************************************");
			 
			 System.out.println("user_LoggedInUserNameDragandDrop          :"+actuser_LoggedInUserNameDragandDrop           +"    Value Expected : "+expuser_LoggedInUserNameDragandDrop);
			 System.out.println("user_LoggedInLoginAbbreviationDragandDrop :"+actuser_LoggedInLoginAbbreviationDragandDrop  +"    Value Expected : "+expuser_LoggedInLoginAbbreviationDragandDrop);
			 System.out.println("user_EndUserDragandDrop                   :"+actuser_EndUserDragandDrop                    +"	  Value Expected : "+expuser_EndUserDragandDrop);
			 
			 
			 boolean res = actuser_LoggedInUserNameDragandDrop==expuser_LoggedInUserNameDragandDrop
					 && actuser_LoggedInLoginAbbreviationDragandDrop==expuser_LoggedInLoginAbbreviationDragandDrop 
					 && actuser_EndUserDragandDrop==expuser_EndUserDragandDrop;
				
			 String actres = Boolean.toString(res);
			 
			 String expres = excelReader.getCellData("BudgetAuthorization", 71, 7);
			 
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 71, 8, actres);
			 
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Fields Options Are Displayed");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userMinimizeBtn));
				userMinimizeBtn.click();
				excelReader.setCellData(xlfile, "BudgetAuthorization", 70, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Fields Options Are NOT Displayed");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userMinimizeBtn));
				userMinimizeBtn.click();
				excelReader.setCellData(xlfile, "BudgetAuthorization", 70, 9, resFail);
				return false;
			}
		}

		
		
	@FindBy(xpath="//span[contains(text(),'Budget Fields')]//i[@class='icon-expand icon-font8']")
	private static WebElement  budgetFieldsExpandBtn;

	@FindBy(xpath="//span[contains(text(),'Budget Fields')]//i[@class='icon-collepse icon-font8']")
	private static WebElement  budgetFieldsMinimizeBtn;



	@FindBy(xpath="//span[contains(text(),'Budget Plan')]")
	private static WebElement  budget_BudgetPlanDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Valid From')]")
	private static WebElement  budget_ValidFromDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Valid To')]")
	private static WebElement  budget_ValidToDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Screen Type')]")
	private static WebElement  budget_ScreenTypeDragandDrop;

	@FindBy(xpath="//li[@class='treeview active']//li[5]//a[1]//span[1]")
	private static WebElement  budget_AccountNameDragandDrop;

	@FindBy(xpath="//li[@class='treeview active']//li[6]//a[1]//span[1]")
	private static WebElement  budget_ProductNameDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Budget Value')]")
	private static WebElement  budget_BudgetValueDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Budget Quantity')]")
	private static WebElement  budget_BudgetQuantityDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Budget Type')]")
	private static WebElement  budget_BudgetTypeDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Last Updated Date')]")
	private static WebElement  budget_LastUpdatedDateDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Version')]")
	private static WebElement  budget_VersionDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Created Date')]")
	private static WebElement  budget_CreatedDateDragandDrop;

	@FindBy(xpath="//li[@class='treeview']//span[contains(text(),'Status')]")
	private static WebElement  budget_StatusDragandDrop;

	@FindBy(xpath="//span[contains(text(),'Narration')]")
	private static WebElement  budget_NarrationDragandDrop;

	@FindBy(xpath="//span[contains(text(),'ToAccount Name')]")
	private static WebElement  budget_ToAccountNameDragandDrop;

	@FindBy(xpath="//span[contains(text(),'ToProduct Name')]")
	private static WebElement  budget_ToProductNameDragandDrop;

		

		
		public boolean 	checkFieldsTabBudgetFieldsOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetFieldsExpandBtn));
			budgetFieldsExpandBtn.click();
			
			Thread.sleep(5000);
			
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budget_BudgetPlanDragandDrop));
		    
		     boolean actbudget_BudgetPlanDragandDrop          = budget_BudgetPlanDragandDrop.isDisplayed();
			 boolean actbudget_ValidFromDragandDrop           = budget_ValidFromDragandDrop.isDisplayed();
			 boolean actbudget_ValidToDragandDrop             = budget_ValidToDragandDrop.isDisplayed();
			 boolean actbudget_ScreenTypeDragandDrop          = budget_ScreenTypeDragandDrop.isDisplayed();
			 boolean actbudget_AccountNameDragandDrop         = budget_AccountNameDragandDrop.isDisplayed();
			 boolean actbudget_ProductNameDragandDrop         = budget_ProductNameDragandDrop.isDisplayed();
			 boolean actbudget_BudgetValueDragandDrop         = budget_BudgetValueDragandDrop.isDisplayed();
			 boolean actbudget_LastUpdatedDateDragandDrop     = budget_LastUpdatedDateDragandDrop.isDisplayed();
			 boolean actbudget_VersionDragandDrop             = budget_VersionDragandDrop.isDisplayed();
			 boolean actbudget_CreatedDateDragandDrop         = budget_CreatedDateDragandDrop.isDisplayed();
			 boolean actbudget_StatusDragandDrop              = budget_StatusDragandDrop.isDisplayed();
			 boolean actbudget_NarrationDragandDrop           = budget_NarrationDragandDrop.isDisplayed();
			 boolean actbudget_ToAccountNameDragandDrop       = budget_ToAccountNameDragandDrop.isDisplayed();
			 boolean actbudget_ToProductNameDragandDrop       = budget_ToProductNameDragandDrop.isDisplayed();
			 
			 boolean expbudget_BudgetPlanDragandDrop          = true;
			 boolean expbudget_ValidFromDragandDrop           = true;
			 boolean expbudget_ValidToDragandDrop             = true;
			 boolean expbudget_ScreenTypeDragandDrop          = true;
			 boolean expbudget_AccountNameDragandDrop         = true;
			 boolean expbudget_ProductNameDragandDrop         = true;
			 boolean expbudget_BudgetValueDragandDrop         = true;
			 boolean expbudget_LastUpdatedDateDragandDrop     = true;
			 boolean expbudget_VersionDragandDrop             = true;
			 boolean expbudget_CreatedDateDragandDrop         = true;
			 boolean expbudget_StatusDragandDrop              = true;
			 boolean expbudget_NarrationDragandDrop           = true;
			 boolean expbudget_ToAccountNameDragandDrop       = true;
			 boolean expbudget_ToProductNameDragandDrop       = true;
			 	
			 System.out.println("****************************** checkFieldsTabBudgetFieldsOptions  ********************************************");
			 
			 System.out.println("budget_BudgetPlanDragandDrop       :"+actbudget_BudgetPlanDragandDrop     +"  Value Expected : "+expbudget_BudgetPlanDragandDrop);
			 System.out.println("budget_ValidFromDragandDrop        :"+actbudget_ValidFromDragandDrop      +"  Value Expected : "+expbudget_ValidFromDragandDrop);
			 System.out.println("budget_ValidToDragandDrop          :"+actbudget_ValidToDragandDrop        +"  Value Expected : "+expbudget_ValidToDragandDrop);
			 System.out.println("budget_ScreenTypeDragandDrop       :"+actbudget_ScreenTypeDragandDrop     +"  Value Expected : "+expbudget_ScreenTypeDragandDrop);
			 System.out.println("budget_AccountNameDragandDrop      :"+actbudget_AccountNameDragandDrop    +"  Value Expected : "+expbudget_AccountNameDragandDrop);
			 System.out.println("budget_ProductNameDragandDrop      :"+actbudget_ProductNameDragandDrop    +"  Value Expected : "+expbudget_ProductNameDragandDrop);
			 System.out.println("budget_BudgetValueDragandDrop      :"+actbudget_BudgetValueDragandDrop    +"  Value Expected : "+expbudget_BudgetValueDragandDrop);
			 System.out.println("budget_LastUpdatedDateDragandDrop  :"+actbudget_LastUpdatedDateDragandDrop+"  Value Expected : "+expbudget_LastUpdatedDateDragandDrop);
			 System.out.println("budget_VersionDragandDrop          :"+actbudget_VersionDragandDrop        +"  Value Expected : "+expbudget_VersionDragandDrop);
			 System.out.println("budget_CreatedDateDragandDrop      :"+actbudget_CreatedDateDragandDrop    +"  Value Expected : "+expbudget_CreatedDateDragandDrop);
			 System.out.println("budget_StatusDragandDrop           :"+actbudget_StatusDragandDrop         +"  Value Expected : "+expbudget_StatusDragandDrop);
			 System.out.println("budget_NarrationDragandDrop        :"+actbudget_NarrationDragandDrop      +"  Value Expected : "+expbudget_NarrationDragandDrop);
			 System.out.println("budget_ToAccountNameDragandDrop    :"+actbudget_ToAccountNameDragandDrop  +"  Value Expected : "+expbudget_ToAccountNameDragandDrop);
			 System.out.println("budget_ToProductNameDragandDrop    :"+actbudget_ToProductNameDragandDrop  +"  Value Expected : "+expbudget_ToProductNameDragandDrop);
				
			 
			 boolean res = actbudget_BudgetPlanDragandDrop==expbudget_BudgetPlanDragandDrop && actbudget_ValidFromDragandDrop==expbudget_ValidFromDragandDrop
					 && actbudget_ValidToDragandDrop==expbudget_ValidToDragandDrop && actbudget_ScreenTypeDragandDrop==expbudget_ScreenTypeDragandDrop
					 && actbudget_AccountNameDragandDrop==expbudget_AccountNameDragandDrop && actbudget_ProductNameDragandDrop==expbudget_ProductNameDragandDrop
					 && actbudget_BudgetValueDragandDrop==expbudget_BudgetValueDragandDrop && actbudget_LastUpdatedDateDragandDrop==expbudget_LastUpdatedDateDragandDrop
					 && actbudget_VersionDragandDrop==expbudget_VersionDragandDrop && actbudget_CreatedDateDragandDrop==expbudget_CreatedDateDragandDrop 
					 && actbudget_StatusDragandDrop==expbudget_StatusDragandDrop && actbudget_NarrationDragandDrop==expbudget_NarrationDragandDrop 
					 && actbudget_ToAccountNameDragandDrop==expbudget_ToAccountNameDragandDrop && actbudget_ToProductNameDragandDrop==expbudget_ToProductNameDragandDrop;
				
			 String actres = Boolean.toString(res);
			 
			 String expres = excelReader.getCellData("BudgetAuthorization", 73, 7);
			 
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 73, 8, actres);
			 
			 
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Field MasterFields Options Are Displayed");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetFieldsMinimizeBtn));
				budgetFieldsMinimizeBtn.click();
				excelReader.setCellData(xlfile, "BudgetAuthorization", 72, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Field MasterFields Options Are NOT Displayed");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetFieldsMinimizeBtn));
				budgetFieldsMinimizeBtn.click();
				excelReader.setCellData(xlfile, "BudgetAuthorization", 72, 9, resFail);
				return false;
			}
		}


		

		public boolean checkEmailTemplateCompanyExpandInFieldsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			 excelReader=new ExcelReader(POJOUtility.getExcelPath());
			 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CompanyExpandBtn));
			 CompanyExpandBtn.click();
			
		
			 Thread.sleep(3000);
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fields_Company_CompanyName));
			 boolean actfields_Company_CompanyName                     = fields_Company_CompanyName.isDisplayed();
		     boolean actfields_Company_CompanyCode                     = fields_Company_CompanyCode.isDisplayed();
			 boolean actfields_Company_AccountingDate                  = fields_Company_AccountingDate.isDisplayed();
			 boolean actfields_Company_Address1                        = fields_Company_Address1.isDisplayed();
			 boolean actfields_Company_Address2                        = fields_Company_Address2.isDisplayed();
			 boolean actfields_Company_City                            = fields_Company_City.isDisplayed();
		     boolean actfields_Company_Zipcode				           = fields_Company_Zipcode.isDisplayed();
			 boolean actfields_Company_Country                         = fields_Company_Country.isDisplayed();
			 boolean actfields_Company_PhoneNumber1                    = fields_Company_PhoneNumber1.isDisplayed();
			 boolean actfields_Company_PhoneNumber2                    = fields_Company_PhoneNumber2.isDisplayed();
			 boolean actfields_Company_Fax                             = fields_Company_Fax.isDisplayed();
			 boolean actfields_Company_TaxRegestrationNumber           = fields_Company_TaxRegestrationNumber.isDisplayed();
			 boolean actfields_Company_ExciseRegistrationNumber        = fields_Company_ExciseRegistrationNumber.isDisplayed();
			 
			 boolean expfields_Company_CompanyName                     = true;
			 boolean expfields_Company_CompanyCode                     = true;
			 boolean expfields_Company_AccountingDate                  = true;
			 boolean expfields_Company_Address1                        = true;
			 boolean expfields_Company_Address2                        = true;
			 boolean expfields_Company_City                            = true;
			 boolean expfields_Company_Zipcode                         = true;
			 boolean expfields_Company_Country                         = true;
			 boolean expfields_Company_PhoneNumber1                    = true;
			 boolean expfields_Company_PhoneNumber2                    = true;
			 boolean expfields_Company_Fax                             = true;
			 boolean expfields_Company_TaxRegestrationNumber           = true;
			 boolean expfields_Company_ExciseRegistrationNumber        = true;
			 
				
			 System.out.println("****************************** checkEmailTemplateCompanyExpandInFieldsTabOptions  ********************************************");
			 
			 System.out.println("fields_Company_CompanyName          :"+actfields_Company_CompanyName              +"    Value Expected : "+expfields_Company_CompanyName);
			 System.out.println("fields_Company_CompanyCode          :"+actfields_Company_CompanyCode              +"    Value Expected : "+expfields_Company_CompanyCode);
			 System.out.println("fields_Company_AccountingDate       :"+actfields_Company_AccountingDate           +"	 Value Expected : "+expfields_Company_AccountingDate);
			 System.out.println("fields_Company_Address1             :"+actfields_Company_Address1                 +"	 Value Expected : "+expfields_Company_Address1);
			 System.out.println("fields_Company_Address2             :"+actfields_Company_Address2                 +"	 Value Expected : "+expfields_Company_Address2);
			 System.out.println("fields_Company_City      	         :"+actfields_Company_City                     +"	 Value Expected : "+expfields_Company_City);
			 System.out.println("fields_Company_Zipcode              :"+actfields_Company_Zipcode                  +"	 Value Expected : "+expfields_Company_Zipcode);
			 System.out.println("fields_Company_Country              :"+actfields_Company_Country                  +"	 Value Expected : "+expfields_Company_Country);
			 System.out.println("fields_Company_PhoneNumber1         :"+actfields_Company_PhoneNumber1             +"	 Value Expected : "+expfields_Company_PhoneNumber1);
			 System.out.println("fields_Company_PhoneNumber2         :"+actfields_Company_PhoneNumber2             +"	 Value Expected : "+expfields_Company_PhoneNumber2);
			 System.out.println("fields_Company_Fax                  :"+actfields_Company_Fax                      +"	 Value Expected : "+expfields_Company_Fax);
			 System.out.println("fields_Company_TaxRegNumber         :"+actfields_Company_TaxRegestrationNumber    +"	 Value Expected : "+expfields_Company_TaxRegestrationNumber);
			 System.out.println("fields_Company_ExciseRegNumber      :"+actfields_Company_ExciseRegistrationNumber +"    Value Expected : "+expfields_Company_ExciseRegistrationNumber);
			 
			 boolean res =  actfields_Company_CompanyName==expfields_Company_CompanyName  && actfields_Company_CompanyCode==expfields_Company_CompanyCode 
					 && actfields_Company_AccountingDate==expfields_Company_AccountingDate
					 && actfields_Company_Address1==expfields_Company_Address1  && actfields_Company_Address2==expfields_Company_Address2
					 && actfields_Company_City==expfields_Company_City && actfields_Company_Zipcode==expfields_Company_Zipcode 
				     && actfields_Company_Country==expfields_Company_Country
				     && actfields_Company_PhoneNumber1==expfields_Company_PhoneNumber1
					 && actfields_Company_PhoneNumber2==expfields_Company_PhoneNumber2
					 && actfields_Company_Fax==expfields_Company_Fax 
					 && actfields_Company_TaxRegestrationNumber==expfields_Company_TaxRegestrationNumber  
					 && actfields_Company_ExciseRegistrationNumber==expfields_Company_ExciseRegistrationNumber;
				
			 String actres = Boolean.toString(res);
			 
			 String expres = excelReader.getCellData("BudgetAuthorization", 75, 7);
			 
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 75, 8, actres);
			 
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Fields Options Are Displayed");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CompanyMinimiseBtn));
				CompanyMinimiseBtn.click();
				excelReader.setCellData(xlfile, "BudgetAuthorization", 74, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Fields Options Are NOT Displayed");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(CompanyMinimiseBtn));
				CompanyMinimiseBtn.click();
				excelReader.setCellData(xlfile, "BudgetAuthorization", 74, 9, resFail);
				return false;
			}
		}
		


	    public boolean 	checkFieldsTabMiscellaneousOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	    {
		
		
		    excelReader=new ExcelReader(POJOUtility.getExcelPath());
		    xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		    Thread.sleep(3000);
		
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fields_MiscellaneousExpandBtn));
		    fields_MiscellaneousExpandBtn.click();
		
		    Thread.sleep(2000);
		
		    boolean actmiscellaneous_Logindate             = miscellaneous_Logindate.isDisplayed();
			boolean actmiscellaneous_PageNumber            = miscellaneous_PageNumber.isDisplayed();
			boolean actmiscellaneous_Printdate             = miscellaneous_Printdate.isDisplayed();
			boolean actmiscellaneous_PrintTime             = miscellaneous_PrintTime.isDisplayed();
			boolean actmiscellaneous_TotalPage             = miscellaneous_TotalPage.isDisplayed();
			boolean actmiscellaneous_Sno                   = miscellaneous_Sno.isDisplayed();
			boolean actmiscellaneous_PageType              = miscellaneous_PageType.isDisplayed();
			boolean actmiscellaneous_AgeingDetails         = miscellaneous_AgeingDetails.isDisplayed();
			boolean actmiscellaneous_AgeingHeadings        = miscellaneous_AgeingHeadings.isDisplayed();
			boolean actmiscellaneous_Ageingbalance         = miscellaneous_Ageingbalance.isDisplayed();
			boolean actmiscellaneous_AgeingSlab1           = miscellaneous_AgeingSlab1.isDisplayed();
			boolean actmiscellaneous_AgeingSlab2           = miscellaneous_AgeingSlab2.isDisplayed();
			boolean actmiscellaneous_AgeingSlab3           = miscellaneous_AgeingSlab3.isDisplayed();
			boolean actmiscellaneous_AgeingSlab4           = miscellaneous_AgeingSlab4.isDisplayed();
			boolean actmiscellaneous_AgeingSlab5           = miscellaneous_AgeingSlab5.isDisplayed();
			boolean actmiscellaneous_AgeingSlab6           = miscellaneous_AgeingSlab6.isDisplayed();
			boolean actmiscellaneous_AgeingMonthWise       = miscellaneous_AgeingMonthWise.isDisplayed();
			
			boolean expmiscellaneous_Logindate             = true;
			boolean expmiscellaneous_PageNumber            = true;
			boolean expmiscellaneous_Printdate             = true;
			boolean expmiscellaneous_PrintTime             = true;
			boolean expmiscellaneous_TotalPage             = true;
			boolean expmiscellaneous_Sno                   = true;
			boolean expmiscellaneous_PageType              = true;
			boolean expmiscellaneous_AgeingDetails         = true;
			boolean expmiscellaneous_AgeingHeadings        = true;
			boolean expmiscellaneous_Ageingbalance         = true;
			boolean expmiscellaneous_AgeingSlab1           = true;
			boolean expmiscellaneous_AgeingSlab2           = true;
			boolean expmiscellaneous_AgeingSlab3           = true;
			boolean expmiscellaneous_AgeingSlab4           = true;
			boolean expmiscellaneous_AgeingSlab5           = true;
			boolean expmiscellaneous_AgeingSlab6           = true;
			boolean expmiscellaneous_AgeingMonthWise       = true;
			
			System.out.println("****************************** checkToolsOptionsDepartmentMenu  ********************************************");
		
			System.out.println("miscellaneous_Logindate           :"+actmiscellaneous_Logindate       +"Value Expected : "+expmiscellaneous_Logindate);
			System.out.println("miscellaneous_PageNumber          :"+actmiscellaneous_PageNumber      +"Value Expected : "+expmiscellaneous_PageNumber);
			System.out.println("miscellaneous_Printdate           :"+actmiscellaneous_Printdate       +"Value Expected : "+expmiscellaneous_Printdate);
			System.out.println("miscellaneous_PrintTime           :"+actmiscellaneous_PrintTime       +"Value Expected : "+expmiscellaneous_PrintTime);
			System.out.println("miscellaneous_TotalPage           :"+actmiscellaneous_TotalPage       +"Value Expected : "+expmiscellaneous_TotalPage);
			System.out.println("miscellaneous_Sno                 :"+actmiscellaneous_Sno             +"Value Expected : "+expmiscellaneous_Sno);
			System.out.println("miscellaneous_PageType            :"+actmiscellaneous_PageType        +"Value Expected : "+expmiscellaneous_PageType);
			System.out.println("miscellaneous_AgeingDetails       :"+actmiscellaneous_AgeingDetails   +"Value Expected : "+expmiscellaneous_AgeingDetails);
			System.out.println("miscellaneous_AgeingHeadings      :"+actmiscellaneous_AgeingHeadings  +"Value Expected : "+expmiscellaneous_AgeingHeadings);
			System.out.println("miscellaneous_Ageingbalance       :"+actmiscellaneous_Ageingbalance   +"Value Expected : "+expmiscellaneous_Ageingbalance);
			System.out.println("miscellaneous_AgeingSlab1         :"+actmiscellaneous_AgeingSlab1     +"Value Expected : "+expmiscellaneous_AgeingSlab1);
			System.out.println("miscellaneous_AgeingSlab2         :"+actmiscellaneous_AgeingSlab2     +"Value Expected : "+expmiscellaneous_AgeingSlab2);
			System.out.println("miscellaneous_AgeingSlab3         :"+actmiscellaneous_AgeingSlab3     +"Value Expected : "+expmiscellaneous_AgeingSlab3);
			System.out.println("miscellaneous_AgeingSlab4         :"+actmiscellaneous_AgeingSlab4     +"Value Expected : "+expmiscellaneous_AgeingSlab4);
			System.out.println("miscellaneous_AgeingSlab5         :"+actmiscellaneous_AgeingSlab5     +"Value Expected : "+expmiscellaneous_AgeingSlab5);
			System.out.println("miscellaneous_AgeingSlab6         :"+actmiscellaneous_AgeingSlab6     +"Value Expected : "+expmiscellaneous_AgeingSlab6);
			System.out.println("miscellaneous_AgeingMonthWise     :"+actmiscellaneous_AgeingMonthWise +"Value Expected : "+expmiscellaneous_AgeingMonthWise);
			
			
			boolean res = actmiscellaneous_Logindate==expmiscellaneous_Logindate  && actmiscellaneous_PageNumber==expmiscellaneous_PageNumber
					&& actmiscellaneous_Printdate==expmiscellaneous_Printdate && actmiscellaneous_PrintTime==expmiscellaneous_PrintTime
					&& actmiscellaneous_TotalPage==expmiscellaneous_TotalPage && actmiscellaneous_PageType==expmiscellaneous_PageType 
					&& actmiscellaneous_Sno==expmiscellaneous_Sno && actmiscellaneous_AgeingDetails==expmiscellaneous_AgeingDetails
					&& actmiscellaneous_AgeingDetails==expmiscellaneous_AgeingDetails && actmiscellaneous_AgeingSlab1==expmiscellaneous_AgeingSlab1
					&& actmiscellaneous_Ageingbalance==expmiscellaneous_Ageingbalance && actmiscellaneous_AgeingSlab2==expmiscellaneous_AgeingSlab2
					&& actmiscellaneous_AgeingSlab3==expmiscellaneous_AgeingSlab3 && actmiscellaneous_AgeingSlab4==expmiscellaneous_AgeingSlab4
		            && actmiscellaneous_AgeingSlab5==expmiscellaneous_AgeingSlab5 && actmiscellaneous_AgeingSlab6==expmiscellaneous_AgeingSlab6 
		            && actmiscellaneous_AgeingMonthWise==expmiscellaneous_AgeingMonthWise;
				
			 String actres = Boolean.toString(res);
			 
			 String expres = excelReader.getCellData("BudgetAuthorization", 77, 7);
			 
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 77, 8, actres);
			 
			
			if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Field Miscellaneous Options Are Displayed");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fields_MinimizeBtn));
				fields_MinimizeBtn.click();
				excelReader.setCellData(xlfile, "BudgetAuthorization", 76, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Field Miscellaneous Options Are NOT Displayed");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fields_MinimizeBtn));
				fields_MinimizeBtn.click();
				
				excelReader.setCellData(xlfile, "BudgetAuthorization", 76, 9, resFail);
				return false;
			}
	    }



		
		
		public boolean 	checkDesignAreaPrintScreen() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			
			Thread.sleep(3000);
			
			boolean actheaderDesignerTab=headerDesignerTab.isDisplayed();
			boolean expheaderDesignerTab=true;
			
			System.out.println("**************************************checkDesignAreaPrintScreen*******************************************************");
			
			System.out.println("headerDesignerTab     : "+actheaderDesignerTab+"  Value Expected  : "+expheaderDesignerTab);
			
			
			boolean res = actheaderDesignerTab==expheaderDesignerTab;
				
			 String actres = Boolean.toString(res);
			 
			 String expres = excelReader.getCellData("BudgetAuthorization", 79, 7);
			 
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 79, 8, actres);
			 
			if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Design Area is Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 78, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Design Area is NOT Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 78, 9, resFail);
				return false;
			}
		}




	public boolean checkFieldsDragInHeaderOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userExpandBtn));
			userExpandBtn.click();
			
			
		    Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\print.exe");
		 
			Thread.sleep(100000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(headerPropertiesTab));
			headerPropertiesTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Savedropdown));
			Savedropdown.click();
			
			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedEmailTemplatePrintSaveBtn));
			receivedEmailTemplatePrintSaveBtn.click();
			
			Thread.sleep(3000);
			
			saveLayoutName.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveLayoutName));
			saveLayoutName.sendKeys(excelReader.getCellData("BudgetAuthorization", 80, 6));
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveLayoutOkBtn));
			saveLayoutOkBtn.click();
			
			String expMessage=excelReader.getCellData("BudgetAuthorization", 81, 7);
			
			String actMessage=checkValidationMessage(expMessage);
			
			excelReader.setCellData(xlfile, "BudgetAuthorization", 81, 8, actMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 80, 9, resPass);
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 80, 9, resFail);
				return false;
			}
			
		
	}	
		



	@FindBy(xpath="//i[@class='icon-font6 icon-save']")
	private static WebElement  Savedropdown;

		@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[1]")
		private static WebElement  PrintSaveBtn;
		
		@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[2]")
		private static WebElement  SaveAsBtn; 
		
		@FindBy(xpath="//input[@id='id_PopuplayoutName']")
		private static WebElement  SaveTxtField;
		
		@FindBy(xpath="//input[@id='id_PopuplayoutName']")
		private static WebElement  Save_OkBtn;
		
		@FindBy(xpath="//*[@id='id_PopUp']/div/div/div[3]/button[2]")
		private static WebElement  Save_CancekBtn;
		
		@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[3]")
		private static WebElement  SaveAsImageBtn; //No Response
		
		@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[4]")
		private static WebElement  SaveAsXMLBtn; // Downloads xml file
		
		@FindBy(xpath="//*[@id='forHf']/span[3]/ul/li/a[5]")
		private static WebElement  SaveAsHTMLBtn; 

		@FindBy(xpath="//span[@class='dropdown theme_icon-color open']//a[1]")
		private static WebElement receivedEmailTemplatePrintSaveBtn;	
		
		@FindBy(xpath="//input[@id='id_PopuplayoutName']")
		private static WebElement saveLayoutName;	
			
		@FindBy(xpath="//button[@id='btn_PopUpOk']")
		private static WebElement saveLayoutOkBtn;	
		
		@FindBy(xpath="//div[@id='id_PopUp']//button[@class='Fbutton'][contains(text(),'Cancel')]")
		private static WebElement saveLayoutCancelBtn;	
		
		@FindBy(xpath="//a[contains(text(),'Static Text')]")
		private static WebElement controlStaticText;	


		
		
		
		@FindBy(xpath="//*[@id='id_InvoiceDesingCancel']/i")
		private static WebElement  ExitBtn;	
		
		
		
		public boolean checkExitOptionInLayout() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ExitBtn));
			ExitBtn.click();
			
	        Thread.sleep(2000);
				      
	        
			 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedChkBox));
			 boolean actreceivedChkBox                   = receivedChkBox.isDisplayed();
		     boolean actapprovedChkBox                   = approvedChkBox.isDisplayed();
			 boolean actrejectedChkBox                   = rejectedChkBox.isDisplayed();
			 boolean actemailHeaderTemplateBtn           = emailHeaderTemplateBtn.isDisplayed();
			 boolean actemailTemplateBtn                 = emailTemplateBtn.isDisplayed();
			 boolean actheaderAlert                      = headerAlert.isDisplayed();
			 
			 boolean expreceivedChkBox                   = true;
			 boolean expapprovedChkBox                   = true;
			 boolean exprejectedChkBox                   = true;
			 boolean expemailHeaderTemplateBtn           = true;
			 boolean expemailTemplateBtn                 = true;
			 boolean expheaderAlert                      = true;
			 
			 boolean res = actreceivedChkBox==expreceivedChkBox&& actapprovedChkBox==expapprovedChkBox  && actrejectedChkBox==exprejectedChkBox 
					  && actemailHeaderTemplateBtn==expemailHeaderTemplateBtn 
					  && actemailTemplateBtn==expemailTemplateBtn && actheaderAlert==expheaderAlert ;
				
			 String actres = Boolean.toString(res);
			 
			 String expres = excelReader.getCellData("BudgetAuthorization", 83, 7);
			 
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 83, 8, actres);
			 	
			 System.out.println("******************************************* checkExitOptionInLayout  ********************************************");
			 
			 System.out.println("ReceivedChkBox                :"+actreceivedChkBox         +"   Value Expected : "+expreceivedChkBox);
			 System.out.println("ApprovedChkBox                :"+actapprovedChkBox         +"   Value Expected : "+expapprovedChkBox);
			 System.out.println("RejectedChkBox                :"+actrejectedChkBox         +"	 Value Expected : "+exprejectedChkBox);
			 System.out.println("EmailHeaderTemplateBtn        :"+actemailHeaderTemplateBtn +"	 Value Expected : "+expemailHeaderTemplateBtn);
			 System.out.println("EmailTemplateBtn      	       :"+actemailTemplateBtn       +"	 Value Expected : "+expemailTemplateBtn);
			 System.out.println("HeaderAlert                   :"+actheaderAlert            +"	 Value Expected : "+expheaderAlert);
			  	
			 if(actres.equalsIgnoreCase(expres))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 82, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 82, 9, resFail);
				return false;
			}
		}
		
		
		@FindBy(xpath="//span[@id='btnNewSmsTemplate']")
		private static WebElement smsTemplateLayoutNewBtn;	
		
		@FindBy(xpath="//span[@id='btnSaveSmsTemplate']")
		private static WebElement smsTemplateLayoutSaveBtn;	

		@FindBy(xpath="//span[@id='btnCloseSMSModal']")
		private static WebElement smsTemplateLayoutCloseBtn;	
			
		@FindBy(xpath="//span[@class='FBold theme_icon-color'][contains(text(),'Users')]")
		private static WebElement smsTemplateLayoutUsers;	
		
		@FindBy(xpath="//span[@class='FBold theme_icon-color'][contains(text(),'Company')]")
		private static WebElement smsTemplateLayoutCompany;	
		
		@FindBy(xpath="//span[contains(text(),'Body')]")
		private static WebElement smsTemplateLayoutBody;	
		
		@FindBy(xpath="//span[contains(text(),'Header')]")
		private static WebElement smsTemplateLayoutHeader;	
		
		@FindBy(xpath="//td[@id='id_tblAlerts_col_1-3']")
		private static WebElement receivedEmailTemplateClick;	
		
		@FindBy(xpath="//select[@id='id_tblAlerts_EmailTemplate']")
		private static WebElement receivedEmailTemplateInput;
		
		

		public boolean checkSavedEmailTemplateInEmailTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recievedEmailHeaderTemplate));
			recievedEmailHeaderTemplate.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedEmailTemplateClick));
			receivedEmailTemplateClick.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recievedEmailHeaderTemplate));
			recievedEmailHeaderTemplate.click();
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedEmailTemplateClick));
			receivedEmailTemplateClick.click();
			
			
			Select s=new Select(receivedEmailTemplateInput);
			s.selectByVisibleText(excelReader.getCellData("BudgetAuthorization", 84, 6));
		
		
			
			receivedEmailTemplateInput.click();
			
			
			
			String actreceivedEmailTemplateInput=receivedEmailTemplateInput.getText();
			
			String expreceivedEmailTemplateInput=excelReader.getCellData("BudgetAuthorization", 85, 7);
			
			 excelReader.setCellData(xlfile, "BudgetAuthorization", 85, 8, actreceivedEmailTemplateInput);
			
			System.out.println("******************************************************checkSavedEmailTemplateInEmailTemplate*******************************"); 
			
			
			System.out.println("receivedEmailTemplateInput     : "+actreceivedEmailTemplateInput+"   VAlue Expected : "+expreceivedEmailTemplateInput); 
			
			if(actreceivedEmailTemplateInput.equalsIgnoreCase(expreceivedEmailTemplateInput))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 84, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				excelReader.setCellData(xlfile, "BudgetAuthorization", 84, 9, resFail);
				return false;
			}
			
		}
		
		
		
		@FindBy(xpath="//td[@id='id_tblAlerts_col_1-4']")
		private static WebElement receivedSMSTemplateClick;	
		
		@FindBy(xpath="//select[@id='id_tblAlerts_SMSTemplate']")
		private static WebElement smsTemplateInput;	
		
		
		@FindBy(xpath="//td[@id='id_tblAlerts_col_1-4']")
		private static WebElement alertsTemplateClick;	
		
		@FindBy(xpath="//select[@id='id_tblAlerts_Alerts']")
		private static WebElement alertsTemplateInput;	
		

		
		
		public boolean checkSavedAlertTemplateInTemplate() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(headerAlert));
			headerAlert.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyMenu));
			emailTemplates.sendKeys(excelReader.getCellData("BudgetAuthorization", 86, 6));
			Thread.sleep(2000);
			emailTemplates.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyNameTemplateArea));
			companyNameTemplateArea.click();
			companyNameTemplateArea.sendKeys("Alerts Template");
			companyNameTemplateArea.sendKeys(excelReader.getCellData("BudgetAuthorization", 87, 6));
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtnInEmailTemplateHeader));
			saveBtnInEmailTemplateHeader.click();
			
			String actValidationMsg=checkValidationMessage("Template saved successfully");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(closeButtonInCreateTemplate));
			closeButtonInCreateTemplate.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(recievedEmailHeaderTemplate));
			recievedEmailHeaderTemplate.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedEmailTemplateClick));
			receivedEmailTemplateClick.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTemplateClick));
			alertsTemplateClick.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTemplateInput));
			alertsTemplateInput.click();
			
			Thread.sleep(3000);
			
			
			Select s=new Select(alertsTemplateInput);
			s.selectByVisibleText("createBudgetAlertsTemplate");
		
			Thread.sleep(3000);
			
			alertsTemplateInput.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(receivedEmailTemplateClick));
			receivedEmailTemplateClick.click();
			
			
			String actalertsTemplateInput=alertsTemplateClick.getText();
			
			
			String expalertsTemplateInput="createBudgetAlertsTemplate";
			
			
			System.out.println("*********************************checkSavedAlertTemplateInTemplate*****************************"); 
		
			
			System.out.println("AlertsTemplateInput     : "+actalertsTemplateInput +"  Value Expected : "+expalertsTemplateInput);
		
			
			if(actalertsTemplateInput.equalsIgnoreCase(expalertsTemplateInput))
			{
				System.out.println("Test Pass : Alerts TAb Options Are Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 124, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Alerts TAb Options Are NOT Displayed");
				excelReader.setCellData(xlfile, xlSheetName, 124, 9, resFail);
				return false;
			}
		}	
		
		
		
		
		
		
		public boolean checkSaveButtonWithCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			System.out.println("********************checkSaveButtonWithCondition*******************************");
			
			boolean actendpoint=endpoint.isDisplayed();
			boolean expendpoint=true;
			
			System.out.println("Endpoint   : "+actendpoint +"  Value Expected :"+expendpoint);
			
			if(actendpoint==expendpoint)
			{
				System.out.println("Test Pass : Image Displayed In Design Area");
				excelReader.setCellData(xlfile, xlSheetName, 125, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area");
				excelReader.setCellData(xlfile, xlSheetName, 125, 9, resFail);
				return false;
			}
			
			
		}
		
		
		@FindBy(xpath="//i[@class='icon-font7 icon-save']")
		private static WebElement saveBtnAuthorization;	
		
		
		public boolean checkSaveAuthorizationWithConditionAndAlerts() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
			
			Thread.sleep(5000);
			

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			System.out.println("***********************************checkSaveAuthorizationWithConditionAndAlerts****************************************************"); 
			
			String expMessage="Authorization Flow saved successfully.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				excelReader.setCellData(xlfile, xlSheetName, 126, 9, resPass);
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				excelReader.setCellData(xlfile, xlSheetName, 126, 9, resFail);
				return false;
			}
		}
		
		

		public boolean 	checkClearButton() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(clearbtnInAuth));
			clearbtnInAuth.click();
			
			Thread.sleep(2000);
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			String actbudgetDropdown=budgetDropdown.getAttribute("placeholder");
			String expbudgetDropdown="Budget Type";
			
		    System.out.println("******************************checkClearButton********************");
			
			System.out.println("budgetDropdown    : "+actbudgetDropdown+"   Value Expected : "+expbudgetDropdown); 
		
			
			if( actbudgetDropdown.equalsIgnoreCase(expbudgetDropdown))
			{
				System.out.println("Test Pass : Master Dropdown Options Are As expected");
				excelReader.setCellData(xlfile, xlSheetName, 127, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Master Dropdown Options Are NOT As expected");
				excelReader.setCellData(xlfile, xlSheetName, 127, 9, resFail);
				return false;
			}
			
		}
		
		
		
		@FindBy(xpath="//li[contains(text(),'Warehouse')]")
		private static WebElement selectWarehouse;	
		
		
		@FindBy(xpath="//span[@id='spnAddLevel']//i[@class='icon-new icon-font7']")
		private static WebElement addLevelButton;	

		
		public boolean checkAlertMessageOnDeleteAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			budgetDropdown.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME, Keys.BACK_SPACE);
			
			budgetDropdown.sendKeys("Revise Budget");
					
			Thread.sleep(2000);
			
			budgetDropdown.sendKeys(Keys.TAB);
			
			
			nameTxt.sendKeys("Revise");
			
			
			nameTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(3000);
			
			addLevelButton.click();
			
			Thread.sleep(3000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			descriptionTxt.sendKeys("Revise");

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userDropdown));
			userDropdown.click();
			
			userDropdown.sendKeys("UserAllOptionsST");
			
			Thread.sleep(3000);
			
			userDropdown.sendKeys(Keys.TAB);
						
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
			errorMessageCloseBtn.click();
			
			Thread.sleep(5000);
			
	        nameTxt.sendKeys("Revise");
			
			Thread.sleep(3000);
			
			nameTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
			deleteBtn.click();
						
			//Working with delete option
			getWaitForAlert();
			
			System.out.println("***********************************checkAlertMessageOnDeleteAuthorization**************************"); 
			
			String actAlert=this.getDriver().switchTo().alert().getText();
			String expAlert="Are you sure that you want to delete?";
		

			System.out.println(" Alert Message Displayed : "+actAlert +"  Value Expected  : "+expAlert);
			
			if(actAlert.equalsIgnoreCase(expAlert))
			{
				System.out.println("Error Message is As Expected");
				excelReader.setCellData(xlfile, xlSheetName, 128, 9, resPass);
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				excelReader.setCellData(xlfile, xlSheetName, 128, 9, resFail);
				return false;
			}
		}
		
		
		
			
		public boolean 	checkCancelInAlertDeleteOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getAlert().dismiss();
			
			Thread.sleep(3000);
			
			String actdata,expdata;
			
			String actnameTxt=nameTxt.getAttribute("value");
			String expnameTxt="Revise";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();

			System.out.println("****************************checkCancelInAlertDeleteOption**********"); 
			
			System.out.println("NameTxt  : "+actnameTxt +"  Value Expected : "+expnameTxt);
			
			if(actnameTxt.equalsIgnoreCase(expnameTxt))
			{
				excelReader.setCellData(xlfile, xlSheetName, 129, 9, resPass);
				return true;
			}
			else
			{
				excelReader.setCellData(xlfile, xlSheetName, 129, 9, resFail);
				return false;
			}
			
		}
		
		
		

		public boolean 	checkOkInAlertDeleteOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
			deleteBtn.click();
			
			
			//Working with delete option
			
			getWaitForAlert();
			
			getAlert().accept();
			
			String expMessage="Selected Authorization flow deleted successfully.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			String actnameTxt=nameTxt.getAttribute("placeholder");
			String expnameTxt="Add New Authorization Flow";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();

			
			System.out.println("*************************************checkOkInAlertDeleteOption************************");
			
			System.out.println("nameTxt     : "+actnameTxt +"  Value Expected : "+expnameTxt);
			System.out.println("Message     : "+actMessage +"  Value Expected : "+expMessage);
			
			if(actnameTxt.equalsIgnoreCase(expnameTxt) && actMessage.equalsIgnoreCase(expMessage))
			{
				excelReader.setCellData(xlfile, xlSheetName, 130, 9, resPass);
				return true;
	         		
			}
			else
			{
				excelReader.setCellData(xlfile, xlSheetName, 130, 9, resFail);
				return false;
			}
			
		}
		
		
		
	@FindBy(xpath="//*[@id='userprofile']/li/span[2]")
	private static WebElement logoutOption;

	/*@FindBy(xpath="//*[@id='txtUsername']")
	private static WebElement username;*/

	@FindBy(id="txtPassword")
	private static WebElement password;

	@FindBy(id="ddlCompany")
	private static WebElement companyDropDownList;

	@FindBy(id="btnSignin")
	private static WebElement signIn;


		
		public boolean checkCloseAndLogoutTheCompany() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			   excelReader=new ExcelReader(POJOUtility.getExcelPath());
			   xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
			   cancelBtn.click();
				
			   Thread.sleep(3000);
			  
			   getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
			   userNameDisplay.click();

			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			   logoutOption.click();
	  
				boolean verifyFiledsStatus=true;
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
				
				boolean actusername                 = username.isDisplayed();
				boolean actcompanyDropDownList      = companyDropDownList.isDisplayed();
				boolean actsignIn                   = signIn.isDisplayed();
				
				boolean expusername                 = true;
				boolean expcompanyDropDownList      = true;
				boolean expsignIn                   = true;
				
				System.out.println("*******************************************checkCloseAndLogoutTheCompany*************************");
				
				System.out.println("username              : "+actusername            +"  Value Expected : "+expusername);
				System.out.println("companyDropDownList   : "+actcompanyDropDownList +"  Value Expected : "+expcompanyDropDownList);
				System.out.println("signIn                : "+actsignIn              +"  Value Expected : "+expsignIn);
				
				if(actusername==expusername&& password.isDisplayed()&&actcompanyDropDownList==expcompanyDropDownList&& actsignIn==expsignIn)
				{
									
					excelReader.setCellData(xlfile, xlSheetName, 76, 9, resPass);
				    return true;
							 
				}
				else
				{
					
					excelReader.setCellData(xlfile, xlSheetName, 76, 9, resFail);
					return false;
			    }
		}
		
		
			
		public boolean 	checkConditionnotRequiredascheckedandSavebuttonwithManadatoryFields() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();

			nameTxt.sendKeys("AuthWithNoCondition");
			nameTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			getAction().doubleClick(startPointBtn).build().perform();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
			conditionTab.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			
			descriptionTxt.sendKeys("AuthWithNoCondition");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();
			
			Thread.sleep(2000);
			
			boolean actconditionNotRequiredChkBox=conditionNotRequiredChkBox.isSelected();
			boolean expconditionNotRequiredChkBox=true;
			
			System.out.println("**********************************************checkConditionnotRequiredascheckedandSavebuttonwithManadatoryFields**********");		
			System.out.println("ConditionNotRequiredChkBox      : "+actconditionNotRequiredChkBox +"  Value Expected : "+expconditionNotRequiredChkBox);
			
			if(actconditionNotRequiredChkBox==expconditionNotRequiredChkBox)
			{
				System.out.println("Test Pass : Successfully created AuthWithNoCondition");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT created AuthWithNoCondition");
				return false;
			}
			
		}
		 
			
			
		public boolean 	checkselectingInactiveintheActionofEntryStatus() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actionOnEntryStatusDropdown));
			actionOnEntryStatusDropdown.click();
			
			Select s=new Select(actionOnEntryStatusDropdown);
			s.selectByVisibleText("Inactive");
			
			String actactionOnEntryStatusDropdown=actionOnEntryStatusDropdown.getAttribute("value");
			String expactionOnEntryStatusDropdown="1";
			
			System.out.println("*****************************checkselectingInactiveintheActionofEntryStatus*************************");
			
			System.out.println("actionOnEntryStatusDropdown     : "+actactionOnEntryStatusDropdown +"  Value Expected : "+expactionOnEntryStatusDropdown);
			
			if(actactionOnEntryStatusDropdown.equals(expactionOnEntryStatusDropdown))
			{
				System.out.println("Test Pass : Selected Active Successfully");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Selected Active Successfully");
				return false;
			}
		}
			
			
			
		public boolean 	checkselectinguserFromUserSelectiontab() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userdropdown));
			userdropdown.click();
			
			Select s=new Select(userdropdown);
			s.selectByVisibleText("ProductTestUser");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			
			System.out.println("*******************************checkselectinguserFromUserSelectiontab**************************");
			boolean actendpoint=endpoint.isDisplayed();
			boolean expendpoint=true;
			
			System.out.println("Endpoint   : "+actendpoint +"  Value Expected :"+expendpoint);
			
			if(actendpoint==expendpoint)
			{
				System.out.println("Test Pass : Image Displayed In Design Area");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area");
				return false;
			}
		}
		
		
		@FindBy(xpath="//div[@id='divLevel1']//div[@class='row clsDivAuthorizationName']")
		private static WebElement level1;
		
		
		
		public boolean 	checkMasterAuthorizationAreaLevel1CheckwithDescriptionNameasAuthWithNoCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(level1));
			
			String actdata,expdata;
			
			String actlevel1=level1.getText();
			String explevel1="AuthWithNoCondition";

			System.out.println("*************************checkMasterAuthorizationAreaLevel1CheckwithDescriptionNameasAuthWithNoCondition*******************");
			
			System.out.println("level1   : "+actlevel1 +"  Value Expected : "+explevel1);
			
			if(actlevel1.equalsIgnoreCase(explevel1))
			{
				System.out.println("Test Pass : level 1 is As Expected");
				return true;
			}
			else
			{
				System.out.println("Test Fail : level 1 is NOT As Expected");
				return false;
			}
			
		}
		
		
		public boolean 	checkSavingMasterAuthorizationAuthWithNoCondition() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{		
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			String expMessage="Authorization Flow saved successfully.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("***************************checkSavingMasterAuthorizationAuthWithNoCondition**********************");
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
				cancelBtn.click();
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
				cancelBtn.click();
				return false;
			}
		}
		
		
		
		
		public boolean 	checkMasterAuthorizationasAuthDateRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();

			nameTxt.sendKeys("AccAuthDateRange");
			nameTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(activeChkBox));
			activeChkBox.click();
			
			
			boolean actactiveChkBox=activeChkBox.isSelected();
			boolean expactiveChkBox=true;
			
			
			String actnameTxt=nameTxt.getAttribute("value");
			String expnameTxt="AccAuthDateRange";
			
			
			System.out.println("*************************checkMasterAuthorizationasAuthDateRange******************************");
			
			System.out.println("nameTxt        : "+actnameTxt       +"  Value Expected  : "+expnameTxt);
			System.out.println("activeChkBox   : "+actactiveChkBox  +"  Value Expected  : "+expactiveChkBox);
			
			if(actnameTxt.equalsIgnoreCase(expnameTxt) && actactiveChkBox==expactiveChkBox)
			{
				System.out.println("Test Pass : Successfully inputed AccAuthDateRange");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT inputed AccAuthDateRange");
				return false;
			}
		}
		

		
		public boolean 	checkStartPointAccAuthDateRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			getAction().doubleClick(startPointBtn).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
	        boolean actconditionTab                 = conditionTab.isDisplayed();
			boolean actuserSelectionTab             = userSelectionTab.isDisplayed();
			boolean actalertsTab                    = alertsTab.isDisplayed();
			boolean actescalationTab                = escalationTab.isDisplayed();
			boolean actdefinitionOkBtn              = definitionOkBtn.isDisplayed();
			boolean actdefinitionCancelBtn          = definitionCancelBtn.isDisplayed();
			
			boolean expconditionTab                 = true;
			boolean expuserSelectionTab             = true;
			boolean expalertsTab                    = true;
			boolean expescalationTab                = true;
			boolean expdefinitionOkBtn              = true;
			boolean expdefinitionCancelBtn          = true;
			
			System.out.println("********************* checkStartPointAccAuthDateRange ******************");
		
			System.out.println("ConditionTab             :"+actconditionTab           +"    Value Expected : "+expconditionTab);
			System.out.println("UserSelectionTab         :"+actuserSelectionTab       +"    Value Expected : "+expuserSelectionTab);
			System.out.println("AlertsTab                :"+actalertsTab              +"	Value Expected : "+expalertsTab);
			System.out.println("EscalationTab            :"+actescalationTab          +"	Value Expected : "+expescalationTab);
			System.out.println("DefinitionOkBtn          :"+actdefinitionOkBtn        +"	Value Expected : "+expdefinitionOkBtn);
			System.out.println("DefinitionCancelBtn      :"+actdefinitionCancelBtn    +"	Value Expected : "+expdefinitionCancelBtn);
			
			if( actconditionTab==expconditionTab  && actuserSelectionTab==expuserSelectionTab &&
					actalertsTab==expalertsTab && actescalationTab==expescalationTab && actdefinitionCancelBtn==expdefinitionCancelBtn && 
					actdefinitionOkBtn==expdefinitionOkBtn)
				
			{
				System.out.println("Test Pass : Start Point works As Expected");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Start Point NOT works As Expected");
				return false;
			}
		}
		
		
		
		
		public boolean 	checkInputingAccWithDateRangeinDescription() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
			conditionTab.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			
			descriptionTxt.sendKeys("AccWithDateRange");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dateRangeChkBox));
			dateRangeChkBox.click();
			
			Thread.sleep(2000);
			
			
			boolean actconditionNotRequiredChkBox     = conditionNotRequiredChkBox.isSelected();
			boolean actdateRangeChkBox                = dateRangeChkBox.isSelected();
			
			boolean expconditionNotRequiredChkBox     = true;
			boolean expdateRangeChkBox                = true;
			
			
			System.out.println("***************************************checkInputingAccWithDateRangeinDescription****************************");
			
			System.out.println("conditionNotRequiredChkBox     : "+actconditionNotRequiredChkBox   +"  Value Expected  : "+expconditionNotRequiredChkBox);
			System.out.println("dateRangeChkBox                : "+actdateRangeChkBox              +"  Value Expected  : "+expdateRangeChkBox);
			
			if(actconditionNotRequiredChkBox==expconditionNotRequiredChkBox && actdateRangeChkBox==expdateRangeChkBox)
			{
				System.out.println("Test Pass : Successfully created AccWithDateRange");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT created AccWithDateRange");
				return false;
			}
			
		}
		
		
		
		public boolean 	checkselectingInActiveintheActionofEntryStatusAccWithDateRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actionOnEntryStatusDropdown));
			actionOnEntryStatusDropdown.click();
			
			Select s=new Select(actionOnEntryStatusDropdown);
			s.selectByVisibleText("Inactive");
			
			String actactionOnEntryStatusDropdown=actionOnEntryStatusDropdown.getAttribute("value");
			String expactionOnEntryStatusDropdown="1";
			
			System.out.println("****************************************checkselectingInActiveintheActionofEntryStatusAccWithDateRange********************");
			
			System.out.println("actionOnEntryStatusDropdown    : "+actactionOnEntryStatusDropdown+"  Value Expected : "+expactionOnEntryStatusDropdown);	
			
			if(actactionOnEntryStatusDropdown.equals(expactionOnEntryStatusDropdown))
				
			{
				System.out.println("Test Pass : Selected Active Successfully");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Selected Active Successfully");
				return false;
			}
		}
		
		
		public boolean 	checkselectinguserFromUserSelectiontabAccWithDateRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userdropdown));
			userdropdown.click();
			
			Select s=new Select(userdropdown);
			s.selectByVisibleText("ProductTestUser");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(level1));
			
			
		    String 	actlevel1=level1.getText();
		    String 	explevel1="AccWithDateRange";
		    
		    boolean actendpoint=endpoint.isDisplayed();
		    boolean expendpoint = true;
		    
		    
		    System.out.println("*************************checkselectinguserFromUserSelectiontabAccWithDateRange*******");
			
		    System.out.println("level1   : "+actlevel1     +"  Value Expected : "+explevel1);
		    System.out.println("Endpoint   : "+actendpoint +"  Value Expected : "+expendpoint);
		    
			if(actlevel1.equalsIgnoreCase(explevel1) && actendpoint==expendpoint)
			{
				System.out.println("Test Pass : Image Displayed In Design Area With Level 1");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area With Level 1");
				return false;
			}
		}
		
		
		
		
		
		public boolean 	checkSavingMasterAuthorizationAccWithDateRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{		
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			
			String expMessage="Authorization Flow saved successfully.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			System.out.println("*************************checkSavingMasterAuthorizationAccWithDateRange******************************");
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
				cancelBtn.click();
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
				cancelBtn.click();
				return false;
			}
		}
		
		
		
		
		
		public boolean 	checkMasterAuthorizationasAuthTimeRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();

			nameTxt.sendKeys("AccAuthTimeRange");
			nameTxt.sendKeys(Keys.TAB);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(activeChkBox));
			activeChkBox.click();
			
			
			String actnameTxt=nameTxt.getAttribute("value");
			String expnameTxt="AccAuthTimeRange";
			
			
			boolean actactiveChkBox=activeChkBox.isSelected();
			boolean expactiveChkBox=true;
			
			
			System.out.println("*********************************checkMasterAuthorizationasAuthTimeRange*****************************");
			
			System.out.println("NameTxt   : "+actnameTxt +"  Value Expected : "+expnameTxt);
			System.out.println("ActiveChkBox   : "+actactiveChkBox+"  Value Expected : "+expactiveChkBox);
			
			if(actnameTxt==expnameTxt && actactiveChkBox==expactiveChkBox)
			{
				System.out.println("Test Pass : Successfully inputed AccAuthTimeRange");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT inputed AccAuthTimeRange");
				return false;
			}
		}
		
		
		
		public boolean 	checkStartPointAccAuthTimeRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			getAction().doubleClick(startPointBtn).build().perform();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
	        boolean actconditionTab                 = conditionTab.isDisplayed();
			boolean actuserSelectionTab             = userSelectionTab.isDisplayed();
			boolean actalertsTab                    = alertsTab.isDisplayed();
			boolean actescalationTab                = escalationTab.isDisplayed();
			boolean actdefinitionOkBtn              = definitionOkBtn.isDisplayed();
			boolean actdefinitionCancelBtn          = definitionCancelBtn.isDisplayed();
			
			boolean expconditionTab                 = true;
			boolean expuserSelectionTab             = true;
			boolean expalertsTab                    = true;
			boolean expescalationTab                = true;
			boolean expdefinitionOkBtn              = true;
			boolean expdefinitionCancelBtn          = true;
			
			System.out.println("********************* checkStartPointAccAuthTimeRange ******************");
		
			System.out.println("ConditionTab             :"+actconditionTab           +"    Value Expected : "+expconditionTab);
			System.out.println("UserSelectionTab         :"+actuserSelectionTab       +"    Value Expected : "+expuserSelectionTab);
			System.out.println("AlertsTab                :"+actalertsTab              +"	Value Expected : "+expalertsTab);
			System.out.println("EscalationTab            :"+actescalationTab          +"	Value Expected : "+expescalationTab);
			System.out.println("DefinitionOkBtn          :"+actdefinitionOkBtn        +"	Value Expected : "+expdefinitionOkBtn);
			System.out.println("DefinitionCancelBtn      :"+actdefinitionCancelBtn    +"	Value Expected : "+expdefinitionCancelBtn);
			
			if( actconditionTab==expconditionTab  && actuserSelectionTab==expuserSelectionTab &&
					actalertsTab==expalertsTab && actescalationTab==expescalationTab && actdefinitionCancelBtn==expdefinitionCancelBtn && 
					actdefinitionOkBtn==expdefinitionOkBtn)
		
			{
				System.out.println("Test Pass : Start Point works As Expected");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Start Point NOT works As Expected");
				return false;
			}
		}
		
		
		
		
		public boolean 	checkInputingAccWithTimeRangeinDescription() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
			conditionTab.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			
			descriptionTxt.sendKeys("AccWithTimeRange");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(timeRangeChkBox));
			timeRangeChkBox.click();
			
			Thread.sleep(2000);
			
			System.out.println("*******************************checkInputingAccWithTimeRangeinDescription*************************");
			boolean actconditionNotRequiredChkBox         = conditionNotRequiredChkBox.isSelected();
			boolean acttimeRangeChkBox                    = timeRangeChkBox.isSelected();
			
			boolean expconditionNotRequiredChkBox        = true;
			boolean exptimeRangeChkBox                   = true;
			
			System.out.println("************************************checkInputingAccWithTimeRangeinDescription****************");
			
			System.out.println("conditionNotRequiredChkBox   : "+actconditionNotRequiredChkBox   +"  Value Expected : "+expconditionNotRequiredChkBox);
			System.out.println("timeRangeChkBox              : "+acttimeRangeChkBox              +"  Value Expected : "+exptimeRangeChkBox);
			
			if(actconditionNotRequiredChkBox==expconditionNotRequiredChkBox && acttimeRangeChkBox==exptimeRangeChkBox)
			{
				System.out.println("Test Pass : Successfully created AccWithTimeRange");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT created AccWithTimeRange");
				return false;
			}
			
		}
		
		
		
		public boolean 	checkselectingInActiveintheActionofEntryStatusAccWithTimeRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actionOnEntryStatusDropdown));
			actionOnEntryStatusDropdown.click();
			
			Select s=new Select(actionOnEntryStatusDropdown);
			s.selectByVisibleText("Inactive");
			
			String actactionOnEntryStatusDropdown  = actionOnEntryStatusDropdown.getAttribute("value");
			String expactionOnEntryStatusDropdown  = "1";
			
			System.out.println("*******************************checkselectingInActiveintheActionofEntryStatusAccWithTimeRange******");
			
			System.out.println("ActionOnEntryStatusDropdown  : "+actactionOnEntryStatusDropdown +"  Value Expected : "+expactionOnEntryStatusDropdown);
			
			if(actactionOnEntryStatusDropdown==expactionOnEntryStatusDropdown)
			{
				System.out.println("Test Pass : Selected Active Successfully");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Selected Active Successfully");
				return false;
			}
		}
		
		
		public boolean 	checkselectinguserFromUserSelectiontabAccWithTimeRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userdropdown));
			userdropdown.click();
			
			Select s=new Select(userdropdown);
			s.selectByVisibleText("ProductTestUser");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(level1));
			
			String 	actlevel1=level1.getText();
		    String 	explevel1="AccWithTimeRange";
		    
		    boolean actendpoint=endpoint.isDisplayed();
		    boolean expendpoint = true;
		    
		    
		    System.out.println("*************************checkselectinguserFromUserSelectiontabAccWithTimeRange*******");
			
		    System.out.println("level1     : "+actlevel1     +"  Value Expected : "+explevel1);
		    System.out.println("Endpoint   : "+actendpoint   +"  Value Expected : "+expendpoint);
		    
			if(actlevel1.equalsIgnoreCase(explevel1) && actendpoint==expendpoint)
			
			{
				System.out.println("Test Pass : Image Displayed In Design Area With Level 1");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area With Level 1");
				return false;
			}
		}
		
		
		
		
		
		public boolean 	checkSavingMasterAuthorizationAccWithTimeRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{		
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
			saveBtn.click();
			
			getWaitForAlert();
			getAlert().accept();
			
			
			System.out.println("************************************checkSavingMasterAuthorizationAccWithTimeRange*****************************");
			
			String expMessage="Authorization Flow saved successfully.";
			
			String actMessage=checkValidationMessage(expMessage);
			
			if(actMessage.equalsIgnoreCase(expMessage))
			{
				System.out.println("Error Message is As Expected");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
				cancelBtn.click();
				return true;
			}
			else
			{
				System.out.println(" NO Error Message");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
				cancelBtn.click();
				return false;
			}
		}
		
		
		public boolean checkConditionOptionsintheMasterAuthorization() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();

			nameTxt.sendKeys("AccAuthConditionOnAccName");
			nameTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			getAction().doubleClick(startPointBtn).build().perform();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionTab));
			conditionTab.click();

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			
			descriptionTxt.sendKeys("AccAuthConditionOnAccName");
			
			Thread.sleep(2000);
			
			boolean actconditionNotRequiredChkBox=conditionNotRequiredChkBox.isSelected();
			boolean expconditionNotRequiredChkBox=true;
			
			System.out.println("*********************************checkConditionOptionsintheMasterAuthorization**********************************");
			
			System.out.println("ConditionNotRequiredChkBox     : "+actconditionNotRequiredChkBox+"  Value Expected  : "+expconditionNotRequiredChkBox );
			
			if(actconditionNotRequiredChkBox==expconditionNotRequiredChkBox)
			{
				System.out.println("Test FAil : NOT created AccAuthConditionOnAccName");
				return false;
			}
			else
			{
				System.out.println("Test Pass : Successfully created AccAuthConditionOnAccName");
				return true;
			}
		}
		 
		
		
		public boolean checkCondtionFiledoptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
		    boolean actconjunctionDropdown                 = conjunctionDropdown.isDisplayed();
			boolean actselectFieldTxt                      = selectFieldTxt.isDisplayed();
			boolean actselectOperatorDropdown              = selectOperatorDropdown.isDisplayed();
			boolean actcompareWithDropdown                 = compareWithDropdown.isDisplayed();
			boolean actvalueTxt                            = valueTxt.isDisplayed();
			
			
			boolean expconjunctionDropdown                 = true;
			boolean expselectFieldTxt                      = true;
			boolean expselectOperatorDropdown              = true;
			boolean expcompareWithDropdown                 = true;
			boolean expvalueTxt                            = true;
			
			System.out.println("********************************************** checkCondtionFiledoptions ******************");
		
			System.out.println("ConjunctionDropdown               :"+actconjunctionDropdown    +"  Value Expected : "+expconjunctionDropdown);
			System.out.println("SelectFieldTxt                    :"+actselectFieldTxt         +"  Value Expected : "+expselectFieldTxt);
			System.out.println("SelectOperatorDropdown            :"+actselectOperatorDropdown +"  Value Expected : "+expselectOperatorDropdown);
			System.out.println("CompareWithDropdown               :"+actcompareWithDropdown    +"  Value Expected : "+expcompareWithDropdown);
			
			if( actconjunctionDropdown==expconjunctionDropdown  && actselectFieldTxt==expselectFieldTxt
					&& actselectOperatorDropdown==expselectOperatorDropdown && actcompareWithDropdown==expcompareWithDropdown && actvalueTxt==expvalueTxt)
			
			{
				System.out.println("Test Pass : Conditions Fileds Are Displayed");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Conditions Fileds Are NOT Displayed");
				return false;
			}
		}
		
		
		
		public boolean checkConjunctionFiledoptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			conjunctionDropdown.click();
			
			Select s=new Select(conjunctionDropdown);
			int actconjunctionDropdown=s.getOptions().size();
			int expconjunctionDropdown=12;
			
			
			System.out.println("******************************checkConjunctionFiledoptions***************");
			
			System.out.println("conjunctionDropdown   : "+actconjunctionDropdown +"   Value Expected  : "+expconjunctionDropdown);
			
			if(actconjunctionDropdown==expconjunctionDropdown)
			{
				System.out.println("Test Pass : Conjunction Filed Options Are As Expected");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Conjunction Filed Options Are NOT As Expected");
				return false;
			}
				
		}
		
		
		public boolean checkSelectFiledoptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
			selectFieldTxt.click();
			
			 boolean actsName                           = sName.isDisplayed();
			 boolean actsCode                           = sCode.isDisplayed();
			 boolean actiAccountType                    = iAccountType.isDisplayed();
			 boolean actfCreditLimit                    = fCreditLimit.isDisplayed();
			 boolean actiCreditDays                     = iCreditDays.isDisplayed();
			 boolean actiDebitCreditProposal            = iDebitCreditProposal.isDisplayed();
			 boolean actiDebitCreditRequired            = iDebitCreditRequired.isDisplayed();
			 boolean actiExchangeAdjustmentGainAC       = iExchangeAdjustmentGainAC.isDisplayed();
			 boolean actiExchangeAdjustmentLossAC       = iExchangeAdjustmentLossAC.isDisplayed();
			 boolean actiPrimaryAccount                 = iPrimaryAccount.isDisplayed();
			 boolean actiConsolidationMethod            = iConsolidationMethod.isDisplayed();
			 boolean actsAddress                 	    = sAddress.isDisplayed();
			 boolean actsDeliveryAddress                = sDeliveryAddress.isDisplayed();
			 boolean actsPin                    		= sPin.isDisplayed();
			 boolean actsDeliveryPin                    = sDeliveryPin.isDisplayed();
			 boolean actiCity                 		    = iCity.isDisplayed();
			 boolean actiDeliveryCity                   = iDeliveryCity.isDisplayed();
			 boolean actsEmail                          = sEmail.isDisplayed();
			 boolean actbSendEmailtocustomer            = bSendEmailtocustomer.isDisplayed();
			 boolean actfChequeDiscountLimit            = fChequeDiscountLimit.isDisplayed();
			 boolean actfRateofinterest                 = fRateofinterest.isDisplayed();
		     boolean actiBankAC                  		= iBankAC.isDisplayed();
			 boolean actsTelNo                       	= sTelNo.isDisplayed();
			 boolean actsFaxNo 							= sFaxNo.isDisplayed();
			 boolean actsPassword            			= sPassword.isDisplayed();
			 boolean actbAllowCustomerPortal            = bAllowCustomerPortal.isDisplayed();
			 boolean actsBankAccountName          		= sBankAccountName.isDisplayed();
			 boolean actsBankAccountNumber         		= sBankAccountNumber.isDisplayed();
			 boolean actsIFSCCode               		= sIFSCCode.isDisplayed();
			 boolean actiPaymentType            		= iPaymentType.isDisplayed();
			 boolean actiPDCDiscountedAC    			= iPDCDiscountedAC.isDisplayed();
			 boolean actsFinanceEmail           		= sFinanceEmail.isDisplayed();
			 boolean actsPortalEmail   					= sPortalEmail.isDisplayed();
			 
			 boolean expsName                        	= true;
			 boolean expsCode           				= true;
			 boolean expiAccountType            		= true;
			 boolean expfCreditLimit                  	= true;
			 boolean expiCreditDays               		= true;
			 boolean expiDebitCreditProposal            = true;
			 boolean expiDebitCreditRequired            = true;
			 boolean expiExchangeAdjustmentGainAC       = true;
			 boolean expiExchangeAdjustmentLossAC       = true;
			 boolean expiPrimaryAccount                	= true;
			 boolean expiConsolidationMethod            = true;
			 boolean expsAddress                        = true;
			 boolean expsDeliveryAddress                = true;
			 boolean expsPin                            = true;
			 boolean expsDeliveryPin                    = true;
			 boolean expiCity                           = true;
			 boolean expiDeliveryCity                   = true;
			 boolean expsEmail                   		= true;
			 boolean expbSendEmailtocustomer            = true;
			 boolean expfChequeDiscountLimit            = true;
			 boolean expfRateofinterest                 = true;
			 boolean expiBankAC                 		= true;
			 boolean expsTelNo                          = true;
			 boolean expsFaxNo 							= true;
			 boolean expsPassword           			= true;
			 boolean expbAllowCustomerPortal            = true;
			 boolean expsBankAccountName          		= true;
			 boolean expsBankAccountNumber         		= true;
			 boolean expsIFSCCode               		= true;
			 boolean expiPaymentType            		= true;
			 boolean expiPDCDiscountedAC    			= true;
			 boolean expsFinanceEmail           		= true;
			 boolean expsPortalEmail   					= true;
				
			 System.out.println("****************************** checkSelectFiledoptions  ********************************************");
			 
			 System.out.println("sName                     :"+actsName                           +"  Value Expected : "+expsName);
			 System.out.println("sCode                     :"+actsCode                           +"  Value Expected : "+expsCode);
			 System.out.println("iAccountType              :"+actiAccountType                    +"	 Value Expected : "+expiAccountType);
			 System.out.println("fCreditLimit              :"+actfCreditLimit                    +"	 Value Expected : "+expfCreditLimit);
			 System.out.println("iCreditDays               :"+actiCreditDays                     +"	 Value Expected : "+expiCreditDays);
			 System.out.println("iDebitCreditProposal      :"+actiDebitCreditProposal            +"	 Value Expected : "+expiDebitCreditProposal);
			 System.out.println("iDebitCreditRequired      :"+actiDebitCreditRequired            +"	 Value Expected : "+expiDebitCreditRequired);
			 System.out.println("iExchangeAdjustmentGainAC :"+actiExchangeAdjustmentGainAC       +"	 Value Expected : "+expiExchangeAdjustmentGainAC);
			 System.out.println("iExchangeAdjustmentLossAC :"+actiExchangeAdjustmentLossAC       +"	 Value Expected : "+expiExchangeAdjustmentLossAC);
			 System.out.println("iPrimaryAccount           :"+actiPrimaryAccount                 +"	 Value Expected : "+expiPrimaryAccount);
			 System.out.println("iConsolidationMethod      :"+actiConsolidationMethod            +"	 Value Expected : "+expiConsolidationMethod);
			 System.out.println("sAddress                  :"+actsAddress                        +"  Value Expected : "+expsAddress);
			 System.out.println("sDeliveryAddress          :"+actsDeliveryAddress                +"	 Value Expected : "+expsDeliveryAddress);
			 System.out.println("sPin                      :"+actsPin                            +"	 Value Expected : "+expsPin);
			 System.out.println("sDeliveryPin              :"+actsDeliveryPin                    +"	 Value Expected : "+expsDeliveryPin);
			 System.out.println("iCity                     :"+actiCity                           +"	 Value Expected : "+expiCity);
			 System.out.println("iDeliveryCity             :"+actiDeliveryCity                   +"	 Value Expected : "+expiDeliveryCity);
			 System.out.println("sEmail                    :"+actsEmail                          +"	 Value Expected : "+expsEmail);
			 System.out.println("bSendEmailtocustomer      :"+actbSendEmailtocustomer            +"	 Value Expected : "+expbSendEmailtocustomer);
			 System.out.println("fChequeDiscountLimit      :"+actfChequeDiscountLimit            +"	 Value Expected : "+expfChequeDiscountLimit);
			 System.out.println("fRateofinterest           :"+actfRateofinterest                 +"	 Value Expected : "+expfRateofinterest);
			 System.out.println("iBankAC                   :"+actiBankAC                         +"  Value Expected : "+expiBankAC);
			 System.out.println("sTelNo                    :"+actsTelNo                          +"  Value Expected : "+expsTelNo);
			 System.out.println("keepUnpostedfinAccChkBox  :"+actsFaxNo                          +"	 Value Expected : "+expsFaxNo);
			 System.out.println("sPassword                 :"+actsPassword                       +"	 Value Expected : "+expsPassword);
			 System.out.println("bAllowCustomerPortal      :"+actbAllowCustomerPortal            +"	 Value Expected : "+expbAllowCustomerPortal);
			 System.out.println("sBankAccountName          :"+actsBankAccountName                +"	 Value Expected : "+expsBankAccountName);
			 System.out.println("sBankAccountNumber        :"+actsBankAccountNumber              +"	 Value Expected : "+expsBankAccountNumber);
			 System.out.println("sIFSCCode                 :"+actsIFSCCode                       +"  Value Expected : "+expsIFSCCode);
			 System.out.println("iPaymentType              :"+actiPaymentType                    +"	 Value Expected : "+expiPaymentType);
			 System.out.println("iPDCDiscountedAC          :"+actiPDCDiscountedAC                +"	 Value Expected : "+expiPDCDiscountedAC);
			 System.out.println("sFinanceEmail             :"+actsFinanceEmail                   +"	 Value Expected : "+expsFinanceEmail);
			 System.out.println("sPortalEmail              :"+actsPortalEmail                    +"  Value Expected : "+expsPortalEmail);
				
			 if(actsName==expsName  && actsCode==expsCode 
					 && actiAccountType==expiAccountType && actfCreditLimit==expfCreditLimit 
					 && actiCreditDays==expiCreditDays && actiDebitCreditProposal==expiDebitCreditProposal
				     && actiDebitCreditRequired==expiDebitCreditRequired && actiExchangeAdjustmentGainAC==expiExchangeAdjustmentGainAC 
				     && actiExchangeAdjustmentLossAC==expiExchangeAdjustmentLossAC && actiPrimaryAccount==expiPrimaryAccount && actiConsolidationMethod==expiConsolidationMethod 
					 && actsAddress==expsAddress && actsDeliveryAddress==expsDeliveryAddress && actsPin==expsPin 
					 && actsDeliveryPin==expsDeliveryPin && actiCity==expiCity && actiDeliveryCity==expiDeliveryCity && actsEmail==expsEmail 
					 && actbSendEmailtocustomer==expbSendEmailtocustomer && actfChequeDiscountLimit==expfChequeDiscountLimit && actsPassword==expsPassword
					 && actfRateofinterest==expfRateofinterest && actiBankAC==expiBankAC && actsTelNo==expsTelNo && actsFaxNo==expsFaxNo 
					 && actbAllowCustomerPortal==expbAllowCustomerPortal&& actsBankAccountName==expsBankAccountName &&  actsBankAccountNumber==expsBankAccountNumber
					 && actsIFSCCode==expsIFSCCode && actiPaymentType==expiPaymentType && actiPDCDiscountedAC==expiPDCDiscountedAC && actsFinanceEmail==expsFinanceEmail 
					 &&  actsPortalEmail==expsPortalEmail)
			
					{
				System.out.println("Test Pass : Select filed Options Are As Expected");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Select filed Options Are NOT As Expected");
				return false;
			}
		}
		
		
		
		public boolean checkSelectOperatoroptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			selectOperatorDropdown.click();
			
			Select s=new Select(selectOperatorDropdown);
			int actselectOperatorDropdown=s.getOptions().size();
			int expselectOperatorDropdown=12;
			
			System.out.println("************************checkSelectOperatoroptions*******************");
			
			System.out.println("selectOperatorDropdown   : "+actselectOperatorDropdown +"  value Expected : "+expselectOperatorDropdown);
			
			
			if(actselectOperatorDropdown==expselectOperatorDropdown)
			{
				System.out.println("Test Pass : Select Operator Options Are As Expected");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Select Operator Options Are NOT As Expected");
				return false;
			}
			
		}
		
		
		
		public boolean checkCompareWithoptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			compareWithDropdown.click();
			
			Select s=new Select(compareWithDropdown);
			int actcompareWithDropdown=s.getOptions().size();
			int expcompareWithDropdown=4;
			
			System.out.println("*************************************checkCompareWithoptions*****************************************");
			
			System.out.println("compareWithDropdown   : "+actcompareWithDropdown +"   Value Expeted : "+expcompareWithDropdown);
			
			if(actcompareWithDropdown==expcompareWithDropdown)
			{
				System.out.println("Test Pass : compareWithDropdown Options Are As Expected");
				return true;
			}
			else
			{
				System.out.println("Test Fail : compareWithDropdown Options Are NOT As Expected");
				return false;
			}
		}
		
		
		
		public boolean checkInputingdatainValue() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Select s=new Select(compareWithDropdown);
			s.selectByVisibleText("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueTxt));
			valueTxt.sendKeys("DATA");
			
			Thread.sleep(2000);
			
			String actvalueTxt=valueTxt.getAttribute("value");
			String expvalueTxt="DATA";
			
			System.out.println("*********************************checkInputingdatainValue********************************************");
			
			System.out.println("valueTxt      : "+actvalueTxt+"  Value Expected  : "+expvalueTxt);
			
			if(actvalueTxt.equalsIgnoreCase(expvalueTxt))
			{
				System.out.println("Test Pass : Able to Enter Data in Value");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOTAble to Enter Data in Value");
				return false;
			}
		}
		
		
		
		
		public boolean checkInputingdatainConditionField() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			conjunctionDropdown.click();
			
			Select s=new Select(conjunctionDropdown);
			s.selectByVisibleText("Where");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
			selectFieldTxt.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sName));
			sName.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			selectOperatorDropdown.click();
			
			Select s1=new Select(selectOperatorDropdown);
			s1.selectByVisibleText("Equal to");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			compareWithDropdown.click();
			
			Select s3=new Select(compareWithDropdown);
			s3.selectByVisibleText("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueTxt));
			valueTxt.clear();
			valueTxt.sendKeys("AccAuth");
			
			Thread.sleep(2000);
			
			String actconjunctionDropdown=conjunctionDropdown.getAttribute("value");
			String expconjunctionDropdown="0";
			
			String actselectFieldTxt=selectFieldTxt.getAttribute("value");
			String expselectFieldTxt="sName";

			String actselectOperatorDropdown=selectOperatorDropdown.getAttribute("value");
			String expselectOperatorDropdown="0";
			
			String actcompareWithDropdown=compareWithDropdown.getAttribute("value");
			String expcompareWithDropdown="0";
			
			String actvalueTxt=valueTxt.getAttribute("value");
			String expvalueTxt="AccAuth";
			
			
			System.out.println("*********************************checkInputingdatainConditionField*********************************************");
			
			System.out.println("conjunctionDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("selectFieldTxt         : "+actselectFieldTxt             +"  Value Expected :"+expselectFieldTxt);
			System.out.println("selectOperatorDropdown : "+actselectOperatorDropdown     +"  Value Expected :"+expselectOperatorDropdown);
			System.out.println("compareWithDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("valueTxt               : "+actvalueTxt                   +"  Value Expected :"+expvalueTxt);
			
			if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && actselectOperatorDropdown.equals(expselectOperatorDropdown)
					&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt))
			{
				System.out.println("Test Pass : Data entered");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT entered");
				return false;
			}
			
		}
		
		
		
		public boolean 	checkselectingInActiveintheActionofEntryStatusAccAuthConditionOnAccName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actionOnEntryStatusDropdown));
			actionOnEntryStatusDropdown.click();
			
			Select s=new Select(actionOnEntryStatusDropdown);
			s.selectByVisibleText("Inactive");
			
			String actactionOnEntryStatusDropdown=actionOnEntryStatusDropdown.getAttribute("value");
			String expactionOnEntryStatusDropdown="1";
			
			System.out.println("***********************************checkselectingInActiveintheActionofEntryStatusAccAuthConditionOnAccName**********");
			
			System.out.println("ActionOnEntryStatusDropdown   : "+actactionOnEntryStatusDropdown +"  Value Expected : "+expactionOnEntryStatusDropdown);
			
			if(actactionOnEntryStatusDropdown.equals(expactionOnEntryStatusDropdown))
			{
				System.out.println("Test Pass : Selected Active Successfully");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Selected Active Successfully");
				return false;
			}
		}
		
		
		public boolean 	checkselectinguserFromUserSelectiontabAndSavingAccAuthConditionOnAccName() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userdropdown));
			userdropdown.click();
			
			Select s=new Select(userdropdown);
			s.selectByVisibleText("ProductTestUser");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(level1));
			
			String actdata,expdata;
			String actlevel1=level1.getText();
			String explevel1="AccAuthConditionOnAccName";
			
			System.out.println("*************************************checkselectinguserFromUserSelectiontabAndSavingAccAuthConditionOnAccName************");
			
			System.out.println("level1           : "+actlevel1 +"  Value Expected : "+explevel1);
			
			if(actlevel1.equalsIgnoreCase(explevel1))
			{
				System.out.println("Test Pass : Image Displayed In Design Area With Level 1");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
							
				String expMessage="Authorization Flow saved successfully.";
				
				String actMessage=checkValidationMessage(expMessage);
				
				if(actMessage.equalsIgnoreCase(expMessage))
				{
					System.out.println("Error Message is As Expected");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return true;
				}
				else
				{
					System.out.println(" NO Error Message");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return false;
				}
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area With Level 1");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
				cancelBtn.click();
				return false;
			}
		}
		
		
		
		
		
		public boolean checkInputingNameAsAccWithConAsAccType() throws InterruptedException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys("AccWithConAsAccType");
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			getAction().doubleClick(startPointBtn).build().perform();
			
			
			boolean actdescriptionTxt=descriptionTxt.isDisplayed();
			boolean expdescriptionTxt=true;
			
			System.out.println("**********************checkInputingNameAsAccWithConAsAccType********************");
			
			System.out.println("descriptionTxt      : "+actdescriptionTxt +"  Value Expected : "+expdescriptionTxt);
			
			if(actdescriptionTxt==expdescriptionTxt)
			{
				System.out.println("Test Pass : Inputed AccWithConAsAccType");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT Inputed AccWithConAsAccType");
				return false;
			}
			
		}
		
		
		public boolean checkInputingdatainConditionFieldAccWithConAsAccType() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.sendKeys("AccWithConAsAccType");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			conjunctionDropdown.click();
			
			Select s=new Select(conjunctionDropdown);
			s.selectByVisibleText("Where");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
			selectFieldTxt.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(iAccountType));
			iAccountType.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			selectOperatorDropdown.click();
			
			Select s1=new Select(selectOperatorDropdown);
			s1.selectByVisibleText("Equal to");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			compareWithDropdown.click();
			
			Select s3=new Select(compareWithDropdown);
			s3.selectByVisibleText("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueDropdown));
			valueDropdown.click();
			Select s2=new Select(valueDropdown);
			s2.selectByVisibleText("Customer/Vendor");
			
			Thread.sleep(2000);
			
		
			
			String actconjunctionDropdown=conjunctionDropdown.getAttribute("value");
			String expconjunctionDropdown="0";
			
			String actselectFieldTxt=selectFieldTxt.getAttribute("value");
			String expselectFieldTxt="iAccountType";

			String actselectOperatorDropdown=selectOperatorDropdown.getAttribute("value");
			String expselectOperatorDropdown="0";
			
			String actcompareWithDropdown=compareWithDropdown.getAttribute("value");
			String expcompareWithDropdown="0";
			
			String actvalueDropdown=valueDropdown.getAttribute("value");
			String expvalueDropdown="7";
			
			
			System.out.println("*********************************checkInputingdatainConditionFieldAccWithConAsAccType*********************************************");
			
			System.out.println("conjunctionDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("selectFieldTxt         : "+actselectFieldTxt             +"  Value Expected :"+expselectFieldTxt);
			System.out.println("selectOperatorDropdown : "+actselectOperatorDropdown     +"  Value Expected :"+expselectOperatorDropdown);
			System.out.println("compareWithDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("valueDropdown          : "+actvalueDropdown              +"  Value Expected :"+expvalueDropdown);
			
			if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && actselectOperatorDropdown.equals(expselectOperatorDropdown)
					&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueDropdown.equalsIgnoreCase(expvalueDropdown))
			
			{
				System.out.println("Test Pass : Data entered");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT entered");
				return false;
			}
			
		}
		
		
		
		public boolean 	checkselectingInActiveintheActionofEntryStatusAccWithConAsAccType() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actionOnEntryStatusDropdown));
			actionOnEntryStatusDropdown.click();
			
			Select s=new Select(actionOnEntryStatusDropdown);
			s.selectByVisibleText("Inactive");
			
			String actactionOnEntryStatusDropdown=actionOnEntryStatusDropdown.getAttribute("value");
			String expactionOnEntryStatusDropdown="1";
			
			System.out.println("****************checkselectingInActiveintheActionofEntryStatusAccWithConAsAccType**********************************");
			
			System.out.println("actionOnEntryStatusDropdown  : "+actactionOnEntryStatusDropdown+"  Value Expected  :"+expactionOnEntryStatusDropdown);
			
			if(actactionOnEntryStatusDropdown.equals(expactionOnEntryStatusDropdown))
			{
				System.out.println("Test Pass : Selected Active Successfully");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Selected Active Successfully");
				return false;
			}
		}
		
		
		public boolean 	checkselectinguserFromUserSelectiontabAndSavingAccWithConAsAccType() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userdropdown));
			userdropdown.click();
			
			Select s=new Select(userdropdown);
			s.selectByVisibleText("ProductTestUser");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(level1));
			
			String 	actlevel1=level1.getText();
		    String 	explevel1="AccWithConAsAccType";
		    
		    boolean actendpoint=endpoint.isDisplayed();
		    boolean expendpoint = true;
		    
		    
		    System.out.println("*************************checkselectinguserFromUserSelectiontabAndSavingAccWithConAsAccType*******");
			
		    System.out.println("level1   : "+actlevel1     +"  Value Expected : "+explevel1);
		    System.out.println("Endpoint   : "+actendpoint +"  Value Expected : "+expendpoint);
		    
			if(actlevel1.equalsIgnoreCase(explevel1) && actendpoint==expendpoint)
			
			{
				System.out.println("Test Pass : Image Displayed In Design Area With Level 1");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage="Authorization Flow saved successfully.";
				
				String actMessage=checkValidationMessage(expMessage);
				
				if(actMessage.equalsIgnoreCase(expMessage))
				{
					System.out.println("Error Message is As Expected");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return true;
				}
				else
				{
					System.out.println(" NO Error Message");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return false;
				}
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area With Level 1");
				return false;
			}
		}
		
		
		
		public boolean checkInputingNameAsAccWithConAsCreditLimit() throws InterruptedException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys("AccWithConAsCreditLimit");
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			getAction().doubleClick(startPointBtn).build().perform();
			
			boolean actdescriptionTxt=descriptionTxt.isDisplayed();
			
			boolean expdescriptionTxt=true;
			
			System.out.println("*************************checkInputingNameAsAccWithConAsCreditLimit********"); 
			
			
			System.out.println("DescriptionTxt   : "+actdescriptionTxt +"  Value Expected : "+expdescriptionTxt);
			
			if(actdescriptionTxt==expdescriptionTxt)
			{
				System.out.println("Test Pass : Inputed AccWithConAsCreditLimit");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT Inputed AccWithConAsCreditLimit");
				return false;
			}
			
		}
		
		
		public boolean checkInputingdatainConditionFieldAccWithConAsCreditLimit() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.sendKeys("AccWithConAsCreditLimit");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			conjunctionDropdown.click();
			
			Select s=new Select(conjunctionDropdown);
			s.selectByVisibleText("Where");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
			selectFieldTxt.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fCreditLimit));
			fCreditLimit.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			selectOperatorDropdown.click();
			
			Select s1=new Select(selectOperatorDropdown);
			s1.selectByVisibleText("Equal to");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			compareWithDropdown.click();
			
			Select s3=new Select(compareWithDropdown);
			s3.selectByVisibleText("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueTxt));
			valueTxt.click();
			valueTxt.sendKeys("90000");
			
			
			String actconjunctionDropdown=conjunctionDropdown.getAttribute("value");
			String expconjunctionDropdown="0";
			
			String actselectFieldTxt=selectFieldTxt.getAttribute("value");
			String expselectFieldTxt="fCreditLimit";

			String actselectOperatorDropdown=selectOperatorDropdown.getAttribute("value");
			String expselectOperatorDropdown="0";
			
			String actcompareWithDropdown=compareWithDropdown.getAttribute("value");
			String expcompareWithDropdown="0";
			
			String actvalueTxt=valueTxt.getAttribute("value");
			String expvalueTxt="90000";
			
			
			System.out.println("*********************************checkInputingdatainConditionFieldAccWithConAsCreditLimit*********************************************");
			
			System.out.println("conjunctionDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("selectFieldTxt         : "+actselectFieldTxt             +"  Value Expected :"+expselectFieldTxt);
			System.out.println("selectOperatorDropdown : "+actselectOperatorDropdown     +"  Value Expected :"+expselectOperatorDropdown);
			System.out.println("compareWithDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("valueTxt               : "+actvalueTxt                   +"  Value Expected :"+expvalueTxt);
			
			if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && actselectOperatorDropdown.equals(expselectOperatorDropdown)
					&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt))
			{
				System.out.println("Test Pass : Data entered");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT entered");
				return false;
			}
			
		}
		
		
		
		public boolean 	checkselectingInActiveintheActionofEntryStatusAccWithConAsCreditLimit() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actionOnEntryStatusDropdown));
			actionOnEntryStatusDropdown.click();
			
			Select s=new Select(actionOnEntryStatusDropdown);
			s.selectByVisibleText("Inactive");
			
			String actactionOnEntryStatusDropdown=actionOnEntryStatusDropdown.getAttribute("value");
			String expactionOnEntryStatusDropdown="1";
			
			System.out.println("************************checkselectingInActiveintheActionofEntryStatusAccWithConAsCreditLimit*************************");
			
			System.out.println("actionOnEntryStatusDropdown  : "+actactionOnEntryStatusDropdown +"   Value Expected  :"+expactionOnEntryStatusDropdown);
			
			if(actactionOnEntryStatusDropdown.equals(expactionOnEntryStatusDropdown))
			{
				System.out.println("Test Pass : Selected Active Successfully");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Selected Active Successfully");
				return false;
			}
		}
		
		
		public boolean 	checkselectinguserFromUserSelectiontabAndSavingAccWithConAsCreditLimit() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userdropdown));
			userdropdown.click();
			
			Select s=new Select(userdropdown);
			s.selectByVisibleText("ProductTestUser");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(level1));
			
			String 	actlevel1=level1.getText();
		    String 	explevel1="AccWithConAsCreditLimit";
		    
		    boolean actendpoint=endpoint.isDisplayed();
		    boolean expendpoint = true;
		    
		    
		    System.out.println("*************************checkselectinguserFromUserSelectiontabAndSavingAccWithConAsCreditLimit*******");
			
		    System.out.println("level1   : "+actlevel1     +"  Value Expected : "+explevel1);
		    System.out.println("Endpoint   : "+actendpoint +"  Value Expected : "+expendpoint);
		    
			if(actlevel1.equalsIgnoreCase(explevel1) && actendpoint==expendpoint)
			{
				System.out.println("Test Pass : Image Displayed In Design Area With Level 1");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage="Authorization Flow saved successfully.";
				
				String actMessage=checkValidationMessage(expMessage);
				
				if(actMessage.equalsIgnoreCase(expMessage))
				{
					System.out.println("Error Message is As Expected");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return true;
				}
				else
				{
					System.out.println(" NO Error Message");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return false;
				}
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area With Level 1");
				return false;
			}
		}
		
		
		
		
		public boolean checkInputingNameAsAccWithConAsCreditDays() throws InterruptedException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys("AccWithConAsCreditDays");
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			getAction().doubleClick(startPointBtn).build().perform();
			

			boolean actdescriptionTxt=descriptionTxt.isDisplayed();
			
			boolean expdescriptionTxt=true;
			
			System.out.println("*************************checkInputingNameAsAccWithConAsCreditDays********"); 
			
			
			System.out.println("DescriptionTxt   : "+actdescriptionTxt +"  Value Expected : "+expdescriptionTxt);
			
			if(actdescriptionTxt==expdescriptionTxt)
			
			{
				System.out.println("Test Pass : Inputed AccWithConAsCreditLimit");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT Inputed AccWithConAsCreditLimit");
				return false;
			}
			
		}
		
		
		public boolean checkInputingdatainConditionFieldAccWithConAsCreditDays() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.sendKeys("AccWithConAsCreditDays");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			conjunctionDropdown.click();
			
			Select s=new Select(conjunctionDropdown);
			s.selectByVisibleText("Where");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
			selectFieldTxt.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(iCreditDays));
			iCreditDays.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			selectOperatorDropdown.click();
			
			Select s1=new Select(selectOperatorDropdown);
			s1.selectByVisibleText("Equal to");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			compareWithDropdown.click();
			
			Select s3=new Select(compareWithDropdown);
			s3.selectByVisibleText("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueTxt));
			valueTxt.click();
			valueTxt.sendKeys("90000");
			
			String actconjunctionDropdown=conjunctionDropdown.getAttribute("value");
			String expconjunctionDropdown="0";
			
			String actselectFieldTxt=selectFieldTxt.getAttribute("value");
			String expselectFieldTxt="iCreditDays";

			String actselectOperatorDropdown=selectOperatorDropdown.getAttribute("value");
			String expselectOperatorDropdown="0";
			
			String actcompareWithDropdown=compareWithDropdown.getAttribute("value");
			String expcompareWithDropdown="0";
			
			String actvalueTxt=valueTxt.getAttribute("value");
			String expvalueTxt="90000";
			
			
			System.out.println("*********************************checkInputingdatainConditionFieldAccWithConAsCreditDays*********************************************");
			
			System.out.println("conjunctionDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("selectFieldTxt         : "+actselectFieldTxt             +"  Value Expected :"+expselectFieldTxt);
			System.out.println("selectOperatorDropdown : "+actselectOperatorDropdown     +"  Value Expected :"+expselectOperatorDropdown);
			System.out.println("compareWithDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("valueTxt               : "+actvalueTxt                   +"  Value Expected :"+expvalueTxt);
			
			if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && actselectOperatorDropdown.equals(expselectOperatorDropdown)
					&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt))
			
			{
				System.out.println("Test Pass : Data entered");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT entered");
				return false;
			}
			
		}
		
		
		
		public boolean 	checkselectingInActiveintheActionofEntryStatusAccWithConAsCreditDays() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actionOnEntryStatusDropdown));
			actionOnEntryStatusDropdown.click();
			
			Select s=new Select(actionOnEntryStatusDropdown);
			s.selectByVisibleText("Inactive");
			
			String actactionOnEntryStatusDropdown=actionOnEntryStatusDropdown.getAttribute("value");
			String expactionOnEntryStatusDropdown="1";
			
			System.out.println("*****************************checkselectingInActiveintheActionofEntryStatusAccWithConAsCreditDays**************");
			
			System.out.println("actionOnEntryStatusDropdown   : "+actactionOnEntryStatusDropdown+"  Value Expected  : "+expactionOnEntryStatusDropdown);
			
			if(actactionOnEntryStatusDropdown.equals(expactionOnEntryStatusDropdown))
			{
				System.out.println("Test Pass : Selected Active Successfully");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Selected Active Successfully");
				return false;
			}
		}
		
		
		public boolean 	checkselectinguserFromUserSelectiontabAndSavingAccWithConAsCreditDays() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userdropdown));
			userdropdown.click();
			
			Select s=new Select(userdropdown);
			s.selectByVisibleText("ProductTestUser");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(level1));
			
			String 	actlevel1=level1.getText();
		    String 	explevel1="AccWithConAsCreditDays";
		    
		    boolean actendpoint=endpoint.isDisplayed();
		    boolean expendpoint = true;
		    
		    
		    System.out.println("*************************checkselectinguserFromUserSelectiontabAndSavingAccWithConAsCreditDays*******");
			
		    System.out.println("level1     : "+actlevel1     +"  Value Expected : "+explevel1);
		    System.out.println("Endpoint   : "+actendpoint   +"  Value Expected : "+expendpoint);
		    
			if(actlevel1.equalsIgnoreCase(explevel1) && actendpoint==expendpoint)
			
			{
				System.out.println("Test Pass : Image Displayed In Design Area With Level 1");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage="Authorization Flow saved successfully.";
				
				String actMessage=checkValidationMessage(expMessage);
				
				if(actMessage.equalsIgnoreCase(expMessage))
				{
					System.out.println("Error Message is As Expected");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return true;
				}
				else
				{
					System.out.println(" NO Error Message");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return false;
				}
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area With Level 1");
				return false;
			}
		}
		
		
		
		
		
		public boolean checkInputingNameAsAccAuthConWithCreditLimitRange() throws InterruptedException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys("AccAuthConWithCreditLimitRange");
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			getAction().doubleClick(startPointBtn).build().perform();
			
	        boolean actdescriptionTxt=descriptionTxt.isDisplayed();
			
			boolean expdescriptionTxt=true;
			
			System.out.println("**********************checkInputingNameAsAccAuthConWithCreditLimitRange********"); 
			
			
			System.out.println("DescriptionTxt   : "+actdescriptionTxt +"  Value Expected : "+expdescriptionTxt);
			
			if(actdescriptionTxt==expdescriptionTxt)
			{
				System.out.println("Test Pass : Inputed AccWithConAsCreditLimit");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT Inputed AccWithConAsCreditLimit");
				return false;
			}
			
		}
		
		
		public boolean checkInputingdatainConditionFieldAccAuthConWithCreditLimitRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.sendKeys("AccAuthConWithCreditLimitRange");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown));
			conjunctionDropdown.click();
			
			Select s=new Select(conjunctionDropdown);
			s.selectByVisibleText("Where");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt));
			selectFieldTxt.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fCreditLimit));
			fCreditLimit.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown));
			selectOperatorDropdown.click();
			
			Select s1=new Select(selectOperatorDropdown);
			s1.selectByVisibleText("Greater than or equal to");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(compareWithDropdown));
			compareWithDropdown.click();
			
			Select s3=new Select(compareWithDropdown);
			s3.selectByVisibleText("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueTxt));
			valueTxt.click();
			valueTxt.sendKeys("100");
			
			
			String actconjunctionDropdown=conjunctionDropdown.getAttribute("value");
			String expconjunctionDropdown="0";
			
			String actselectFieldTxt=selectFieldTxt.getAttribute("value");
			String expselectFieldTxt="fCreditLimit";

			String actselectOperatorDropdown=selectOperatorDropdown.getAttribute("value");
			String expselectOperatorDropdown="5";
			
			String actcompareWithDropdown=compareWithDropdown.getAttribute("value");
			String expcompareWithDropdown="0";
			
			String actvalueTxt=valueTxt.getAttribute("value");
			String expvalueTxt="100";
			
			
			System.out.println("*********************************checkInputingdatainConditionFieldAccAuthConWithCreditLimitRange*********************************************");
			
			System.out.println("conjunctionDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("selectFieldTxt         : "+actselectFieldTxt             +"  Value Expected :"+expselectFieldTxt);
			System.out.println("selectOperatorDropdown : "+actselectOperatorDropdown     +"  Value Expected :"+expselectOperatorDropdown);
			System.out.println("compareWithDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("valueTxt               : "+actvalueTxt                   +"  Value Expected :"+expvalueTxt);
			
			if(actconjunctionDropdown.equals(expconjunctionDropdown) && actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt) && actselectOperatorDropdown.equals(expselectOperatorDropdown)
					&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt))
			{
				System.out.println("Test Pass : Data entered");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(addRowBtn));
				addRowBtn.click();
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT entered");
				return false;
			}
			
		}
		

		public boolean checkInputingdatainConditionFieldTWOAccAuthConWithCreditLimitRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conjunctionDropdown2));
			conjunctionDropdown2.click();
			
			Select s=new Select(conjunctionDropdown2);
			s.selectByVisibleText("And");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectFieldTxt2));
			selectFieldTxt2.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fCreditLimit));
			fCreditLimit.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(selectOperatorDropdown2));
			selectOperatorDropdown2.click();
			
			Select s1=new Select(selectOperatorDropdown2);
			s1.selectByVisibleText("Less than or equal to");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(comparewithDropdown2));
			comparewithDropdown2.click();
			
			Select s3=new Select(comparewithDropdown2);
			s3.selectByVisibleText("value");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(valueTxt2));
			valueTxt2.click();
			valueTxt2.sendKeys("1000");
			
			
			Thread.sleep(2000);
			
			String actconjunctionDropdown=conjunctionDropdown2.getAttribute("value");
			String expconjunctionDropdown="2";
			
		/*	String actselectFieldTxt=selectFieldTxt.getAttribute("value");
			String expselectFieldTxt="fCreditLimit";*/

			String actselectOperatorDropdown=selectOperatorDropdown2.getAttribute("value");
			String expselectOperatorDropdown="4";
			
			String actcompareWithDropdown=comparewithDropdown2.getAttribute("value");
			String expcompareWithDropdown="0";
			
			String actvalueTxt=valueTxt2.getAttribute("value");
			String expvalueTxt="1000";
			
			
			System.out.println("******************************checkInputingdatainConditionFieldTWOAccAuthConWithCreditLimitRange*********************************************");
			
			System.out.println("conjunctionDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			/*System.out.println("selectFieldTxt         : "+actselectFieldTxt             +"  Value Expected :"+expselectFieldTxt);*/
			System.out.println("selectOperatorDropdown : "+actselectOperatorDropdown     +"  Value Expected :"+expselectOperatorDropdown);
			System.out.println("compareWithDropdown    : "+actconjunctionDropdown        +"  Value Expected :"+expconjunctionDropdown);
			System.out.println("valueTxt               : "+actvalueTxt                   +"  Value Expected :"+expvalueTxt);
			
			if(actconjunctionDropdown.equals(expconjunctionDropdown) /*&& actselectFieldTxt.equalsIgnoreCase(expselectFieldTxt)*/ && actselectOperatorDropdown.equals(expselectOperatorDropdown)
					&& actcompareWithDropdown.equals(expcompareWithDropdown) && actvalueTxt.equalsIgnoreCase(expvalueTxt))
			
			{
				System.out.println("Test Pass : Data entered");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Data NOT entered");
				return false;
			}
			
		}


		
		public boolean 	checkselectingInActiveintheActionofEntryStatusAccAuthConWithCreditLimitRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actionOnEntryStatusDropdown));
			actionOnEntryStatusDropdown.click();
			
			Select s=new Select(actionOnEntryStatusDropdown);
			s.selectByVisibleText("Inactive");
			
			String actactionOnEntryStatusDropdown=actionOnEntryStatusDropdown.getAttribute("value");
			String expactionOnEntryStatusDropdown="1";
			
			System.out.println("****************************************checkselectingInActiveintheActionofEntryStatusAccAuthConWithCreditLimitRange*********************");
			
			System.out.println("actionOnEntryStatusDropdown   : "+actactionOnEntryStatusDropdown +"  Value Expected  : "+expactionOnEntryStatusDropdown);
			
			
			if(actactionOnEntryStatusDropdown.equals(expactionOnEntryStatusDropdown))
			{
				System.out.println("Test Pass : Selected Active Successfully");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Selected Active Successfully");
				return false;
			}
		}
		
		
		public boolean 	checkselectinguserFromUserSelectiontabAndSavingAccAuthConWithCreditLimitRange() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
			userSelectionTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userdropdown));
			userdropdown.click();
			
			Select s=new Select(userdropdown);
			s.selectByVisibleText("ProductTestUser");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(level1));
			
			String 	actlevel1=level1.getText();
		    String 	explevel1="AccAuthConWithCreditLimitRange";
		    
		    boolean actendpoint=endpoint.isDisplayed();
		    boolean expendpoint = true;
		    
		    
		    System.out.println("*************************checkselectinguserFromUserSelectiontabAndSavingAccWithConAsCreditDays*******");
			
		    System.out.println("level1   : "+actlevel1     +"  Value Expected : "+explevel1);
		    System.out.println("Endpoint   : "+actendpoint +"  Value Expected : "+expendpoint);
		    
			if(actlevel1.equalsIgnoreCase(explevel1) && actendpoint==expendpoint)
			{
				System.out.println("Test Pass : Image Displayed In Design Area With Level 1");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage="Authorization Flow saved successfully.";
				
				String actMessage=checkValidationMessage(expMessage);
				
				if(actMessage.equalsIgnoreCase(expMessage))
				{
					System.out.println("Error Message is As Expected");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return true;
				}
				else
				{
					System.out.println(" NO Error Message");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return false;
				}
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area With Level 1");
				return false;
			}
		}
		
		
		
		

		public boolean checkInputingNameAsAccFullAuth() throws InterruptedException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys("AccFullAuth");
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			getAction().doubleClick(startPointBtn).build().perform();
			
	        boolean actdescriptionTxt=descriptionTxt.isDisplayed();
			boolean expdescriptionTxt=true;
			
			System.out.println("**********************checkInputingNameAsAccFullAuth********"); 
			
			
			System.out.println("DescriptionTxt   : "+actdescriptionTxt +"  Value Expected : "+expdescriptionTxt);
			
			if(actdescriptionTxt==expdescriptionTxt)
			{
				System.out.println("Test Pass : Inputed AccFullAuth");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT Inputed AccFullAuth");
				return false;
			}
		}

		

		public boolean checkInputingDescriptionAsAccFullAuthAndselectConditionNotRequired() throws InterruptedException
		{

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			
			descriptionTxt.sendKeys("AccFullAuth");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();
			
			
			boolean actconditionNotRequiredChkBox=conditionNotRequiredChkBox.isSelected();
			boolean expconditionNotRequiredChkBox=true;

			System.out.println("*******************************checkInputingDescriptionAsAccFullAuthAndselectConditionNotRequired******");
			
			System.out.println("conditionNotRequiredChkBox     : "+actconditionNotRequiredChkBox +"   Value Expected  : "+expconditionNotRequiredChkBox);
			
			if(actconditionNotRequiredChkBox==expconditionNotRequiredChkBox)
			{
				System.out.println("Test pass : Entered Description and checked contion not required");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT Entered Description and checked contion not required");
				return false;
			}
		}

		
		
		public boolean checkSelectingFullAuthorization() throws InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fullAuthorizationChkBox));
			fullAuthorizationChkBox.click();
			
			Thread.sleep(2000);
			
			boolean actfullAuthorizationChkBox=fullAuthorizationChkBox.isSelected();
			boolean expfullAuthorizationChkBox=true;

			System.out.println("*******************************checkSelectingFullAuthorization******");
			
			System.out.println("fullAuthorizationChkBox     : "+actfullAuthorizationChkBox +"   Value Expected  : "+expfullAuthorizationChkBox);
			
			if(actfullAuthorizationChkBox==expfullAuthorizationChkBox)
			
			{
				System.out.println("Test pass : checked fullAuthorizationChkBox");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
				userSelectionTab.click();
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT checked fullAuthorizationChkBox");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
				userSelectionTab.click();
				return false;
			}
		}
		
		
		
		
		public boolean 	checkselectinguserFromUserSelectiontabAndSavingAccFullAuth() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userdropdown));
			userdropdown.click();
			
			Select s=new Select(userdropdown);
			s.selectByVisibleText("ProductTestUser");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(level1));
			
			String 	actlevel1=level1.getText();
		    String 	explevel1="AccFullAuth";
		    
		    boolean actendpoint=endpoint.isDisplayed();
		    boolean expendpoint = true;
		    
		    
		    System.out.println("*************************checkselectinguserFromUserSelectiontabAndSavingAccFullAuth*******");
			
		    System.out.println("level1   : "+actlevel1     +"  Value Expected : "+explevel1);
		    System.out.println("Endpoint   : "+actendpoint +"  Value Expected : "+expendpoint);
		    
			if(actlevel1.equalsIgnoreCase(explevel1) && actendpoint==expendpoint)
			
			{
				System.out.println("Test Pass : Image Displayed In Design Area With Level 1");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage="Authorization Flow saved successfully.";
				
				String actMessage=checkValidationMessage(expMessage);
				
				if(actMessage.equalsIgnoreCase(expMessage))
				{
					System.out.println("Error Message is As Expected");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return true;
				}
				else
				{
					System.out.println(" NO Error Message");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return false;
				}
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area With Level 1");
				return false;
			}
		}
		
		
		
		
		

		public boolean checkInputingNameAsAccUndoAuthOnEditing() throws InterruptedException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys("AccUndoAuthOnEditing");
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(startPointBtn));
			getAction().doubleClick(startPointBtn).build().perform();
			
			boolean actdescriptionTxt=descriptionTxt.isDisplayed();
			boolean expdescriptionTxt=true;
				
			System.out.println("**********************checkInputingNameAsAccUndoAuthOnEditing********"); 
				
				
			System.out.println("DescriptionTxt   : "+actdescriptionTxt +"  Value Expected : "+expdescriptionTxt);
				
			if(actdescriptionTxt==expdescriptionTxt)
			
			{
				System.out.println("Test Pass : Inputed AccUndoAuthOnEditing ");
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT Inputed AccUndoAuthOnEditing ");
				return false;
			}
		}

		

		public boolean checkInputingDescriptionAsAccUndoAuthOnEditingAndselectConditionNotRequired() throws InterruptedException
		{

			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(descriptionTxt));
			descriptionTxt.click();
			
			descriptionTxt.sendKeys("AccUndoAuthOnEditing");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(conditionNotRequiredChkBox));
			conditionNotRequiredChkBox.click();
			
			boolean actconditionNotRequiredChkBox=conditionNotRequiredChkBox.isSelected();
			boolean expconditionNotRequiredChkBox=true;

			System.out.println("*******************************checkInputingDescriptionAsAccUndoAuthOnEditingAndselectConditionNotRequired******");
			
			System.out.println("conditionNotRequiredChkBox     : "+actconditionNotRequiredChkBox +"   Value Expected  : "+expconditionNotRequiredChkBox);
			
			if(actconditionNotRequiredChkBox==expconditionNotRequiredChkBox)
		
			{
				System.out.println("Test pass : Entered Description and checked contion not required"  );
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT Entered Description and checked contion not required");
				return false;
			}
		}

		
		
		public boolean checkSelectingFullinUndoAuthorizationAccUndoAuthOnEditing() throws InterruptedException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(undoAuthEditingDropdown));
			undoAuthEditingDropdown.click();
			
			Thread.sleep(2000);
			
			Select s=new Select(undoAuthEditingDropdown);
			s.selectByVisibleText("Full");
			
			Thread.sleep(2000);
			
			String actval,expval;
			String actundoAuthEditingDropdown=undoAuthEditingDropdown.getAttribute("value");
			String expundoAuthEditingDropdown="1";
			
			
			System.out.println("****************************checkSelectingFullinUndoAuthorizationAccUndoAuthOnEditing***********************");
			
			System.out.println("UndoAuthEditingDropdown      : "+actundoAuthEditingDropdown +"  Value Expected  : "+expundoAuthEditingDropdown);
			
			if(actundoAuthEditingDropdown.endsWith(expundoAuthEditingDropdown))
			{
				System.out.println("Test pass : Selected Full in Undo Auth On Editing");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
				userSelectionTab.click();
				return true;
			}
			else
			{
				System.out.println("Test Fail : NOT Selected Full in Undo Auth On Editing");
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userSelectionTab));
				userSelectionTab.click();
				return false;
			}
		}
		
		
		
		
		public boolean 	checkselectinguserFromUserSelectiontabAndSavingAccUndoAuthOnEditing() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userrow1));
			userrow1.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userdropdown));
			userdropdown.click();
			
			Select s=new Select(userdropdown);
			s.selectByVisibleText("ProductTestUser");
			
			Thread.sleep(5000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(definitionOkBtn));
			definitionOkBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(level1));
			
			String 	actlevel1=level1.getText();
		    String 	explevel1="AccUndoAuthOnEditing";
		    
		    boolean actendpoint=endpoint.isDisplayed();
		    boolean expendpoint = true;
		    
		    
		    System.out.println("*************************checkselectinguserFromUserSelectiontabAndSavingAccFullAuth*******");
			
		    System.out.println("level1   : "+actlevel1     +"  Value Expected : "+explevel1);
		    System.out.println("Endpoint   : "+actendpoint +"  Value Expected : "+expendpoint);
		    
			if(actlevel1.equalsIgnoreCase(explevel1) && actendpoint==expendpoint)
			{
				System.out.println("Test Pass : Image Displayed In Design Area With Level 1");
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
				saveBtn.click();
				
				String expMessage="Authorization Flow saved successfully.";
				
				String actMessage=checkValidationMessage(expMessage);
				
				if(actMessage.equalsIgnoreCase(expMessage))
				{
					System.out.println("Error Message is As Expected");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return true;
				}
				else
				{
					System.out.println(" NO Error Message");
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(cancelBtn));
					cancelBtn.click();
					return false;
				}
			}
			else
			{
				System.out.println("Test Fail : Image NOT Displayed In Design Area With Level 1");
				return false;
			}
		}
		
		@FindBy(xpath="//div[@id='divLevel1']//i[@class='icon-edit icon-font7']")
		private static WebElement Level1editBtn;
		
		
		
		public boolean 	checkEmailTemplateOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(settingsmenuBtn));
			settingsmenuBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(masterAuthorizationBtn));
			masterAuthorizationBtn.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDropdown));
			budgetDropdown.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(account));
			account.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(nameTxt));
			nameTxt.click();
			nameTxt.clear();
			nameTxt.sendKeys("AccUndoAuthOnEditing");
			nameTxt.sendKeys(Keys.TAB);
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(Level1editBtn));
			Level1editBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(alertsTab));
			alertsTab.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(emailTemplateBtn));
			emailTemplateBtn.click();
			
			Thread.sleep(2000);
			
			boolean actnewdropdown=newdropdown.isDisplayed();
			boolean expnewdropdown=true;
			
			System.out.println("**************************checkEmailTemplateOption***********"); 
			
			System.out.println("newdropdown    : "+actnewdropdown +" Value Expected  : "+expnewdropdown);
			
			if(actnewdropdown==expnewdropdown)
			{
				System.out.println("Test Pass : email Template Opened Print Screen");
				return true;
			}
			else
			{
				System.out.println("Test Fail : email Template NOT Opened Print Screen");
				return false;
			}
			
		}
		
		
		
		public boolean 	checkFieldsAndPropertiesTab() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			boolean actheaderFieldsTab=headerFieldsTab.isDisplayed();
			boolean actheaderPropertiesTab=headerPropertiesTab.isDisplayed();
			
			boolean expheaderFieldsTab=true;
			boolean expheaderPropertiesTab=true;
			
			System.out.println(" ****************************************checkFieldsAndPropertiesTab*******************************");
			
			System.out.println("headerFieldsTab        : "+actheaderFieldsTab     +"  Value Expected : "+expheaderFieldsTab);
			System.out.println("headerPropertiesTab    : "+actheaderPropertiesTab +"  Value Expected : "+expheaderPropertiesTab);
			
			
			if(actheaderFieldsTab==expheaderFieldsTab  && actheaderPropertiesTab==expheaderPropertiesTab)
			{
				System.out.println("Test Pass : Fields And Properties Are Displayed");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Fields And Properties Are NOT Displayed");
				return false;
			}
		}
		
		
		
		
		public boolean 	checkFieldsTabOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(headerFieldsTab));
			headerFieldsTab.click();
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fields_CompanyExpandBtn));
			boolean actfields_CompanyExpandBtn              = fields_CompanyExpandBtn.isDisplayed();
			boolean actfields_MasterFieldsExpandBtn         = fields_MasterFieldsExpandBtn.isDisplayed();
			boolean actfields_UserDetailsExpandBtn          = fields_UserDetailsExpandBtn.isDisplayed();
			boolean actfields_MiscellaneousExpandBtn        = fields_MiscellaneousExpandBtn.isDisplayed();
			
			boolean expfields_CompanyExpandBtn              = true;
			boolean expfields_MasterFieldsExpandBtn         = true;
			boolean expfields_UserDetailsExpandBtn          = true;
			boolean expfields_MiscellaneousExpandBtn        = true;
			
			System.out.println("*******************************************checkSaveOptionInEmailTemplate************************************"); 
			System.out.println("Fields_CompanyExpandBtn           :"+actfields_CompanyExpandBtn        +" Value Expected  :"+expfields_CompanyExpandBtn);
			System.out.println("Fields_MasterFieldsExpandBtn      :"+actfields_MasterFieldsExpandBtn   +" Value Expected  :"+expfields_MasterFieldsExpandBtn);
			System.out.println("Fields_UserDetailsExpandBtn       :"+actfields_UserDetailsExpandBtn    +" Value Expected  :"+expfields_UserDetailsExpandBtn);
			System.out.println("Fields_MiscellaneousExpandBtn     :"+actfields_MiscellaneousExpandBtn  +" Value Expected  :"+expfields_MiscellaneousExpandBtn);
			
			if(actfields_CompanyExpandBtn==expfields_CompanyExpandBtn && actfields_MasterFieldsExpandBtn==expfields_MasterFieldsExpandBtn &&
					actfields_UserDetailsExpandBtn==expfields_UserDetailsExpandBtn  && actfields_MiscellaneousExpandBtn==expfields_MiscellaneousExpandBtn)

			{
				System.out.println("Test Pass : Field Options Are Displayed");
				return true;
			}
			else
			{
				System.out.println("Test Fail : Field Options Are NOT Displayed");
				return false;
			}
		}
		
		@FindBy(xpath="//i[@class='icon-collepse icon-font8']")
		private static WebElement fields_MinimizeBtn;
		
		
		
		
		
		
		@FindBy(xpath="//p[contains(text(),'Company Name')]")
		private static WebElement draggedCompanyName;
		
		@FindBy(xpath="//div[@id='id_LayoutDiv_0']")
		private static WebElement designArea;
		
		
		public boolean checkMovingCompanyNameToDesignArea() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fields_CompanyExpandBtn));
			fields_CompanyExpandBtn.click();
			
			Thread.sleep(3000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fields_Company_CompanyName));
			fields_Company_CompanyName.click();

			getAction().dragAndDropBy(fields_Company_CompanyName, 250, -100).build().perform();
			
			//getAction().clickAndHold(fields_Company_CompanyName).dragAndDrop(fields_Company_CompanyName, designArea).build().perform();
			
			Thread.sleep(2000);
		
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(draggedCompanyName));
			boolean actdraggedCompanyName=draggedCompanyName.isDisplayed();
			boolean expdraggedCompanyName=true;
			
			
			System.out.println("******************************checkMovingCompanyNameToDesignArea******************");
			
			System.out.println("draggedCompanyName       : "+actdraggedCompanyName +"  Value Expected : "+expdraggedCompanyName);
			
			if(actdraggedCompanyName==expdraggedCompanyName)
			{
				System.out.println("Test Pass : CompanyName Moved To Design Area");
				return true;
			}
			else
			{
				System.out.println("Test Fail : CompanyName NOT Moved To Design Area");
				return false;
			}
			
		}

		
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[14]/table/tbody/tr[1]/td[2]/i")
		private static WebElement row1SplitBtn; 
		
		@FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div/div[2]/div[14]/table/tbody/tr[2]/td[2]/i")
		private static WebElement row2SplitBtn; 
				
		@FindBy(xpath="//*[@id='cmbSplitType']")
		private static WebElement splitType; 
		
		@FindBy(xpath="//*[@id='cmbSpillOver']")
		private static WebElement spillOver;
		
		
		
		@FindBy(xpath="//*[@id='SplitBudgetContent']/div[2]/div[2]/label")
		private static WebElement  splitGetAccount; 
	
		@FindBy(xpath="//*[@id='SplitBudgetContent']/div[2]/div[6]/label")
		private static WebElement  splitGetBudget; 
		
		@FindBy(xpath="//*[@id='grdSplitBudget_body']/tr[1]/td[4]")
		private static WebElement  splitGetRowOneValue; 
	
		@FindBy(xpath="//*[@id='grdSplitBudget_body']/tr[2]/td[4]")
		private static WebElement  splitGetRowTwoValue; 
	
		@FindBy(xpath="//*[@id='grdSplitBudget_body']/tr[3]/td[4]")
		private static WebElement  splitGetRowThreeValue; 
	
		@FindBy(xpath="//*[@id='grdSplitBudget_body']/tr[4]/td[4]")
		private static WebElement  splitGetRowFourValue; 
		
		@FindBy(xpath="//*[@id='lblDstAmt']")
		private static WebElement  splitGetAmount; 
		
		@FindBy(xpath="//*[@id='btnLoadSplit']")
		private static WebElement  splitSplitBtn; 
		
		@FindBy(xpath="//*[@id='btnOk']")
		private static WebElement  splitOkBtn; 
		
		@FindBy(xpath="//*[@id='splitBudgetModel']/div/div[3]/div/div[2]/div[1]/div[2]")
		private static WebElement  splitCloseBtn; 
		
		
		
		@FindBy(xpath="//*[@id='grdDefineBudget_col_2-2']")
		private static WebElement  budgetTypeR2C2SelectCell;
		
		@FindBy(xpath="//*[@id='grdDefineBudget_col_2-3']")
		private static WebElement  budgetDepartmentR2C3SelectCell; 		
		
		@FindBy(xpath="//*[@id='grdDefineBudget_col_2-6']")
		private static WebElement  budgetItemR2C4SelectCell; 
	
		@FindBy(xpath="//*[@id='grdDefineBudget_col_2-7']")
		private static WebElement  budgetAccountR2C5SelectCell; 
		
		@FindBy(xpath="//*[@id='grdDefineBudget_col_2-8']")
		private static WebElement  budgetQtyR2C6SelectCell; 
		
		@FindBy(xpath="//*[@id='grdDefineBudget_col_2-9']")
		private static WebElement  budgetValueR2C3SelectCell;
  
		@FindBy(xpath="//*[@id='grdDefineBudget_col_2-13']")
		private static WebElement  budgetStatusR2C4SelectCell; 
		
		@FindBy(xpath="//*[@id='grdDefineBudget_col_3-2']")
		private static WebElement  budgetTypeR3C2SelectCell; 
	
		@FindBy(xpath="//*[@id='grdDefineBudget_col_3-3']")
		private static WebElement  budgetDepartmentR3C3SelectCell; 		
		
		@FindBy(xpath="//*[@id='grdDefineBudget_col_3-6']")
		private static WebElement  budgetItemR3C4SelectCell; 
	
		@FindBy(xpath="//*[@id='grdDefineBudget_col_3-7']")
		private static WebElement  budgetAccountR3C5SelectCell; 
		
		@FindBy(xpath="//*[@id='grdDefineBudget_col_3-8']")
		private static WebElement  budgetQtyR3C6SelectCell; 
		
		@FindBy(xpath="//*[@id='grdDefineBudget_col_3-9']")
		private static WebElement  budgetValueR3C3SelectCell;
  
		@FindBy(xpath="//*[@id='grdDefineBudget_col_3-13']")
		private static WebElement  budgetStatusR3C4SelectCell; 
		
		
 	    public static boolean checkSplitRowOneAndRowTwoInDefineBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
		 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		 	
		 	/*try
		 	{*/ 				 		
		 	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
			    financialsMenu.click();

			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetMenu));
			    financialsBudgetMenu.click();
			 				   
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetDefineBudget));
	            financialsBudgetDefineBudget.click();
			  		   
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));     	
		 		planName.sendKeys("AuthPlanBudget");
		 		Thread.sleep(2000);
		 		versionNo.click();
		 		versionNo.sendKeys(Keys.TAB);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetTypeR1C2SelectCell));     	
		 		budgetTypeR1C2SelectCell.click();
		 		
		 		budgetTypeR1C1InputCell.click();
		 		
		 		budgetTypeR1C1InputCell.sendKeys("By Account By Value");  	
		 		Thread.sleep(2000);
		 		budgetTypeR1C1InputCell.sendKeys(Keys.TAB);
		    	
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDepartmentR1C3SelectCell));     	
		 		budgetDepartmentR1C3SelectCell.click();	 	
		 		
		 		budgetDepartmentR1C3InputCell.click();
		 
		 		budgetDepartmentR1C3InputCell.sendKeys("DUBAI");
		 		Thread.sleep(2000);
		 		budgetDepartmentR1C3InputCell.sendKeys(Keys.TAB);
		 			 				 				
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAccountR1C5SelectCell));     	
		 		budgetAccountR1C5SelectCell.click();	 	
		 	
		 		budgetAccountR1C5InputCell.click();
		 		budgetAccountR1C5InputCell.sendKeys("WA COGS ACC INV");
		 		Thread.sleep(2000);
		 		budgetAccountR1C5InputCell.sendKeys(Keys.TAB);	 				 		
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetValueR1C3SelectCell));     	
		 		budgetValueR1C3SelectCell.click();	 	
		 		Thread.sleep(2000);
		 		budgetValueR1C9InputCell.click();
		 		budgetValueR1C9InputCell.sendKeys("120");
		 		budgetValueR1C9InputCell.sendKeys(Keys.TAB);
		 		
		 		
		 		// Second Row
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetTypeR2C2SelectCell));     	
		 		budgetTypeR2C2SelectCell.click();
		 		
		 		budgetTypeR1C1InputCell.click();
		 		
		 		budgetTypeR1C1InputCell.sendKeys("By Product By Value");  	
		 		Thread.sleep(2000);
		 		budgetTypeR1C1InputCell.sendKeys(Keys.TAB);
		    	
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetDepartmentR2C3SelectCell));     	
		 		budgetDepartmentR2C3SelectCell.click();	 	
		 		
		 		budgetDepartmentR1C3InputCell.click();
		 
		 		budgetDepartmentR1C3InputCell.sendKeys("INDIA");
		 		Thread.sleep(2000);
		 		budgetDepartmentR1C3InputCell.sendKeys(Keys.TAB);
		 			 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetItemR2C4SelectCell));     	
		 		budgetItemR2C4SelectCell.click();	 	
		 		
		 		budgetItemR1C4InputCell.click();
		 		budgetItemR1C4InputCell.sendKeys("BR COGS ITEM");
		 		Thread.sleep(2000);
		 		budgetItemR1C4InputCell.sendKeys(Keys.TAB);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAccountR2C5SelectCell));     	
		 		budgetAccountR2C5SelectCell.click();	 	
		 		Thread.sleep(2000);
		 		
                String  actGetAccountLoadOnSelectItem = budgetAccountR1C5InputCell.getAttribute("value");
		 		
		 		System.out.println("The account name "+actGetAccountLoadOnSelectItem);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetAccountR1C5InputCell));
		 		budgetAccountR1C5InputCell.click();
		 	
		 		Thread.sleep(2000);
		 		budgetAccountR1C5InputCell.sendKeys(Keys.TAB);
		 				 		 			 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetValueR2C3SelectCell));     	
		 		budgetValueR2C3SelectCell.click();	 	
		 		Thread.sleep(2000);
		 		budgetValueR1C9InputCell.click();
		 		budgetValueR1C9InputCell.sendKeys("120.12");
		 		budgetValueR1C9InputCell.sendKeys(Keys.TAB);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(row1SplitBtn));  
		 		row1SplitBtn.click();
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitType));  
		 		Select spiltTypeValue=new Select(splitType);
		 		String actSplitSelected=spiltTypeValue.getFirstSelectedOption().getText();
		 				
		 		System.out.println(actSplitSelected);
		 		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(spillOver));
		 		Select getSpillValues = new Select(spillOver);
		 		String actSpillValue = getSpillValues.getFirstSelectedOption().getText();
		 		System.out.println(actSpillValue);
		 						 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitSplitBtn));  
		 		splitSplitBtn.click();
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitGetRowOneValue));  
		 		System.out.println(splitGetRowOneValue.getText());
		 		System.out.println(splitGetRowTwoValue.getText());		 		
		 		System.out.println(splitGetRowThreeValue.getText());
		 		System.out.println(splitGetRowFourValue.getText());
		 		 		
		 		System.out.println(splitGetAmount.getText());
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitOkBtn));
		 		splitOkBtn.click();		 		
		 		
		 		Thread.sleep(2000);
		 		
		 		// SECOND ROW SPLIT
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(row2SplitBtn));  
		 		row2SplitBtn.click();
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitType));  
		 		
		 		Select spiltTypeR2Value=new Select(splitType);
		 		String actSplitR2Selected=spiltTypeR2Value.getFirstSelectedOption().getText();
		 				
		 		System.out.println(actSplitR2Selected);
		 		 		
		 		
		 		Select getSpillR2Values = new Select(spillOver);
		 		String actSpillR2Value = getSpillR2Values.getFirstSelectedOption().getText();
		 		System.out.println(actSpillR2Value);
		 		Thread.sleep(2000);		
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitSplitBtn));  
		 		splitSplitBtn.click();
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitGetRowOneValue));  
		 		System.out.println(splitGetRowOneValue.getText());
		 		System.out.println(splitGetRowTwoValue.getText());		 		
		 		System.out.println(splitGetRowThreeValue.getText());
		 		System.out.println(splitGetRowFourValue.getText());
		 		 		
		 		System.out.println(splitGetAmount.getText());
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitOkBtn));
		 		splitOkBtn.click();
		 		
		 		Thread.sleep(2000);
		 			 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetSaveBtn));     	
		 		budgetSaveBtn.click();	
		 			 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
		 		String actErrorMessage=errorMessage.getText();
		 		String expErrorMessage="Budget Saved succesfully";

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
		 		errorMessageCloseBtn.click();
		 				 		
		 		System.out.println("ValidationMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage);
		 				 				 			 		
		 		if( actErrorMessage.equalsIgnoreCase(expErrorMessage))
		 		{
		 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
		 			return true;
		 		}
		 		else
		 		{
		 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
		 			return false;
		 		}
		 	}
		 	/*catch (Exception e) 
		     {
		 		 String exception =e.getMessage().substring(0, 50);
		         excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
		         System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
		         return false;
		     }
		  }*/

		

 	   @FindBy(xpath="//*[@id='btnSendBudget']")
 	   public static WebElement budgetSendToAuth;
 	    
 	   @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[2]/td[3]")
 	   public static WebElement budgetTypeR2GetText;
 		 
 	   @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[2]/td[4]")
       public static WebElement budgetDepartmentR2GetText;
 		 	
 	   @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[2]/td[7]")
       public static WebElement budgetItemR2GetText;
 		 
 	   @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[2]/td[8]")
       public static WebElement budgetAccountR2GetText;
 		 
       @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[2]/td[9]")
 	   public static WebElement budgetQtyR2GetText;	
 		 
       @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[2]/td[10]")
 	   public static WebElement budgetValueR2GetText;
 		
 	   @FindBy(xpath="//*[@id='grdDefineBudget_body']/tr[2]/td[14]")
 	   public static WebElement budgetStatusR2GetText;
  
 	   public static boolean checSavedBudgetInfoThroughSplitOptionDefineBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	   {
		 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		 	
		 	/*try
		 	{*/ 				 		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planName));     	
		 		planName.sendKeys("AuthPlanBudget");
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(versionNo)); 
		 		versionNo.click();
		 		
		 		System.out.println(versionNo.getAttribute("value"));	 	
		 	 			 		
		 		String actBudgetType     = budgetTypeGetText.getText();
		 		String actDepartment     = budgetDepartmentGetText.getText();
		 		String actAccount        = budgetAccountGetText.getText();
		 		boolean actItem          = budgetItemGetText.getText().isEmpty();
		 		String actValue          = budgetValueGetText.getText();
		 		String actBudgetStatus   = budgetStatusGetText.getText();
		 				 		
		 		String expBudgetType     =  "By Account By Value";
		 		String expDepartment     = "DUBAI";
		 		String expAccount        = "WA COGS ACC INV";
		 		boolean expItem           = true;
		 		String expValue          = "120.00";
		 		String expBudgetStatus   = "Pending";
		 		
		    	System.out.println("Budget Menu Confirm Budget        : " + actBudgetType +       "  value expected  " + expBudgetType);
		    	System.out.println("Budget Menu Revise Budget         : " + actDepartment +       "  value expected  " + expDepartment);
		    	System.out.println("Budget Menu Approval Budget       : " + actAccount +          "  value expected  " + expAccount);
		    	System.out.println("Budget Menu Rejected Budget       : " + actItem +             "  value expected  " + expItem);
		    	System.out.println("Budget Menu Revise Budget Budget  : " + actValue +            "  value expected  " + expValue);
		    	System.out.println("Budget Menu Budget Report Budget  : " + actBudgetStatus +     "  value expected  " + actBudgetStatus);
		     	    			    	
		 		String actBudgetTypes     = budgetTypeR2GetText.getText();
		 		String actDepartments     = budgetDepartmentR2GetText.getText();
		 		String actAccounts        = budgetAccountR2GetText.getText();
		 		String actItems           = budgetItemR2GetText.getText();
		 		String actValues          = budgetValueR2GetText.getText();
		 		String actBudgetStatuss   = budgetStatusR2GetText.getText();
		 		
		 	
		 		String expBudgetTypes     = "By Product By Value";
		 		String expDepartments     = "INIDA";
		 		String expAccounts        = "BR COGS ACC INV";
		 		String expItems           = "BR COGS ITEM";
		 		String expValues          = "120.12";
		 		String expBudgetStatuss   = "Pending";
		 
		    	System.out.println("Budget Menu Confirm Budget        : " + actBudgetTypes +       "  value expected  " + expBudgetTypes);
		    	System.out.println("Budget Menu Revise Budget         : " + actDepartments +       "  value expected  " + expDepartments);
		    	System.out.println("Budget Menu Approval Budget       : " + actAccounts +          "  value expected  " + expAccounts);
		    	System.out.println("Budget Menu Rejected Budget       : " + actItems +             "  value expected  " + expItems);
		    	System.out.println("Budget Menu Revise Budget Budget  : " + actValues +            "  value expected  " + expValues);
		    	System.out.println("Budget Menu Budget Report Budget  : " + actBudgetStatuss +     "  value expected  " + expBudgetStatuss);	    	    
		    	
		    	Thread.sleep(2000);
		    	
		    	// TO CHECK SPLIT SCREEN
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(row1SplitBtn));  
		 		row1SplitBtn.click();
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitType));  
		 		Select spiltTypeValue=new Select(splitType);
		 		String actSplitSelected=spiltTypeValue.getFirstSelectedOption().getText();
		 				
		 		System.out.println(actSplitSelected);
		 		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(spillOver));
		 		Select getSpillValues = new Select(spillOver);
		 		String actSpillValue = getSpillValues.getFirstSelectedOption().getText();
		 		System.out.println("SplitValues     : " + actSpillValue);
		 						 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitGetRowOneValue));  
		 		System.out.println("Row One Value   : " + splitGetRowOneValue.getText());
		 		System.out.println("Row Two Value   : " + splitGetRowTwoValue.getText());		 		
		 		System.out.println("Row Three Value : " + splitGetRowThreeValue.getText());
		 		System.out.println("Row Four Value  : " + splitGetRowFourValue.getText());
		 		 		
		 		System.out.println("Row Split Amount : " + splitGetAmount.getText());
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitOkBtn));
		 		splitOkBtn.click();		 		
		 		
		 		Thread.sleep(5000);
		 		
		 		// SECOND ROW SPLIT
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(row2SplitBtn));  
		 		row2SplitBtn.click();
		 		
		 		Thread.sleep(5000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitType));  
		 		
		 		Select spiltTypeR2Value=new Select(splitType);
		 		String actSplitR2Selected=spiltTypeR2Value.getFirstSelectedOption().getText();
		 				
		 		System.out.println(actSplitR2Selected);
		 		 			 		
		 		Select getSpillR2Values = new Select(spillOver);
		 		String actSpillR2Value = getSpillR2Values.getFirstSelectedOption().getText();
		 		System.out.println(actSpillR2Value);
		 		Thread.sleep(2000);		
		 		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitGetRowOneValue));  
		 		System.out.println(splitGetRowOneValue.getText());
		 		System.out.println(splitGetRowTwoValue.getText());		 		
		 		System.out.println(splitGetRowThreeValue.getText());
		 		System.out.println(splitGetRowFourValue.getText());
		 		 		
		 		System.out.println(splitGetAmount.getText());
		 		
		 		Thread.sleep(2000);
		 		
		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(splitOkBtn));
		 		splitOkBtn.click();
		 			 	
		 		Thread.sleep(2000);
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetSendToAuth));     	
		    	budgetSendToAuth.click();
		    	
		    	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
		 		String actErrorMessages=errorMessage.getText();
		 		String expErrorMessages="Budget Confirmed successfully";

		 		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
		 		errorMessageCloseBtn.click();
		 		
		 		
		 		System.out.println("ValidationMessage  :  "+actErrorMessages +" Value Expected : "+expErrorMessages);
                 		 		
		 			 		
		 			 		
		 		if( true)
		 		{
		 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
		 			return true;
		 		}
		 		else
		 		{
		 			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
		 			return false;
		 		}
		 	}
		 	/*catch (Exception e) 
		     {
		 		 String exception =e.getMessage().substring(0, 50);
		         excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
		         System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
		         return false;
		     }
		  }*/

		

 	   // AUTHORIZATION OF THE BUDGETS IN AUTHORIZE USER
 	   
		
 	   
 	   



 public boolean checkAlertsInWrapperLayoutAuthorizeUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
 {
	excelReader=new ExcelReader(POJOUtility.getExcelPath());
	xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";

    /*try
	{*/
		
    	    getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
		    userNameDisplay.click();
		
		    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		    logoutOption.click();
		 
			boolean verifyFiledsStatus=true;
			
			getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
						
			boolean actusername              = username.isDisplayed();
			boolean actpassword              = password.isDisplayed();
			boolean actcompanyDropDownList   = companyDropDownList.isDisplayed();
			boolean actsignIn                = signIn.isDisplayed();
			
			// Expected
			boolean expusername              = true;
			boolean exppassword              = true;
			boolean expcompanyDropDownList   = true;
			boolean expsignIn                = true;
			
			System.out.println("************************************* checkSignOut  *********************************");
			System.out.println("username             : "+actusername             +" Value Expected : "+expusername);
			System.out.println("password             : "+actpassword             +" Value Expected : "+exppassword);
			System.out.println("companyDropDownList  : "+actcompanyDropDownList  +" Value Expected : "+expcompanyDropDownList);
			System.out.println("signIn               : "+actsignIn               +" Value Expected : "+expsignIn);

	
		
		  getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(username));
		  username.click();
		 
		  username.sendKeys(excelReader.getCellData("PurchaseVoucherN", 252, 6));
		  getAction().moveToElement(username).sendKeys(Keys.TAB).perform();
		 
		  getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(password));
		  password.sendKeys(excelReader.getCellData("PurchaseVoucherN", 253, 6));
		 
		  getAction().moveToElement(password).sendKeys(Keys.TAB).perform();
		  
		  getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(signIn));
		  signIn.click();
	 

	
		Thread.sleep(5000);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
		String actUserInfo=userNameDisplay.getText();
		
		System.out.println("User Info : "+actUserInfo);
		
		System.out.println("User Info Capture Text :"+userNameDisplay.getText());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
		
		companyLogo.click();
		
		String actCompanyTxt=companyName.getText();
		String actLoginCompanyName=actCompanyTxt.substring(0, 19);
		System.out.println("company name :"+ actLoginCompanyName);
		companyLogo.click();
					    	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
		String getDashboard=labelDashboard.getText();
			
		boolean getnoOfAlertsDisplay=noOfAlertsDisplay.isDisplayed();
				
		// Expected
		String expuserInfo="UserAllOptionsST";
		String expLoginCompanyName="Automation Company ";
		String expDashboard="Dashboard";
		boolean expnoOfAlertsDisplay=true;
		
		System.out.println("************************************* checkAlertsToSignIAsAuthorizeUser  *********************************");
		System.out.println("userNameDisplay    : "+actUserInfo              +" Value Expected : "+expuserInfo);
		System.out.println("LoginCompanyName   : "+actLoginCompanyName      +" Value Expected : "+expLoginCompanyName);
		System.out.println("labelDashboard     : "+getDashboard             +" Value Expected : "+expDashboard);
		System.out.println("noOfAlertsDisplay  : "+getnoOfAlertsDisplay     +" Value Expected : "+expnoOfAlertsDisplay);
		
		
		boolean actMethod=actUserInfo.equalsIgnoreCase(expuserInfo) && actLoginCompanyName.equalsIgnoreCase(expLoginCompanyName)
				&& getDashboard.equalsIgnoreCase(expDashboard) && getnoOfAlertsDisplay==expnoOfAlertsDisplay;
		
		String actResult=Boolean.toString(actMethod);
		String expResult=excelReader.getCellData("PurchaseVoucherN", 254, 7);
		
		excelReader.setCellData(xlfile, "PurchaseVoucherN", 254, 8, actResult);
		
		
		if(actMethod==true)
		{
			System.out.println("Test Pass : Transaction Authorization is Displayed");
			excelReader.setCellData(xlfile, "PurchaseVoucherN", 251, 9, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Transaction Authorization is Displayed");
			excelReader.setCellData(xlfile, "PurchaseVoucherN", 251, 9, resFail);
			return false;
		}
	}
	/*catch (Exception e) 
	{
		String exception =e.getMessage();
		excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 251, 10, exception);
		System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ exception);
		return false;
	}
  }*/
 
 
 
  @FindBy(xpath="//span[@id='noofalerts']")
  private static WebElement noOfAlertsDisplay;

  @FindBy(xpath="//li[@class='header']")
  private static WebElement alertsMainHeader;

  @FindBy(xpath="//li[@class='treeview-menu']//a[@id='1']")
  private static WebElement alertTransactionsTreeView;

  @FindBy(xpath="//*[@id='28']/span")
  private static WebElement alertBudget;

  @FindBy(xpath="//*[@id='0']/span")
  private static WebElement budgetAuthorization;
  
  @FindBy(xpath="//*[@id='masterChild_0']/a/span")
  private static WebElement budgetAuthorizationText;
  
  @FindBy(xpath="//*[@id='id_PlanName 1']/label")
  private static WebElement pendingBudgetInfo;

  @FindBy(xpath="//*[@id='Grid_ApproveBudgetMore_control_heading_ctrl_1']")
  private static WebElement checkboxInBudgetAuthorization;
  
  @FindBy(xpath="//*[@id='btnAuthorize']")
  private static WebElement btnAuthorizeInBudgetAuthorization;
  
  @FindBy(xpath="//*[@id='btnClose']")
  private static WebElement btnCloseInBudgetAuthorization;
  
  
  public boolean checkClickOnAlertsInWrapperLayoutAuthorizeUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
  {
   excelReader=new ExcelReader(POJOUtility.getExcelPath());
   xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
   /*try
   {*/
		
	getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(noOfAlertsDisplay));
	noOfAlertsDisplay.click();
	
	Thread.sleep(2000);
	
	// Actual
	String  actAlertsMainHeader=alertsMainHeader.getText();
	excelReader.setCellData(xlfile, "PurchaseVoucherN", 256, 8, actAlertsMainHeader);
	boolean actNoOfAlertsDisplay=alertsMainHeader.isDisplayed();
	boolean expNoOfAlertsDisplay=true;
	System.out.println("Alerts is display"+ actAlertsMainHeader);
	
	String actAlertTransactionsTreeView=alertBudget.getText();
	String expAlertTransactionsTreeView="  Budgeting (2)";

	alertBudget.click();
	
	String actBudgetAuthorization = budgetAuthorization.getText();
	String expBudgetAuthorization = " Define Budget(s) to Authorization (2)";
	
	budgetAuthorization.click();
	
	String actBudgetText =budgetAuthorizationText.getText();
	String actBudgetSubString = actBudgetText.substring(0,33);
	String expBudgetText = "Alerts TemplateAutomation Company";
	
	budgetAuthorizationText.click();
	
	Thread.sleep(2000);
	
	System.out.println("************************************* checkClickOnAlertsInAuthorizeUser  *********************************");
	System.out.println("noOfAlertsDisplay         : " + actNoOfAlertsDisplay         + " Value Expected  : " + expNoOfAlertsDisplay);
	System.out.println("noOfAlertsDisplay         : " + actAlertTransactionsTreeView + " Value Expected  : " + expAlertTransactionsTreeView);
	System.out.println("noOfAlertsDisplay         : " + actBudgetAuthorization       + " Value Expected  : " + expBudgetAuthorization);
	System.out.println("noOfAlertsDisplay         : " + actBudgetSubString           + " Value Expected  : " + expBudgetText);
	
	if(actNoOfAlertsDisplay==expNoOfAlertsDisplay && actAlertTransactionsTreeView.equalsIgnoreCase(expAlertTransactionsTreeView)
			&& actBudgetAuthorization.equalsIgnoreCase(expBudgetAuthorization) && actBudgetSubString.equalsIgnoreCase(expBudgetText))
	{
		System.out.println("Test Pass : Transaction Authorization is Displayed");
		excelReader.setCellData(xlfile, "PurchaseVoucherN", 255, 9, resPass);
		return true;
		
	}
	else
	{
		System.out.println("Test Fail : Transaction Authorization is Displayed");
		excelReader.setCellData(xlfile, "PurchaseVoucherN", 255, 9, resFail);
		return false;
	
	}
	}
	/*catch (Exception e) 
	{
		String exception =e.getMessage();
		excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 255, 10, exception);
		System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ exception);
   	 	return false;
	}
}
*/

  
  
  public boolean checkPendingBudgetInAuthorizeUser() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
  {
   excelReader=new ExcelReader(POJOUtility.getExcelPath());
   xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
   /*try
   {*/
		
	getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(pendingBudgetInfo));
    boolean actPendingBudgetIsDisplay =pendingBudgetInfo.isDisplayed();
    boolean expPendingBudgetIsDisplay =true;

	pendingBudgetInfo.click();
	
	Thread.sleep(2000);
	
    checkboxInBudgetAuthorization.click();
	Thread.sleep(2000);
	btnAuthorizeInBudgetAuthorization.click();
	
	
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(budgetSendToAuth));     	
	budgetSendToAuth.click();
	*/
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
	String actErrorMessages=errorMessage.getText();
	String expErrorMessages="Approved Successfully";
	
	System.out.println("ValidationMessage  :  "+actErrorMessages +" Value Expected : "+expErrorMessages);
		
	if(actErrorMessages.equalsIgnoreCase(expErrorMessages))
	{
		System.out.println("Test Pass : Transaction Authorization is Displayed");
		excelReader.setCellData(xlfile, "PurchaseVoucherN", 255, 9, resPass);
		return true;
		
	}
	else
	{
		System.out.println("Test Fail : Transaction Authorization is Displayed");
		excelReader.setCellData(xlfile, "PurchaseVoucherN", 255, 9, resFail);
		return false;
	
	}
	}
	/*catch (Exception e) 
	{
		String exception =e.getMessage();
		excelReader.setExceptionInExcel(xlfile, "PurchaseVoucherN", 255, 10, exception);
		System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ exception);
   	 	return false;
	}
   }*/



  
    public boolean checkCloseButtonInAuthorizeBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
    {
	 	excelReader=new ExcelReader(POJOUtility.getExcelPath());
		xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
		
		
		if(btnCloseInBudgetAuthorization.isDisplayed())
		{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(btnCloseInBudgetAuthorization));
		btnCloseInBudgetAuthorization.click();
		}
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
		
		boolean actdashboard=dashboard.isDisplayed();
		boolean expdashboard=true;
			
		System.out.println("dashboard  :"+actdashboard +"  Value Expected  : "+expdashboard);
		
		if(actdashboard==expdashboard)	
		{
			System.out.println("Test Pass : Cancel Option Working ");
			excelReader.setCellData(xlfile, xlSheetName, 53, 9, resPass);
			return true;
		}
		else
		{
			System.out.println("Test Fail : Cancel Option Not Working");
			excelReader.setCellData(xlfile, xlSheetName, 53, 9, resFail);
			return false;
		}
      }
 	   
 	    
    
    // REVISE BUDGET DETAILS REPORT
    
    @FindBy(xpath="//*[@id='ReviseBudgetType']")
    private static WebElement reviseBudgetType;
    
    @FindBy(xpath="//*[@id='page_Content']/div[1]/div/div[2]/div[1]/div[7]/ul/li/span")
    private static WebElement refreshBtn;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr/td[3]")
    private static WebElement revisePlanName;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr/td[4]")
    private static WebElement reviseAccount;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr/td[5]")
    private static WebElement reviseItem;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr/td[6]")
    private static WebElement reviseDepartment;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr/td[7]")
    private static WebElement reviseQty;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr/td[8]")
    private static WebElement reviseValue;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr/td[11]")
    private static WebElement reviseQtyAdd;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr/td[12]")
    private static WebElement reviseQtyReduce;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr/td[13]")
    private static WebElement reviseStatus;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr/td[14]")
    private static WebElement reviseReduceValue;
    
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr[2]/td[3]")
    private static WebElement revisePlanNameR2;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr[2]/td[4]")
    private static WebElement reviseAccountR2;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr[2]/td[5]")
    private static WebElement reviseItemR2;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr[2]/td[6]")
    private static WebElement reviseDepartmentR2;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr[2]/td[9]")
    private static WebElement reviseQtyR2;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr[2]/td[10]")
    private static WebElement reviseValueR2;
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr[2]/td[13]")
    private static WebElement reviseStatusR2;
    
    public static boolean checkReviseBudgetDetails() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	 excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
	 /*try
	 {*/
   	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsMenu));
	   financialsMenu.click();

	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetMenu));
	   financialsBudgetMenu.click();
		 
	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsBudgetReviseBudgetDetailsBudget));
	   financialsBudgetReviseBudgetDetailsBudget.click();
	  	   
	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reviseBudgetType));
	   reviseBudgetType.sendKeys("Append Budget");
	   
	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(refreshBtn));
	   refreshBtn.click();
  
       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));
	   String actRevisePlanName        = revisePlanName.getText();
	   String actReviseAccount         = reviseAccount.getText();
	   String actReviseDepartment      = reviseDepartment.getText();
	  // String actReviseItem            = reviseItem.getText();		 
	   String actReviseQty             = reviseQty.getText();
	   String actReviseValue           = reviseValue.getText();
	   String actReviseStatus          = reviseStatus.getText();
	   	 
	   String expRevisePlanName        = "ConfirmPlanBudget";
	   String expReviseAccount         = "1";
	   String expReviseDeparment       = "SU";
	   //String expReviseItem            = "SU";	   
	   String expReviseQty             = "24.00";
	   String expReviseValue           = "240.00";
	   String expReviseStatus          = "Authorized";
	 	
	   System.out.println("Plan Name         : " + actRevisePlanName +     "  value expected  " + expRevisePlanName);
	   System.out.println("Revision          : " + actReviseAccount +      "  value expected  " + expReviseAccount);
	   System.out.println("Revision By       : " + actReviseDepartment +   "  value expected  " + expReviseDeparment);
	   //System.out.println("Revision By       : " + actReviseItem +         "  value expected  " + expReviseItem);
	   System.out.println("Qty               : " + actReviseQty +          "  value expected  " + expReviseQty);
	  System.out.println("Value             : " + actReviseValue +        "  value expected  " + expReviseValue);
	  // System.out.println("Status            : " + actReviseStatus +       "  value expected  " + expReviseStatus);
	
	  
	   
		if(actRevisePlanName.equalsIgnoreCase(expRevisePlanName) && actReviseAccount.equalsIgnoreCase(expReviseAccount)
				&& actReviseDepartment.equalsIgnoreCase(expReviseDeparment)
				&& actReviseQty.equalsIgnoreCase(expReviseQty) && actReviseValue.equalsIgnoreCase(expReviseValue))				
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
			return false;
		}
	}
	/*catch (Exception e) 
    {
	   String exception =e.getMessage().substring(0, 50);
       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
       return false;
    }
  }*/
    
 	    
    
    public static boolean checkAddReduceReviseBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	 excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
	 /*try
	 { */  	  
	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reviseBudgetType));
	   reviseBudgetType.sendKeys("Add/Reduce Budget");
	   
	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(refreshBtn));
	   refreshBtn.click();	   
	  
       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));
	   String actRevisePlanName        = revisePlanName.getText();
	   String actReviseAccount         = reviseAccount.getText();
	   String actReviseDepartment      = reviseDepartment.getText();
	   //String actReviseItem            = reviseItem.getText();		 
	   String actReviseQty             = reviseQtyAdd.getText();
	   String actReviseValue           = reviseQtyReduce.getText();
	   String actReviseStatus          = reviseStatus.getText();
	   String actReviseReduceValue     = reviseReduceValue.getText();
	   
	   String expRevisePlanName        = "ConfirmPlanBudget";
	   String expReviseAccount         = "2";
	   String expReviseDeparment       = "SU";
	  // String expReviseItem            = "BR COGS ITEM";	   
	   String expReviseQty             = "6.00";
	   String expReviseValue           = "-6.00";
	   String expReviseStatus          = "120.00";
	   String expReviseReduceValue     = "-60.00";
	 	   	   	
	   System.out.println("Budget Menu Define Budget         : " + actRevisePlanName +     "  value expected  " + expRevisePlanName);
	   System.out.println("Budget Menu Confirm Budget        : " + actReviseAccount +      "  value expected  " + expReviseAccount);
	   System.out.println("Budget Menu Budget Budget         : " + actReviseDepartment +   "  value expected  " + expReviseDeparment);
	   //System.out.println("Budget Menu Approval Budget       : " + actReviseItem +         "  value expected  " + expReviseItem);
	   System.out.println("Budget Menu Budget Report Budget  : " + actReviseQty +          "  value expected  " + expReviseQty);
	   System.out.println("Budget Menu Define Budget         : " + actReviseValue +        "  value expected  " + expReviseValue);
	   System.out.println("Budget Menu Confirm Budget        : " + actReviseStatus +       "  value expected  " + expReviseStatus);
	   System.out.println("Budget Menu Confirm Budget        : " + actReviseReduceValue +  "  value expected  " + expReviseReduceValue);
	   
	   
	  /* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));
	   String actRevisePlanNameR2        = revisePlanNameR2.getText();
	   String actReviseAccountR2         = reviseAccountR2.getText();
	   String actReviseDepartmentR2      = reviseDepartmentR2.getText();
	   String actReviseItemR2            = reviseItemR2.getText();		 
	   String actReviseQtyR2             = reviseQtyR2.getText();
	   String actReviseValueR2            = reviseValueR2.getText();
	   String actReviseStatusR2          = reviseStatusR2.getText();
	   	 
	   String expRevisePlanNameR2        = "ConfirmPlanBudget";
	   String expReviseAccountR2         = "STD RATE COGS ACC INV";
	   String expReviseDeparmentR2       = "DUBAI";
	   String expReviseItemR2            = "STD RATE COGS ITEM";	   
	   String expReviseQtyR2             = "6.00";
	   String expReviseValueR2           = "120.00";
	   String expReviseStatusR2          = "Authorized";
	 	
	   System.out.println("Budget Menu Define Budget         : " + actRevisePlanNameR2 +     "  value expected  " + expRevisePlanNameR2);
	   System.out.println("Budget Menu Confirm Budget        : " + actReviseAccountR2 +      "  value expected  " + expReviseAccountR2);
	   System.out.println("Budget Menu Budget Budget         : " + actReviseDepartmentR2 +   "  value expected  " + expReviseDeparmentR2);
	   System.out.println("Budget Menu Approval Budget       : " + actReviseItemR2 +         "  value expected  " + expReviseItemR2);
	   System.out.println("Budget Menu Budget Report Budget  : " + actReviseQtyR2 +          "  value expected  " + expReviseQtyR2);
	   System.out.println("Budget Menu Define Budget         : " + actReviseValueR2 +        "  value expected  " + expReviseValueR2);
	   System.out.println("Budget Menu Confirm Budget        : " + actReviseStatusR2 +       "  value expected  " + expReviseStatusR2);
*/	   
		if(actRevisePlanName.equalsIgnoreCase(expRevisePlanName) && actReviseAccount.equalsIgnoreCase(expReviseAccount)
				&& actReviseDepartment.equalsIgnoreCase(expReviseDeparment) /*&& actReviseItem.equalsIgnoreCase(expReviseItem)*/
				&& actReviseQty.equalsIgnoreCase(expReviseQty) && actReviseValue.equalsIgnoreCase(expReviseValue)
				&& actReviseStatus.equalsIgnoreCase(expReviseStatus)
				
				/*&& actRevisePlanNameR2.equalsIgnoreCase(expRevisePlanNameR2) && actReviseAccountR2.equalsIgnoreCase(expReviseAccountR2)
				&& actReviseDepartmentR2.equalsIgnoreCase(expReviseDeparmentR2) && actReviseItemR2.equalsIgnoreCase(expReviseItemR2)
				&& actReviseQtyR2.equalsIgnoreCase(expReviseQtyR2) && actReviseValueR2.equalsIgnoreCase(expReviseValueR2)
				&& actReviseStatusR2.equalsIgnoreCase(expReviseStatusR2)*/)				
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
			return false;
		}
	}
	/*catch (Exception e) 
    {
	   String exception =e.getMessage().substring(0, 50);
       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
       return false;
    }
  }*/
  
    
    @FindBy(xpath="//*[@id='ReviseHome_Grid_body']/tr[1]/td[2]/input")
    private static WebElement reviseGridBodyFirstCheckBox;
    
    @FindBy(xpath="//*[@id='myNavbar']/ul/li[3]/div")
    private static WebElement reviseViewOption;
    
    @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div[2]/div/div[1]/div[1]/div[2]/table/tbody/tr[2]/td[1]/input[2]")
    private static WebElement revisePagePlanName;
    
    @FindBy(xpath="//*[@id='grdAddEditBudget_col_1-4']")
    private static WebElement revisePageDepartmentName;
    
    @FindBy(xpath="//*[@id='grdAddEditBudget_col_1-7']")
    private static WebElement revisePageItemName;
    
    @FindBy(xpath="//*[@id='grdAddEditBudget_col_1-8']")
    private static WebElement revisePageAccountName;

    @FindBy(xpath="//*[@id='grdAddEditBudget_col_1-9']")
    private static WebElement revisePageQtyName;
    
    @FindBy(xpath="//*[@id='grdAddEditBudget_col_1-10']")
    private static WebElement revisePageValueName;

    @FindBy(xpath="//*[@id='grdAddEditBudget_col_2-4']")
    private static WebElement revisePageDepartmentNameR2;
    
    @FindBy(xpath="//*[@id='grdAddEditBudget_col_2-7']")
    private static WebElement revisePageItemNameR2;
    
    @FindBy(xpath="//*[@id='grdAddEditBudget_col_2-8']")
    private static WebElement revisePageAccountNameR2;

    @FindBy(xpath="//*[@id='grdAddEditBudget_col_2-9']")
    private static WebElement revisePageQtyNameR2;
    
    @FindBy(xpath="//*[@id='grdAddEditBudget_col_2-10']")
    private static WebElement revisePageValueNameR2;
    
    @FindBy(xpath="/html/body/section/div[2]/div/section[1]/div/div/div[1]/nav/div/div[2]/ul/li/div/div[2]")
    private static WebElement revisePageValueCancel;
    
    public static boolean checkViewOptionInReviseBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	 excelReader=new ExcelReader(POJOUtility.getExcelPath());
	 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
	
	 /*try
	 {*/
		 
	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reviseBudgetType));
   	   reviseBudgetType.sendKeys("Transfer Budget");
   	   
   	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(refreshBtn));
   	   refreshBtn.click();
	   	   
   	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reviseGridBodyFirstCheckBox));
   	   reviseGridBodyFirstCheckBox.click();

	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reviseViewOption));
	   reviseViewOption.click();
		   		   
	   Thread.sleep(2000);
	   
	   /*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(revisePagePlanName));
	   String actRevisePlanName       = revisePagePlanName.getText();*/
       String actReviseDepartmentName = revisePageDepartmentName.getText();
       String actReviseItemName       = revisePageItemName.getText();
       String actReviseAccountName    = revisePageAccountName.getText();
       String actReviseQtyName        = revisePageQtyName.getText();
       String actReviseValueName      = revisePageValueName.getText();
   
	   String expRevisePlanName       = "ConfirmPlanBudget";
       String expReviseDepartmentName = "DUBAI";
       String expReviseItemName       = "BR COGS ITEM";
       String expReviseAccountName    = "BR COGS ACC INV";
       String expReviseQtyName        = "-2.00";
       String expReviseValueName      = "-20.00";
	   
	 
	   //System.out.println("Append Budget          : " + actRevisePlanName +        "  value expected  " + expRevisePlanName);
	   System.out.println("Add Or Reduce Budget   : " + actReviseDepartmentName +  "  value expected  " + expReviseDepartmentName);
	   System.out.println("Transfer Budget        : " + actReviseItemName +        "  value expected  " + expReviseItemName);
	   System.out.println("Append Budget          : " + actReviseAccountName +     "  value expected  " + expReviseAccountName);
	   System.out.println("Add Or Reduce Budget   : " + actReviseQtyName +         "  value expected  " + expReviseQtyName);
	   System.out.println("Transfer Budget        : " + actReviseValueName +       "  value expected  " + expReviseValueName);
	 
	   
	   
	   
       String actReviseDepartmentNameR2 = revisePageDepartmentNameR2.getText();
       String actReviseItemNameR2       = revisePageItemNameR2.getText();
       String actReviseAccountNameR2    = revisePageAccountNameR2.getText();
       String actReviseQtyNameR2        = revisePageQtyNameR2.getText();
       String actReviseValueNameR2      = revisePageValueNameR2.getText();
   
	   String expRevisePlanNameR2       = "ConfirmPlanBudget";
       String expReviseDepartmentNameR2 = "DUBAI";
       String expReviseItemNameR2       = "STD RATE COGS ITEM";
       String expReviseAccountNameR2    = "STD RATE COGS ACC INV";
       String expReviseQtyNameR2        = "2.00";
       String expReviseValueNameR2      = "20.00";
	   
	 
	   
	   System.out.println("Add Or Reduce Budget   : " + actReviseDepartmentNameR2 +  "  value expected  " + expReviseDepartmentNameR2);
	   System.out.println("Transfer Budget        : " + actReviseItemNameR2 +        "  value expected  " + expReviseItemNameR2);
	   System.out.println("Append Budget          : " + actReviseAccountNameR2 +     "  value expected  " + expReviseAccountNameR2);
	   System.out.println("Add Or Reduce Budget   : " + actReviseQtyNameR2 +         "  value expected  " + expReviseQtyNameR2);
	   System.out.println("Transfer Budget        : " + actReviseValueNameR2 +       "  value expected  " + expReviseValueNameR2);
	 
	   
	   Thread.sleep(5000);
	   revisePageValueCancel.click();
	   
		if(actReviseDepartmentName.equalsIgnoreCase(expReviseDepartmentName)
				&& actReviseItemName.equalsIgnoreCase(expReviseItemName) && actReviseAccountName.equalsIgnoreCase(expReviseAccountName)
				&& actReviseQtyName.equalsIgnoreCase(expReviseQtyName) && actReviseValueName.equalsIgnoreCase(expReviseValueName)
				
				&& actReviseDepartmentNameR2.equalsIgnoreCase(expReviseDepartmentNameR2)
				&& actReviseItemNameR2.equalsIgnoreCase(expReviseItemNameR2) && actReviseAccountNameR2.equalsIgnoreCase(expReviseAccountNameR2)
				&& actReviseQtyNameR2.equalsIgnoreCase(expReviseQtyNameR2) && actReviseValueNameR2.equalsIgnoreCase(expReviseValueNameR2))				
		{	   
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
			return true;
		}
		else
		{	   
			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
			return false;
		}
	}
	/*catch (Exception e) 
    {
	   String exception =e.getMessage().substring(0, 50);
       excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
       System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
       return false;
    }
  }*/
	
    
    
    
    public static boolean checkTransferReviseBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
   	{
   	 excelReader=new ExcelReader(POJOUtility.getExcelPath());
   	 xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
   	
   	 /*try
   	 { */  	  
   	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reviseBudgetType));
   	   reviseBudgetType.sendKeys("Transfer Budget");
   	   
   	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(refreshBtn));
   	   refreshBtn.click();	   
   	  
       getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));
   	   String actRevisePlanName        = revisePlanName.getText();
   	   String actReviseAccount         = reviseAccount.getText();
   	   String actReviseDepartment      = reviseDepartment.getText();
   	   String actReviseItem            = reviseItem.getText();		 
   	   String actReviseQty             = reviseQty.getText();
   	   String actReviseValue           = reviseValue.getText();
   	   String actReviseStatus          = reviseStatus.getText();
   	   	 
   	   String expRevisePlanName        = "ConfirmPlanBudget";
   	   String expReviseAccount         = "BR COGS ACC INV";
   	   String expReviseDeparment       = "DUBAI";
   	   String expReviseItem            = "BR COGS ITEM";	   
   	   String expReviseQty             = "-2.00";
   	   String expReviseValue           = "-20.00";
   	   String expReviseStatus          = "Authorized";
   	 	   	   	
   	   System.out.println("Budget Menu Define Budget         : " + actRevisePlanName +     "  value expected  " + expRevisePlanName);
   	   System.out.println("Budget Menu Confirm Budget        : " + actReviseAccount +      "  value expected  " + expReviseAccount);
   	   System.out.println("Budget Menu Budget Budget         : " + actReviseDepartment +   "  value expected  " + expReviseDeparment);
   	   System.out.println("Budget Menu Approval Budget       : " + actReviseItem +         "  value expected  " + expReviseItem);
   	   System.out.println("Budget Menu Budget Report Budget  : " + actReviseQty +          "  value expected  " + expReviseQty);
   	   System.out.println("Budget Menu Define Budget         : " + actReviseValue +        "  value expected  " + expReviseValue);
   	   System.out.println("Budget Menu Confirm Budget        : " + actReviseStatus +       "  value expected  " + expReviseStatus);
   	
   	   
   	   
   	   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(planNameTxt));
   	   String actRevisePlanNameR2        = revisePlanNameR2.getText();
   	   String actReviseAccountR2         = reviseAccountR2.getText();
   	   String actReviseDepartmentR2      = reviseDepartmentR2.getText();
   	   String actReviseItemR2            = reviseItemR2.getText();		 
   	   String actReviseQtyR2             = reviseQtyR2.getText();
   	   String actReviseValueR2            = reviseValueR2.getText();
   	   String actReviseStatusR2          = reviseStatusR2.getText();
   	   	 
   	   String expRevisePlanNameR2        = "ConfirmPlanBudget";
   	   String expReviseAccountR2         = "STD RATE COGS ACC INV";
   	   String expReviseDeparmentR2       = "DUBAI";
   	   String expReviseItemR2            = "STD RATE COGS ITEM";	   
   	   String expReviseQtyR2             = "2.00";
   	   String expReviseValueR2           = "20.00";
   	   String expReviseStatusR2          = "Authorized";
   	 	
   	   System.out.println("Budget Menu Define Budget         : " + actRevisePlanNameR2 +     "  value expected  " + expRevisePlanNameR2);
   	   System.out.println("Budget Menu Confirm Budget        : " + actReviseAccountR2 +      "  value expected  " + expReviseAccountR2);
   	   System.out.println("Budget Menu Budget Budget         : " + actReviseDepartmentR2 +   "  value expected  " + expReviseDeparmentR2);
   	   System.out.println("Budget Menu Approval Budget       : " + actReviseItemR2 +         "  value expected  " + expReviseItemR2);
   	   System.out.println("Budget Menu Budget Report Budget  : " + actReviseQtyR2 +          "  value expected  " + expReviseQtyR2);
   	   System.out.println("Budget Menu Define Budget         : " + actReviseValueR2 +        "  value expected  " + expReviseValueR2);
   	   System.out.println("Budget Menu Confirm Budget        : " + actReviseStatusR2 +       "  value expected  " + expReviseStatusR2);
   	   
   		if(actRevisePlanName.equalsIgnoreCase(expRevisePlanName) && actReviseAccount.equalsIgnoreCase(expReviseAccount)
   				&& actReviseDepartment.equalsIgnoreCase(expReviseDeparment) && actReviseItem.equalsIgnoreCase(expReviseItem)
   				&& actReviseQty.equalsIgnoreCase(expReviseQty) && actReviseValue.equalsIgnoreCase(expReviseValue)
   				&& actReviseStatus.equalsIgnoreCase(expReviseStatus)
   				
   				&& actRevisePlanNameR2.equalsIgnoreCase(expRevisePlanNameR2) && actReviseAccountR2.equalsIgnoreCase(expReviseAccountR2)
   				&& actReviseDepartmentR2.equalsIgnoreCase(expReviseDeparmentR2) && actReviseItemR2.equalsIgnoreCase(expReviseItemR2)
   				&& actReviseQtyR2.equalsIgnoreCase(expReviseQtyR2) && actReviseValueR2.equalsIgnoreCase(expReviseValueR2)
   				&& actReviseStatusR2.equalsIgnoreCase(expReviseStatusR2))				
   		{
   			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resPass);
   			return true;
   		}
   		else
   		{
   			excelReader.setCellData(xlfile, xlSheetName, 12, 9, resFail);
   			return false;
   		}
   	}
   	/*catch (Exception e) 
    {
   	   String exception =e.getMessage().substring(0, 50);
          excelReader.setExceptionInExcel(xlfile, xlSheetName, 12, 10, e.getMessage());
          System.err.println("CATCH BLOCK : EXECPTION MESSAGE IS DISPLAY : "+ e.getMessage());
          return false;
    }
   }*/
    
    
    
    @FindBy(xpath="//*[@id='myNavbar']/ul/li[4]/span")
	private static WebElement  reviseCloseBtn; 
    
    public boolean checkCloseButtonInReviseBudget() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
			excelReader=new ExcelReader(POJOUtility.getExcelPath());
			xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reviseCloseBtn));
			reviseCloseBtn.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dashboard));
			
			boolean actdashboard=dashboard.isDisplayed();
			boolean expdashboard=true;
				
			System.out.println("dashboard  :"+actdashboard +"  Value Expected  : "+expdashboard);
			
			Thread.sleep(3000);
			  
			   getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(userNameDisplay));
			   userNameDisplay.click();

			   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			   logoutOption.click();
			
			if(actdashboard==expdashboard)	
			{
				System.out.println("Test Pass : Cancel Option Working ");
				excelReader.setCellData(xlfile, xlSheetName, 53, 9, resPass);
				return true;
			}
			else
			{
				System.out.println("Test Fail : Cancel Option Not Working");
				excelReader.setCellData(xlfile, xlSheetName, 53, 9, resFail);
				return false;
			}
		}
    
    
  public SmokeBudgetPage(WebDriver driver)
  {
	  PageFactory.initElements(driver, this);	
  }
			
 }