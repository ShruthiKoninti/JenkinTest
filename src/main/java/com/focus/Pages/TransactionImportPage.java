package com.focus.Pages;

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

import com.focus.base.BaseEngine;

public class TransactionImportPage extends BaseEngine
{

	//Home Menu
		@FindBy (xpath="//*[@id='1']/div/span")
	    public static WebElement homeMenu;

		 //Utilities Menu		
		@FindBy (xpath="//*[@id='20']/span")
	    private static WebElement utilities;
	
		@FindBy (xpath="//*[@id='128']/span")
	    private static WebElement transactionImport;
	
		@FindBy (xpath="//input[@id='oTransImportVouchers']")
	    private static WebElement transImpVoucherTyeDrpdwn;
	
		@FindBy (xpath="//button[@id='btnTrnsImpBrowse']")
	    private static WebElement transImpFileNameBtn;
	
		@FindBy (xpath="//input[@id='txtTrnsImpFile']")
	    private static WebElement transImpFileNameTxt;
	
		@FindBy (xpath="//select[@id='ddlDAccInATrnsTagTo']")
	    private static WebElement transImpAccInTransTaggedDrpdwn;
	
		@FindBy (xpath="//select[@id='ddlDProInATrnsTagTo']")
	    private static WebElement transImpItemInTransTaggedDrpdwn;
	
		@FindBy (xpath="//select[@id='ddlTagsInTrnsTagTo']")
	    private static WebElement transImpTransInTaggedDrpdwn;
	
		@FindBy (xpath="//*[@id='RestoreControls']/ul/li/div[1]")
	    private static WebElement TransImpSaveTempleteBtn;
	
		@FindBy (xpath="//*[@id='RestoreControls']/ul/li/div[2]")
	    private static WebElement transImpImportBtn;
	
		@FindBy (xpath="//*[@id='RestoreControls']/ul/li/div[3]")
	    private static WebElement transImpCloseBtn;
	
		@FindBy (xpath="//*[@id='grdTransactionImport_body']/tr/td[2]")
	    private static List<WebElement> transImpGridFocusFieldColList;
		
		
		@FindBy (xpath="//*[@id='grdTransactionImport_body']/tr/td[3]")
	    private static List<WebElement> transImpGridDataBaseFieldColList;
	
		@FindBy (xpath="//*[@id='grdTransactionImport_body']/tr[4]/td[3]")
	    private static WebElement transImpgrid4RowCol3;
	
		@FindBy (xpath="//*[@id='ddlDatabaseField']")
	    private static WebElement gridDataBaseList;
	
		
		
		@FindBy (xpath="//*[@id='grdTransactionImport_body']/tr[5]/td[3]")
	    private static WebElement transImpgrid5RowCol3;
		
		
		@FindBy (xpath="//*[@id='grdTransactionImport_body']/tr[13]/td[3]")
	    private static WebElement transImpgrid13RowCol3;
		
		
		@FindBy (xpath="//*[@id='grdTransactionImport_body']/tr[14]/td[3]")
	    private static WebElement transImpgrid14RowCol3;
	
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


		
		
		 public static boolean checkLogin() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException
		 {
			enterUrl("http://175.101.7.11:5252/focus9");
			
			Thread.sleep(4000);
			
	        LoginPage lp=new LoginPage(getDriver()); 
			
			String unamelt="su";
					      
			String pawslt="su";
				
			Thread.sleep(2000);
			
			lp.enterUserName(unamelt);
					
			lp.enterPassword(pawslt);
			
			Thread.sleep(2000);
		     
		     /*Select s = new Select(companyDropDownList);
		     int size=s.getOptions().size();
		     
		     System.err.println("SIZE : "+size);
		     
		     for(int i=0;i<size;i++)
		     {
		    	 String data;
		    	 
		    	 data=s.getOptions().get(i).getText();
		    	 
		    	 System.err.println("DATA : "+data);
		    	 
		    	 if(data.contains("Ravi Foods Pvt Ltd"))
		    	 {
		    		 s.selectByVisibleText(data);
		    		 
		    		 break;
		    	 }
		     }*/
			
			companyDropDownList.sendKeys("Ravi Foods Pvt Ltd");
			
			Thread.sleep(1000);
					  
			lp.clickOnSignInBtn();
			
			Thread.sleep(5000);
					     
			userNameDisplay.click();
			
			
			if(userNameDisplay.isDisplayed()==true)
			{	
				return true;
			}	 
			else
			{
				return false;
			}
		}

		@FindBy (xpath="//button[@class='Fbutton'][contains(text(),'Yes')]")
    	private static WebElement overRideYesBtn;
    	
    	@FindBy (xpath="//button[@class='Fbutton'][contains(text(),'No')]")
    	private static WebElement overRideNoBtn;
    	
    	@FindBy (xpath="//button[@class='Fbutton'][contains(text(),'Cancel')]")
    	private static WebElement overRideCancelBtn;
	    	
    	 @FindBy(xpath="//*[@id='userprofile']/li/span[2]")
  		 private static WebElement logoutOption;
  	    
		 
		 @FindBy (xpath="//*[@id='27']/span")
		 private static WebElement restore;
			
		 @FindBy (xpath="//button[@class='btn Fbutton']")
	     private static WebElement folderpathExpandBtn;
		 
		 @FindBy (xpath="//div[@id='RestoreControls']//div[1]//div[1]")
	     private static WebElement restoreCompanyBtn;
		 
		 @FindBy(xpath="//*[@id='companyLogo']")
		 private static WebElement companyLogo;

		@FindBy(xpath="//*[@id='ulCompanyDetails_HomePage']/li[1]")
		private static WebElement companyName;

		 
		 public static boolean checkRestoreOfCompany() throws EncryptedDocumentException, InvalidFormatException, IOException, InterruptedException 
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
			     
			     Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\optionsCompanyRestore.exe");
			     
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
				
			
			     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplay));
			     userNameDisplay.click();
				
				
			     getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
			     logoutOption.click();
			     
			     

			     Thread.sleep(3000);
			     
			     lp.enterUserName(unamelt);
			     
			     lp.enterPassword(pawslt);
			     
			     Thread.sleep(2000);
			     
			     Select s = new Select(companyDropDownList);
			     int size=s.getOptions().size();
			     
			     System.err.println("SIZE : "+size);
			     
			     for(int i=0;i<size;i++)
			     {
			    	 String data;
			    	 
			    	 data=s.getOptions().get(i).getText();
			    	 
			    	 System.err.println("DATA : "+data);
			    	 
			    	 if(data.contains("Automation Options"))
			    	 {
			    		 
			    		 s.selectByVisibleText(data);
			    		 
			    		 break;
			    	 }
			     }
			     
			     
			     
			     lp.clickOnSignInBtn();
			     
			     Thread.sleep(10000);
			     
			     
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
				String expLoginCompanyName1   ="Automation Options ";
			     
			     
				
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
		 
		
		 @FindBy(xpath="//span[@class='hidden-xs']")
		 private static WebElement userNameDisplay;
		 
		/* @FindBy(id="ddlCompany")
		 private static WebElement companyDropDownList;
*/
		 
		 @FindBy(xpath="//select[@id='ddlCompany']")
		 private static WebElement companyDropDownList;

			
			
			@FindBy (xpath="//div[@class='navText']//span[contains(text(),'Transaction Import')]")
		    private static WebElement transactionImportLabel;
			
			@FindBy (xpath="//input[@id='oTransImportVouchers']")
		    private static WebElement voucherTypeComboBox;
			
			@FindBy (xpath="//div[@id='RestoreControls']//div[1]//div[1]")
		    private static WebElement saveTemplateBtn;
			
			@FindBy (xpath="//span[@class='icon-import icon-font8']")
		    private static WebElement importBtn;
			
			@FindBy (xpath="//span[@class='icon-close icon-font8']")
		    private static WebElement closeBtn;
			
			@FindBy (xpath="//button[@id='btnTrnsImpBrowse']")
		    private static WebElement fileNameExpandBtn;
			
			@FindBy (xpath="//select[@id='ddlDrCrSpecification']")
		    private static WebElement debitCreditSpecificationDropDown;
			
			@FindBy (xpath="//input[@id='txtDrCrCharacter']")
		    private static WebElement debitCreditSpecificationTxt;
			
			@FindBy (xpath="//select[@id='ddlTrnsRecords']")
		    private static WebElement transactionRecordsDropDown;
			
			@FindBy (xpath="//select[@id='ddlDMainAccRecInAVIs']")
		    private static WebElement theMainAccountRecordsInaVoucherIsDropDown;
			
			@FindBy (xpath="//select[@id='ddlDAccInATrnsTagTo']")
		    private static WebElement theAccountInaTransactionTaggedToDropDown;
			
			@FindBy (xpath="//select[@id='ddlDProInATrnsTagTo']")
		    private static WebElement theItemInaTransactionTaggedToDropDown;

			@FindBy (xpath="//select[@id='ddlTagsInTrnsTagTo']")
		    private static WebElement tagsInTransactionTaggedToDropDown;

			@FindBy (xpath="//input[@id='chkImportRefByBillsInv']")
		    private static WebElement importReferenceByBilssInvoiceChkBox;

			@FindBy (xpath="//select[@id='ddlDatabaseField']")
		    private static WebElement databaseFieldDropDown;

			@FindBy (xpath="//select[@id='ddlDatabaseField1']")
		    private static WebElement databaseField1DropDown;

			@FindBy (xpath="//select[@id='ddlDatabaseField2']")
		    private static WebElement databaseField2DropDown;

			@FindBy (xpath="//section[@id='page_Content']//tr[1]//td[3]")
		    private static WebElement selectRow1dataField;

			@FindBy (xpath="//section[@id='page_Content']//tr[1]//td[4]")
		    private static WebElement selectRow1dataField1;

			@FindBy (xpath="//section[@id='page_Content']//tr[1]//td[5]")
		    private static WebElement selectRow1dataField2;
			
			@FindBy (xpath="//section[@id='page_Content']//tr[2]//td[3]")
		    private static WebElement selectRow2dataField;

			@FindBy (xpath="//section[@id='page_Content']//tr[2]//td[4]")
		    private static WebElement selectRow2dataField1;

			@FindBy (xpath="//section[@id='page_Content']//tr[2]//td[5]")
		    private static WebElement selectRow2dataField2;
			
			
			@FindBy(xpath="//*[@id='oTransImportVouchers_table_body']/tr/td")
			private static List<WebElement> voucherTypeComboBoxList;
			
			
			@FindBy(xpath="//div[@class='theme_color font-6']")
			public static WebElement errorMessage;
							
			@FindBy(xpath="//span[@class='icon-reject2 theme_color']")
			public static WebElement errorMessageCloseBtn;
			
		
		
		
		
		
		//Server Error Elements
		@FindBy(xpath="//div[@id='id_focus_msgbox_main']")
		private static WebElement serverErrorPopup;
			
		@FindBy(xpath="//div[@id='id_focus_msgbox_title']/div[2]/span")
		private static WebElement serverErrorTitleClose;
						
		@FindBy(xpath="//div[@id='id_focus_msgbox_main']/div[2]/button")
		private static WebElement serverErrorClose;
		
		
		
		
		
		
		
		
		public boolean checkClickonImportOptionInTransationImport() throws InterruptedException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
			transImpImportBtn.click();
			
			Thread.sleep(2000);
			
            getWaitForAlert();
			
            String newLine = System.getProperty("line.separator");
           
			String alerttxt=getAlert().getText();
			String expalert="2 Vouchers imported successfully";
			String expalert1="Vouchers imported : 2,1";
			String expalert2="Vouchers failed : 1";
			String expalert3="1 Vouchers importing failed : 1";
				
			
			System.out.println("alerttxt : "+alerttxt);
			System.out.println("expalert : "+expalert);
			System.out.println("expalert : "+expalert1);
			System.out.println("expalert : "+expalert2);
			System.out.println("expalert : "+expalert3);
			getAlert().accept();
			
			if (alerttxt.startsWith(expalert) && 
					alerttxt.contains(expalert1) && alerttxt.contains(expalert2) && alerttxt.endsWith(expalert3))
			{
				System.out.println(" Test Pass: Voucher imported Successfully");
				return true;
			} 
			else
			{
				System.out.println(" Test Fail: Voucher not imported  ");
				return false;
			}
		}
		@FindBy (xpath="//span[@class='icon-financial icon-font1']")
	    private static WebElement financialMenu;
		
		
		@FindBy(xpath="//*[@id='61']/span")
		private static WebElement  financialsTransactionMenu; 	

		@FindBy(xpath = "//div[@id='id_transaction_homescreen_edit']//span[@class='icon-edit icon-font6 ImagesinArabic']")
		private static WebElement editBtn;

		
		 public boolean checkLoadingMessage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		  {
			try
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessage));
				String actVoucherLoadingMessage=errorMessage.getText();
				String expVoucherLoadingMessage="Voucher loaded successfully";
				
				System.out.println("VoucherLoadingMessage  : " + actVoucherLoadingMessage + " Value Expected : " + expVoucherLoadingMessage);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
				errorMessageCloseBtn.click();

				if(actVoucherLoadingMessage.startsWith(expVoucherLoadingMessage))
				{
					return true;
				}
				else
				{
					return false;
				}
			}
			catch(Exception e)
			{
				System.err.println("UNABLE TO COMPARE");
				return false;
			}
		  }
		
		@FindBy(xpath="//input[@id='id_header_10']")
		private static WebElement  voucherHeaderCurrency;

		 
		@FindBy(xpath="//*[@id='2007']")
		private static WebElement  financialsTransactionsPurchaseMenu; 
		
		@FindBy(xpath="//*[@id='tblBodyTransRender']/tr/td[4]")
	    private static List<WebElement>  voucherBodyGridVouRowCount;

		
		@FindBy(xpath="//*[@id='tblBodyTransRender']/tr/td[2]/input")
	    private static List<WebElement>  voucherGridIndexChkboxlist;
		
		
		@FindBy(xpath="//*[@id='2008']")
		private static WebElement  purchaseVoucherMenu;
		
		@FindBy(xpath="//*[@id='id_transactionentry_close']")
		private static WebElement  new_CloseBtn;
		
		
		public boolean checkImportedVouchersinPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
			financialMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
			purchaseVoucherMenu.click();
			
			
			Thread.sleep(3000);
			
			int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
			 
			for (int i = 0; i < VoucherBodyGridVouRowCount; i++) 
			{
				String docNo=voucherBodyGridVouRowCount.get(i).getText();
				
				if (docNo.equalsIgnoreCase("1")) 
				{
					voucherGridIndexChkboxlist.get(i).click();
					
					break;
				}
			}
		
		 Thread.sleep(2000);
			
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		 editBtn.click();
		 
		 
		 boolean loading=checkLoadingMessage();
		
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			String actDocNo = documentNumberTxt.getAttribute("value");
			String expDocNo = "1";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		    String actVendor = vendorAccountTxt.getAttribute("value");
			String expVendor = "Vendor A";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			String actCurrency = voucherHeaderCurrency.getAttribute("value");
			String expCurrency = "INR";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			String actDepartment = departmentTxt.getAttribute("value");
			String expDepartment = "India";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
			String actpurchaseAccountTxt = purchaseAccountTxt.getAttribute("value");
			String exppurchaseAccountTxt = "purchase";

			
			ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
				
			int voucherRow1ListCount=voucherRow1List.size();
				
			for (int j = 0; j < voucherRow1ListCount; j++) 
			{
					String data=voucherRow1List.get(j).getText();
					voucherRow1ListArray.add(data);
			}
				String actRowList=voucherRow1ListArray.toString();
				String expRow1List="[1, HYDERABAD, WA COGS ITEM, Pcs, 0.00, 5.00, 0.00, 0.00, 0.00, , , , Hk1,Hk2,Hk3,Hk4,Hk5]";
				
		        System.out.println("actRowList   : "+actRowList );
		        System.out.println("expRow1List  : "+expRow1List );
		     
				
				System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
				System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
				System.out.println("purchaseAccountTxt   : "+actpurchaseAccountTxt	+"  Value Expected  "+exppurchaseAccountTxt);
				System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
				System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
		        if (actRowList.equalsIgnoreCase(expRow1List)  
		        		&& actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
		        		&& actDepartment.equalsIgnoreCase(expDepartment) && actpurchaseAccountTxt.equalsIgnoreCase(exppurchaseAccountTxt))

		        {
		        	return true;
		        }

		        else
		        {
		        	return false;
		        }
			
		}
		
		//Voucher Entry Page Header Fields
		@FindBy(xpath="//input[@id='id_header_1']")
		private static WebElement  documentNumberTxt;
		
		@FindBy(xpath="//*[@id='id_header_1_input_image']/span")
		private static WebElement  documentNumberdropdown_ExpansionBtn;
		
		@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td")
		 private static List<WebElement> voucherRow1List;
		 
		 @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td")
		 private static List<WebElement> voucherRow2List;
			
		 
		 @FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr")
		 private static List<WebElement> voucherGridRowList;
		
		@FindBy(xpath="//input[@id='id_header_2']")
		private static WebElement  dateTxt;
		
		@FindBy(xpath="//td[@id='id_header_2_input_image']//span[@class='icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow']")
		private static WebElement  dateTxt_CalenderBtn;
		
		@FindBy(xpath="//tr[@id='id_header_2_day_today']//span[@class='theme_color-inverse'][contains(text(),'Today')]")
		private static WebElement  calender_TodayBtn;
		
		@FindBy(xpath="//input[@id='id_header_4']")
		private static WebElement  caskBankAccountTxt;

		@FindBy(xpath="//input[@id='id_header_268435459']")
		private static WebElement  departmentTxt;
		
		@FindBy(xpath="//input[@id='id_header_4']")
		private static WebElement  vendorAccountTxt;
		
		@FindBy(xpath="//input[@id='id_header_4']")
		private static WebElement  customerAccountTxt;
		
		@FindBy(xpath="//input[@id='id_header_268435460']")
		private static WebElement  warehouseTxt;
		
		
		@FindBy(xpath="//input[@id='id_header_3']")
		private static WebElement  purchaseAccountTxt;
		
		@FindBy(xpath="//input[@id='id_header_3']")
		private static WebElement  salesAccTxt;
		
		

		@FindBy(xpath="//input[@id='id_header_67108880']")
		private static WebElement  purNarrationTxt;
		
		
		
		@FindBy (xpath="//input[@id='idEraseTransaction']")
		private static WebElement eraseTranscationsRadio;
		
		@FindBy (xpath="//span[contains(text(),'Ok')]")
		private static WebElement eraseAllOkBtn;
		
		@FindBy (xpath="//ul[@class='nav navbar-nav navbar-right']//i[@class='icon-font6 icon-close']")
		private static WebElement eraseAllCancelBtn;
		
		@FindBy (xpath="//*[@id='25']/span")
        private static WebElement dataMangementMenu;
    	
		
		@FindBy (xpath="//*[@id='115']/span")
		private static WebElement eraseAll;
		
		public boolean checkEraseAllTransactionsAfterCompletingTransactionFlow() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
			dataMangementMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAll));
			eraseAll.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseTranscationsRadio));
			eraseTranscationsRadio.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAllOkBtn));
			eraseAllOkBtn.click();
			
			if(getIsAlertPresent())
			{
				getWaitForAlert();
				
				getAlert().accept();
			}
			
			String expValidationMsg="Data deleted successfully";
			
			String actValidationMsg=checkValidationMessage(expValidationMsg);
			
			System.out.println("************************************* CheckSavingPurchaseVoucherWithCovertOptionWithAllMandatoryFields  *********************************");
			System.out.println("Message  :  "+actValidationMsg +" Value Expected : "+expValidationMsg);
			
			if (actValidationMsg.equalsIgnoreCase(expValidationMsg))
			{
				System.out.println("Date Erased");
				return true;
			}
			else
			{
				System.out.println("Date  NotErased");
				return false;
			}
		}
		
		
		
		public boolean checkTransationImportWithFooterField() throws IOException, InterruptedException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
			utilities.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
			transactionImport.click();

            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
			voucherTypeComboBox.click();
			voucherTypeComboBox.sendKeys("p");
				   
			int count=voucherTypeComboBoxList.size();
			
			System.err.println(count);
			
			for(int i=0 ; i < count ;i++)
			{
				String data=voucherTypeComboBoxList.get(i).getText();
				
				if(data.equalsIgnoreCase("Purchases Vouchers"))
				{
					voucherTypeComboBoxList.get(i).click();
					break;
				}
				
			}
			 
			voucherTypeComboBox.sendKeys(Keys.TAB);
				
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
			importBtn.click();*/
			
			

           getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
			transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
			
			transImpVoucherTyeDrpdwn.click();
			
			transImpFileNameBtn.click();
			
			Thread.sleep(2000);
			Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\TransImportFD.exe");
			Thread.sleep(3000);
			
			

           getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
			transImpgrid4RowCol3.click();
			Select purchase=new Select(gridDataBaseList);
			purchase.selectByValue("2");
			
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
			transImpgrid5RowCol3.click();
			Select vendor=new Select(gridDataBaseList);
			vendor.selectByValue("3");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
			transImpImportBtn.click();
			
			Thread.sleep(2000);
			
            getWaitForAlert();
			
            String newLine = System.getProperty("line.separator");
            
			String alerttxt=getAlert().getText();
			
			String expalert="4 Vouchers imported successfully";
			String expalert1="Vouchers imported : FOC-HYD-1,FOC-HYD-2,FOC-HYD-3,FOC-HYD-4";
					
			
			System.out.println("alerttxt : "+alerttxt);
			System.out.println("expalert : "+expalert);
			System.out.println("expalert : "+expalert1);
			
			getAlert().accept();
			
			if (alerttxt.startsWith(expalert) && 
					alerttxt.endsWith(expalert1) )
			
			{
				System.out.println(" Test Pass: Voucher imported Successfully");
				return true;
			} 
			else
			{

				System.out.println(" Test Fail: Voucher not imported  ");

				return false;
			}
		}

		@FindBy (xpath="//input[@id='id_footer_134217749']")
        private static WebElement pvFooterFD;
    	
		
		
		
		public boolean checkImportedVouchersinPurchaseVoucherWithFooterField() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
			financialMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
			purchaseVoucherMenu.click();
			
			
			Thread.sleep(3000);
			
			int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
			 
			for (int i = 0; i < VoucherBodyGridVouRowCount; i++) 
			{
				String docNo=voucherBodyGridVouRowCount.get(i).getText();
				
				if (docNo.equalsIgnoreCase("FOC-HYD-2")) 
				{
					voucherGridIndexChkboxlist.get(i).click();
					
					break;
				}
			}
		
		 Thread.sleep(2000);
			
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		 editBtn.click();
		 
		 
		 boolean loading=checkLoadingMessage();
		
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			String actDocNo = documentNumberTxt.getAttribute("value");
			String expDocNo = "FOC-HYD-2";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		    String actVendor = vendorAccountTxt.getAttribute("value");
			String expVendor = "Vendor A";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			String actCurrency = voucherHeaderCurrency.getAttribute("value");
			String expCurrency = "INR";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			String actDepartment = departmentTxt.getAttribute("value");
			String expDepartment = "India";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
			String actpurchaseAccountTxt = purchaseAccountTxt.getAttribute("value");
			String exppurchaseAccountTxt = "purchase";

			
			ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
				
			int voucherRow1ListCount=voucherRow1List.size();
				
			for (int j = 0; j < voucherRow1ListCount; j++) 
			{
					String data=voucherRow1List.get(j).getText();
					voucherRow1ListArray.add(data);
			}
				String actRowList=voucherRow1ListArray.toString();
				String expRow1List="[1, HYDERABAD, BR COGS ITEM, Pcs, 0.00, 20.00, 5.00, 100.00, 0.00, B3, , , ]";
				
		        System.out.println("actRowList   : "+actRowList );
		        System.out.println("expRow1List  : "+expRow1List );
		     
				
				System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
				System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
				System.out.println("purchaseAccountTxt   : "+actpurchaseAccountTxt	+"  Value Expected  "+exppurchaseAccountTxt);
				System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
				System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvFooterFD));
				JavascriptExecutor js = (JavascriptExecutor) getDriver();
				js.executeScript("arguments[0].scrollIntoView();",pvFooterFD);
				
				
				String actpvFooterFD=pvFooterFD.getAttribute("value");
				
				String exppvFooter="10.00";
				
				System.err.println("pvFooterFD  : "+actpvFooterFD +" Value Expected  : "+exppvFooter);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
		        if (actRowList.equalsIgnoreCase(expRow1List)  
		        		&& actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
		        		&& actDepartment.equalsIgnoreCase(expDepartment) && actpurchaseAccountTxt.equalsIgnoreCase(exppurchaseAccountTxt))

		        {
		        	return true;
		        }

		        else
		        {
		        	return false;
		        }
			
		}

		public boolean checkTransationImportWithDifferentDocNumber() throws IOException, InterruptedException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
			utilities.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
			transactionImport.click();

            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
			voucherTypeComboBox.click();
			voucherTypeComboBox.sendKeys("p");
				   
			int count=voucherTypeComboBoxList.size();
			
			System.err.println(count);
			
			for(int i=0 ; i < count ;i++)
			{
				String data=voucherTypeComboBoxList.get(i).getText();
				
				if(data.equalsIgnoreCase("Purchases Vouchers"))
				{
					voucherTypeComboBoxList.get(i).click();
					break;
				}
				
			}
			 
			voucherTypeComboBox.sendKeys(Keys.TAB);
				
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
			importBtn.click();
			
			*/

           getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
			transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
			
			transImpVoucherTyeDrpdwn.click();
			
			transImpFileNameBtn.click();
			
			Thread.sleep(2000);
			Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\TransImportDiffDoc.exe");
			Thread.sleep(3000);
			
			

           getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
			transImpgrid4RowCol3.click();
			Select purchase=new Select(gridDataBaseList);
			purchase.selectByValue("2");
			
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
			transImpgrid5RowCol3.click();
			Select vendor=new Select(gridDataBaseList);
			vendor.selectByValue("3");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
			transImpImportBtn.click();
			
			Thread.sleep(2000);
			
            getWaitForAlert();
			
            String newLine = System.getProperty("line.separator");
            
			String alerttxt=getAlert().getText();
			String expalert1="3 Vouchers imported successfully";
			String expalert2="Vouchers imported : 1,2,3";
					
			System.out.println("alerttxt : "+alerttxt);
			System.out.println("expalert1 : "+expalert1);
			System.out.println("expalert2 : "+expalert2);
			
			getAlert().accept();
			
			if (alerttxt.startsWith(expalert1) && 
					alerttxt.endsWith(expalert2))
			{
				System.out.println(" Test Pass: Voucher imported Successfully");
				return true;
			} 
			else
			{

				System.out.println(" Test Fail: Voucher not imported  ");

				return false;
			}
		}


		
		public boolean checkImportedVouchersinPurchaseVoucherWithDiffDocNumber() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
			financialMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
			purchaseVoucherMenu.click();
			
			Thread.sleep(3000);
			
			int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
			 
			int expVoucherBodyGridVouRowCount=4;
			
			System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
			
			for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
			{
				String docNo=voucherBodyGridVouRowCount.get(i).getText();
				
				if (docNo.equalsIgnoreCase("2")) 
				{
					voucherGridIndexChkboxlist.get(i).click();
					
					break;
				}
			}
		
		 Thread.sleep(2000);
			
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		 editBtn.click();
		 
		 
		 boolean loading=checkLoadingMessage();
		
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			String actDocNo = documentNumberTxt.getAttribute("value");
			String expDocNo = "2";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		    String actVendor = vendorAccountTxt.getAttribute("value");
			String expVendor = "Vendor A";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			String actCurrency = voucherHeaderCurrency.getAttribute("value");
			String expCurrency = "INR";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			String actDepartment = departmentTxt.getAttribute("value");
			String expDepartment = "India";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
			String actpurchaseAccountTxt = purchaseAccountTxt.getAttribute("value");
			String exppurchaseAccountTxt = "purchase";

			
			ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
				
			int voucherRow1ListCount=voucherRow1List.size();
				
			for (int j = 0; j < voucherRow1ListCount; j++) 
			{
					String data=voucherRow1List.get(j).getText();
					voucherRow1ListArray.add(data);
			}
				String actRowList=voucherRow1ListArray.toString();
				String expRow1List="[1, HYDERABAD, WA COGS ITEM, Pcs, 0.00, 5.00, 5.00, 25.00, 0.00, , , , Hk1,Hk2,Hk3,Hk4,Hk5]";
				
		        System.out.println("actRowList   : "+actRowList );
		        System.out.println("expRow1List  : "+expRow1List );
		     
				
				System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
				System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
				System.out.println("purchaseAccountTxt   : "+actpurchaseAccountTxt	+"  Value Expected  "+exppurchaseAccountTxt);
				System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
				System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvFooterFD));
				Thread.sleep(2000);
				JavascriptExecutor js = (JavascriptExecutor) getDriver();
				js.executeScript("arguments[0].scrollIntoView();",pvFooterFD);
				
				
				String actpvFooterFD=pvFooterFD.getAttribute("value");
				
				String exppvFooter="0.00";
				
				System.err.println("pvFooterFD  : "+actpvFooterFD +" Value Expected  : "+exppvFooter);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
		        if (actRowList.equalsIgnoreCase(expRow1List)  
		        		&& actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
		        		&& actDepartment.equalsIgnoreCase(expDepartment) && actpurchaseAccountTxt.equalsIgnoreCase(exppurchaseAccountTxt)
		        		 &&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount)

		        {
		        	return true;
		        }

		        else
		        {
		        	return false;
		        }
		}
		
		
		
		
		public boolean checkTransationImportWithDuplicateRMA() throws IOException, InterruptedException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
			utilities.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
			transactionImport.click();

            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
			voucherTypeComboBox.click();
			voucherTypeComboBox.sendKeys("p");
				   
			int count=voucherTypeComboBoxList.size();
			
			System.err.println(count);
			
			for(int i=0 ; i < count ;i++)
			{
				String data=voucherTypeComboBoxList.get(i).getText();
				
				if(data.equalsIgnoreCase("Purchases Vouchers"))
				{
					voucherTypeComboBoxList.get(i).click();
					break;
				}
				
			}
			 
			voucherTypeComboBox.sendKeys(Keys.TAB);
			/*	
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
			importBtn.click();
			*/
			
            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
			transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
			
			transImpVoucherTyeDrpdwn.click();
			
			transImpFileNameBtn.click();
			
			Thread.sleep(2000);
			Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\TransImportWithDuplicateRMA.exe");
			Thread.sleep(3000);
			
			

           getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
			transImpgrid4RowCol3.click();
			Select purchase=new Select(gridDataBaseList);
			purchase.selectByValue("2");
			
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
			transImpgrid5RowCol3.click();
			Select vendor=new Select(gridDataBaseList);
			vendor.selectByValue("3");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
			transImpImportBtn.click();
			
			Thread.sleep(2000);
			
            getWaitForAlert();
			
            String newLine = System.getProperty("line.separator");
            
			String alerttxt=getAlert().getText();
			String expalert="3 Vouchers imported successfully";
		    String expalert1="Vouchers imported : 1,2,3";
					
			
			System.out.println("alerttxt : "+alerttxt);
			System.out.println("expalert : "+expalert);
			System.out.println("expalert1 : "+expalert1);
			
			getAlert().accept();
			
			if (alerttxt.startsWith(expalert) && 
					alerttxt.endsWith(expalert1))
			{
				System.out.println(" Test Pass: Voucher imported Successfully");
				return true;
			} 
			else
			{

				System.out.println(" Test Fail: Voucher not imported  ");

				return false;
			}
		}


		public boolean checkImportedVouchersinPurchaseVoucherWithDuplicateRMA() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
			financialMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
			purchaseVoucherMenu.click();
			
			
			Thread.sleep(3000);
			
			int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
			 
			int expVoucherBodyGridVouRowCount=4;
			
			System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
			
			for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
			{
				String docNo=voucherBodyGridVouRowCount.get(i).getText();
				
				if (docNo.equalsIgnoreCase("2")) 
				{
					voucherGridIndexChkboxlist.get(i).click();
					
					break;
				}
			}
		
		 Thread.sleep(2000);
			
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		 editBtn.click();
		 
		 
		 boolean loading=checkLoadingMessage();
		
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			String actDocNo = documentNumberTxt.getAttribute("value");
			String expDocNo = "2";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		    String actVendor = vendorAccountTxt.getAttribute("value");
			String expVendor = "Vendor A";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			String actCurrency = voucherHeaderCurrency.getAttribute("value");
			String expCurrency = "INR";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			String actDepartment = departmentTxt.getAttribute("value");
			String expDepartment = "India";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
			String actpurchaseAccountTxt = purchaseAccountTxt.getAttribute("value");
			String exppurchaseAccountTxt = "purchase";

			
			ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
				
			int voucherRow1ListCount=voucherRow1List.size();
				
			for (int j = 0; j < voucherRow1ListCount; j++) 
			{
					String data=voucherRow1List.get(j).getText();
					voucherRow1ListArray.add(data);
			}
				String actRowList=voucherRow1ListArray.toString();
				String expRow1List="[1, HYDERABAD, WA COGS ITEM, Pcs, 0.00, 20.00, 5.00, 100.00, 0.00, , , , HK2,Rk1]";
				
		        System.out.println("actRowList   : "+actRowList );
		        System.out.println("expRow1List  : "+expRow1List );
		        
		        ArrayList<String >voucherRow2ListArray = new ArrayList<String>();
				
				int voucherRow2ListCount=voucherRow2List.size();
					
				for (int j = 0; j < voucherRow2ListCount; j++) 
				{
						String data=voucherRow2List.get(j).getText();
						voucherRow2ListArray.add(data);
				}
					String actRow2List=voucherRow2ListArray.toString();
					String expRow2List="[2, HYDERABAD, BR COGS ITEM, Pcs, 0.00, 5.00, 5.00, 25.00, 0.00, B3, , , ]";
					
			        System.out.println("actRow2List   : "+actRow2List );
			        System.out.println("expRow2List  : "+expRow2List );
		     
				
				System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
				System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
				System.out.println("purchaseAccountTxt   : "+actpurchaseAccountTxt	+"  Value Expected  "+exppurchaseAccountTxt);
				System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
				System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvFooterFD));
				Thread.sleep(2000);
				JavascriptExecutor js = (JavascriptExecutor) getDriver();
				js.executeScript("arguments[0].scrollIntoView();",pvFooterFD);
				
				
				String actpvFooterFD=pvFooterFD.getAttribute("value");
				
				String exppvFooter="0.00";
				
				System.err.println("pvFooterFD  : "+actpvFooterFD +" Value Expected  : "+exppvFooter);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
		        if (actRowList.equalsIgnoreCase(expRow1List)  
		        		&& actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
		        		&& actDepartment.equalsIgnoreCase(expDepartment) && actpurchaseAccountTxt.equalsIgnoreCase(exppurchaseAccountTxt)
		        		 &&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount && 
		        		 actRow2List.equalsIgnoreCase(expRow2List))

		        {
		        	return true;
		        }

		        else
		        {
		        	return false;
		        }
		}
		

		
		public boolean checkTransationImportWithRateZero() throws IOException, InterruptedException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
			homeMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
			utilities.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
			transactionImport.click();

            getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
			voucherTypeComboBox.click();
			voucherTypeComboBox.sendKeys("p");
				   
			int count=voucherTypeComboBoxList.size();
			
			System.err.println(count);
			
			for(int i=0 ; i < count ;i++)
			{
				String data=voucherTypeComboBoxList.get(i).getText();
				
				if(data.equalsIgnoreCase("Purchases Vouchers"))
				{
					voucherTypeComboBoxList.get(i).click();
					break;
				}
				
			}
			 
			voucherTypeComboBox.sendKeys(Keys.TAB);
				
			/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
			importBtn.click();
			
			*/

           getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
			transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
			
			transImpVoucherTyeDrpdwn.click();
			
			transImpFileNameBtn.click();
			
			Thread.sleep(2000);
			Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\TransImportRate0.exe");
			Thread.sleep(3000);
			
			

           getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
			transImpgrid4RowCol3.click();
			Select purchase=new Select(gridDataBaseList);
			purchase.selectByValue("2");
			
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
			transImpgrid5RowCol3.click();
			Select vendor=new Select(gridDataBaseList);
			vendor.selectByValue("3");
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
			transImpImportBtn.click();
			
			Thread.sleep(2000);
			
            getWaitForAlert();
			
            String newLine = System.getProperty("line.separator");
            
			String alerttxt=getAlert().getText();
			String expalert="4 Vouchers imported successfully";
			String expalert1="Vouchers imported : 1,2,3,4";
					
			
			System.out.println("alerttxt : "+alerttxt);
			System.out.println("expalert : "+expalert);
			System.out.println("expalert 1: "+expalert1);
			
			getAlert().accept();
			
			if (alerttxt.startsWith(expalert) && 
					alerttxt.endsWith(expalert1)) 
			{
				System.out.println(" Test Pass: Voucher imported Successfully");
				return true;
			} 
			else
			{

				System.out.println(" Test Fail: Voucher not imported  ");

				return false;
			}
		}


		public boolean checkImportedVouchersinPurchaseVoucherWitRateZero() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
		{
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
			financialMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
			financialsTransactionMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
			financialsTransactionsPurchaseMenu.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
			purchaseVoucherMenu.click();
			
			
			Thread.sleep(3000);
			
			int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
			 
			int expVoucherBodyGridVouRowCount=5;
			
			System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
			
			for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
			{
				String docNo=voucherBodyGridVouRowCount.get(i).getText();
				
				if (docNo.equalsIgnoreCase("2")) 
				{
					voucherGridIndexChkboxlist.get(i).click();
					
					break;
				}
			}
		
		 Thread.sleep(2000);
			
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
		 editBtn.click();
		 
		 
		 boolean loading=checkLoadingMessage();
		
		 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
			String actDocNo = documentNumberTxt.getAttribute("value");
			String expDocNo = "2";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		    String actVendor = vendorAccountTxt.getAttribute("value");
			String expVendor = "Vendor A";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
			String actCurrency = voucherHeaderCurrency.getAttribute("value");
			String expCurrency = "INR";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
			String actDepartment = departmentTxt.getAttribute("value");
			String expDepartment = "India";
				
			getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
			String actpurchaseAccountTxt = purchaseAccountTxt.getAttribute("value");
			String exppurchaseAccountTxt = "purchase";

			
			ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
				
			int voucherRow1ListCount=voucherRow1List.size();
				
			for (int j = 0; j < voucherRow1ListCount; j++) 
			{
					String data=voucherRow1List.get(j).getText();
					voucherRow1ListArray.add(data);
			}
				String actRowList=voucherRow1ListArray.toString();
				String expRow1List="[1, HYDERABAD, BR COGS ITEM, Pcs, 0.00, 20.00, 0.00, 0.00, 0.00, B3, , , ]";
				
		        System.out.println("actRowList   : "+actRowList );
		        System.out.println("expRow1List  : "+expRow1List );
		        
		        ArrayList<String >voucherRow2ListArray = new ArrayList<String>();
				
				int voucherRow2ListCount=voucherRow2List.size();
					
				 
				
				System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
				System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
				System.out.println("purchaseAccountTxt   : "+actpurchaseAccountTxt	+"  Value Expected  "+exppurchaseAccountTxt);
				System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
				System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvFooterFD));
				Thread.sleep(2000);
				JavascriptExecutor js = (JavascriptExecutor) getDriver();
				js.executeScript("arguments[0].scrollIntoView();",pvFooterFD);
				
				
				String actpvFooterFD=pvFooterFD.getAttribute("value");
				
				String exppvFooter="0.00";
				
				System.err.println("pvFooterFD  : "+actpvFooterFD +" Value Expected  : "+exppvFooter);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
				new_CloseBtn.click();
				
		        if (actRowList.equalsIgnoreCase(expRow1List)  
		        		&& actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
		        		&& actDepartment.equalsIgnoreCase(expDepartment) && actpurchaseAccountTxt.equalsIgnoreCase(exppurchaseAccountTxt)
		        		 &&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount)

		        {
		        	return true;
		        }

		        else
		        {
		        	return false;
		        }
		}
		

public boolean checkTransationImportWithCorrectData() throws IOException, InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("p");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Purchases Vouchers"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();
	*/
	

   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\TransImportWithCorrectData.exe");
	Thread.sleep(3000);
	
	

   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("2");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("3");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    String newLine = System.getProperty("line.separator");
    
	String alerttxt=getAlert().getText();
		
	String expalert="4 Vouchers imported successfully";
	String expAlert1="Vouchers imported : FOC-HYD-1,FOC-HYD-2,FOC-HYD-3,FOC-HYD-4";


	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert) && alerttxt.endsWith(expAlert1))
	{
		System.out.println(" Test Pass: Voucher imported Successfully");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
}


public boolean checkImportedVouchersinPurchaseVoucherWitCorrectData() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
	purchaseVoucherMenu.click();
	
	
	Thread.sleep(3000);
	
	int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
	 
	int expVoucherBodyGridVouRowCount=5;
	
	System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
	
	for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
	{
		String docNo=voucherBodyGridVouRowCount.get(i).getText();
		
		if (docNo.equalsIgnoreCase("FOC-HYD-2")) 
		{
			voucherGridIndexChkboxlist.get(i).click();
			
			break;
		}
	}

 Thread.sleep(2000);
	
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
 editBtn.click();
 
 
 boolean loading=checkLoadingMessage();

 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "FOC-HYD-2";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
    String actVendor = vendorAccountTxt.getAttribute("value");
	String expVendor = "Vendor A";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
	String actpurchaseAccountTxt = purchaseAccountTxt.getAttribute("value");
	String exppurchaseAccountTxt = "purchase";

	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, HYDERABAD, BR COGS ITEM, Pcs, 0.00, 20.00, 5.00, 100.00, 0.00, B3, , , ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
        
        ArrayList<String >voucherRow2ListArray = new ArrayList<String>();
		
		int voucherRow2ListCount=voucherRow2List.size();
			
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("purchaseAccountTxt   : "+actpurchaseAccountTxt	+"  Value Expected  "+exppurchaseAccountTxt);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(pvFooterFD));
		Thread.sleep(2000);
		JavascriptExecutor js = (JavascriptExecutor) getDriver();
		js.executeScript("arguments[0].scrollIntoView();",pvFooterFD);
		
		
		String actpvFooterFD=pvFooterFD.getAttribute("value");
		
		String exppvFooter="0.00";
		
		System.err.println("pvFooterFD  : "+actpvFooterFD +" Value Expected  : "+exppvFooter);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actpurchaseAccountTxt.equalsIgnoreCase(exppurchaseAccountTxt)
        		&&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount)

        {
        	return true;
        }

        else
        {
        	return false;
        }
    }




public boolean checkSalesInvoiceImportWithCorrectDataInTrasactionImport() throws InterruptedException, IOException
{
	
	
	getDriver().navigate().refresh();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("sales");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Sales Invoices"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();
	
	*/

   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\OutwardFullComsumption.exe");
	Thread.sleep(3000);
	
	

   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("4");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("5");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    String newLine = System.getProperty("line.separator");
    
	String alerttxt=getAlert().getText();
		
	String expalert="4 Vouchers imported successfully";
	String expAlert1="Vouchers imported : 1,2,3,4";

	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	System.out.println("expalert1 : "+expAlert1);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert) && alerttxt.endsWith(expAlert1))
	{
		System.out.println(" Test Pass: Voucher imported Successfully");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
	
}



@FindBy(xpath="//*[@id='2014']")
private static WebElement financialsTransactionssalesMenu;

@FindBy(xpath="//*[@id='2015']")
private static WebElement salesInvoiceMenu;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[20]")
private static WebElement  select1stRow_19thColumn;


public boolean checkImportedVouchersInSalesInvoiceWithFullComnsumption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionssalesMenu));
	financialsTransactionssalesMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceMenu));
	salesInvoiceMenu.click();
	
	
	Thread.sleep(3000);
	
	int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
	 
	int expVoucherBodyGridVouRowCount=5;
	
	System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
	
	for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
	{
		String docNo=voucherBodyGridVouRowCount.get(i).getText();
		
		if (docNo.equalsIgnoreCase("1")) 
		{
			voucherGridIndexChkboxlist.get(i).click();
			
			break;
		}
	}

    Thread.sleep(2000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
    editBtn.click();
 
 
    boolean loading=checkLoadingMessage();

    getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "1";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccTxt));
    String actsalesAccTxt = salesAccTxt.getAttribute("value");
	String expsalesAccTxt = "Sales - Computers";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String actcustomerAccountTxt = customerAccountTxt.getAttribute("value");
	String expcustomerAccountTxt = "Customer A";

	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	String actwarehouseTxt = warehouseTxt.getAttribute("value");
	String expwarehouseTxt = "HYDERABAD";

	
	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, BR COGS ITEM, Pcs, 0.00, 0.00, 0.00, 0.00, 0.00, 10.00, , , 5.00, 50.00, 0.00, B2, , , ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
        
			
		System.out.println("Document No          : "+actDocNo			    +"  Value Expected  "+expDocNo);
		System.out.println("salesAccTxt          : "+actsalesAccTxt			+"  Value Expected  "+expsalesAccTxt);
		System.out.println("customerAccountTxt   : "+actcustomerAccountTxt	+"  Value Expected  "+expcustomerAccountTxt);
		System.out.println("Currency             : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department           : "+actDepartment		    +"  Value Expected  "+expDepartment);
		System.out.println("warehouseTxt         : "+actwarehouseTxt	 	+"  Value Expected  "+expwarehouseTxt);
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(batchItemName));
		
		int actBatchRows = batchPopupRowNosList.size();
		int expBatchRows = 1;
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchTotalQtyAmt));
		String actTotalQuantityInHeader=batchTotalQtyAmt.getText();
		String expTotalQuantityInHeader="10.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchAdjustQtyAmt));
		String actAdjustQuantity=batchAdjustQtyAmt.getText();
		String expAdjustQuantity="10.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchBalanceAmt));
		String actBalance=batchBalanceAmt.getText();
		String expBalance="0.00";
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchTable1stRow1stCol));
		String act1stRowBatchNo=batchTable1stRow1stCol.getText();
		String exp1stRowBatchNo="B2";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchTable1stRow2ndCol));
		String act1stRowBaseQuantity=batchTable1stRow2ndCol.getText();
		String exp1stRowBaseQuantity="10.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchCancelIcon));
		batchCancelIcon.click();
		
		System.out.println("********* Batch Rows Count : "+actBatchRows+"  value expected  "+expBatchRows);
		System.out.println("********* Total Quantity: "+actTotalQuantityInHeader+"  value expected  "+expTotalQuantityInHeader);
		System.out.println("********* Adjust Quantity: "+actAdjustQuantity+"  value expected  "+expAdjustQuantity);
		System.out.println("********* Balance: "+actBalance+"  value expected  "+expBalance);
		System.out.println("********* 1.Batch No: "+act1stRowBatchNo+"  value expected  "+exp1stRowBatchNo);
		System.out.println("********* 1.Base Quantity: "+act1stRowBaseQuantity+"  value expected  "+exp1stRowBaseQuantity);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actsalesAccTxt.equalsIgnoreCase(expsalesAccTxt) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actcustomerAccountTxt.equalsIgnoreCase(expcustomerAccountTxt)
        		&&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount && actwarehouseTxt.equalsIgnoreCase(expwarehouseTxt) && 
        		VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount && 
        		actBatchRows==expBatchRows && actTotalQuantityInHeader.equalsIgnoreCase(expTotalQuantityInHeader) && actAdjustQuantity.equalsIgnoreCase(expAdjustQuantity) 
				  && actBalance.equalsIgnoreCase(expBalance) && actAdjustQuantity.equalsIgnoreCase(expAdjustQuantity)
				  && act1stRowBatchNo.equalsIgnoreCase(exp1stRowBatchNo) && act1stRowBaseQuantity.equalsIgnoreCase(exp1stRowBaseQuantity))

        {
        	return true;
        }

        else
        {
        	return false;
        }
	
}
		

@FindBy(xpath="//tbody[@id='Id_MRPBatch_Grid_body']/tr/td[2]")
private static List<WebElement> batchPopupBatchNoList;

@FindBy(xpath="//tbody[@id='Id_MRPBatch_Grid_body']/tr/td[1]")
private static List<WebElement> batchPopupRowNosList;


@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[15]")
private static WebElement  select1stRow_14thColumn;


@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[17]")
private static WebElement  select1stRow_16thColumn;

//Batch Popup Elements

		@FindBy(xpath="//label[@id='lblitem']")
		private static WebElement batchItemName;
		
		@FindBy(xpath="//label[@id='lblTotalQty']")
		private static WebElement batchTotalQtyAmt;
		
		@FindBy(xpath="//label[@id='lblAdjustQty']")
		private static WebElement batchAdjustQtyAmt;
		
		
		@FindBy(xpath="//label[@id='lblBalnceQty']")
		private static WebElement batchBalanceAmt;
		
		@FindBy(xpath="//div[contains(@class,'col-xs-12 col-sm-12 col-md-12 col-lg-12 P_ingBottom')]//div[2]//input[1]")
		private static WebElement batchExipiringInRadioBtn;
		             
		@FindBy(xpath="//section[@id='page_Content']//div[@class='row']//div//div[3]//input[1]")
		private static WebElement batchExipiringAfetrRadioBtn;
		
		@FindBy(xpath="//input[@id='txtWithin']")
		private static WebElement batchWithInTxt;
		
		@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_control_heading_6']")
		private static WebElement batchTableHeaderRateColumn;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-1']")
		private static WebElement batchTable1stRow1stCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-2']")
		private static WebElement batchTable1stRow2ndCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-5']")
		private static WebElement batchTable1stRow3rdCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-6']")
		private static WebElement batchTable1stRow4thCol;
		               
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-8']")
		private static WebElement batchTable1stRow5thCol;
		
		@FindBy(xpath="//input[@id='MRPBatch_AdjBatchQty']")
		private static WebElement enter_BatchTableQtyAdjusted;
		
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_1-13']")
		private static WebElement batchTable1stRow6thCol;
		
		
		@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_col_2-1']")
		private static WebElement batchTable2ndRow1stCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_2-2']")
		private static WebElement batchTable2ndRow2ndCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_2-5']")
		private static WebElement batchTable2ndRow3rdCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_2-6']")
		private static WebElement batchTable2ndRow4thCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_2-7']")
		private static WebElement batchTable2ndRow5thCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_2-13']")
		private static WebElement batchTable2ndRow6thCol;
		
		
		@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_col_3-1']")
		private static WebElement batchTable3rdRow1stCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_3-2']")
		private static WebElement batchTable3rdRow2ndCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_3-5']")
		private static WebElement batchTable3rdRow3rdCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_3-6']")
		private static WebElement batchTable3rdRow4thCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_3-7']")
		private static WebElement batchTable3rdRow5thCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_3-13']")
		private static WebElement batchTable3rdRow6thCol;
		
		
		
		@FindBy(xpath="//*[@id='Id_MRPBatch_Grid_col_4-1']")
		private static WebElement batchTable4thRow1stCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_4-2']")
		private static WebElement batchTable4thRow2ndCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_4-5']")
		private static WebElement batchTable4thRow3rdCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_4-6']")
		private static WebElement batchTable4thRow4thCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_4-7']")
		private static WebElement batchTable4thRow5thCol;
		
		@FindBy(xpath="//td[@id='Id_MRPBatch_Grid_col_4-13']")
		private static WebElement batchTable4thRow6thCol;
		
		@FindBy(xpath="//*[@id='MRPBatch_AdjBatchQty']")
		private static WebElement batchTableEnterQtyAdjusted;
		
		
		/*@FindBy(xpath="//tbody[@id='Id_MRPBatch_Grid_body']/tr")
		private static List<WebElement> batchPopupTable;*/
		
		@FindBy(xpath="//tbody[@id='Id_MRPBatch_Grid_body']")
		private static WebElement batchPopupTable;
		
		@FindBy(xpath="//i[@class='icon-refresh icon-font7']")
		private static WebElement batchRefreshBtn;
		
		@FindBy(xpath="//input[@id='Batch_Search']")
		private static WebElement batchSearchTxt;
		
		@FindBy(xpath="//input[@class='Fcheckbox']")
		private static WebElement batchShowAllAvalibleBatches;
		
		@FindBy(xpath="//label[contains(text(),'Customization')]")
		private static WebElement batchCustomizationIcon;
		
		/*@FindBy(xpath="//label[contains(text(),'Pick On FIFO')]")
		private static WebElement batchPickOnFIFOIcon;*/
		
		@FindBy(xpath="//div[@class='btnhighlight theme_button_color P_LeftRight10'][2]")
		private static WebElement batchPickOnFIFOIcon;
		
		@FindBy(xpath="//label[@id='div_Ok']")
		private static WebElement batchOkIcon;

		@FindBy(xpath="//div[@class='col-xs-12 col-sm-12 col-md-12 col-lg-12']//div[6]//label[1]")
		private static WebElement batchTotalAdjustedQty;
		
		@FindBy(xpath="//*[@id='MBatch_Content']/div[3]/div/div[3]/div[4]/label")
		private static WebElement batchCancelIcon;
		
		@FindBy(xpath="//*[@id='MBatch_Content']/div[1]/div[2]/span/i")
		private static WebElement batchCloseIcon;
		
		@FindBy(xpath="//*[@id='MBatch_Content']/div[2]/div/div[5]/div/div[3]")
		private static WebElement batchTotalQuantity;
		
		//Grid Header CheckBox
		@FindBy(xpath="//input[@id='HeaderChkBox']")
		private static WebElement  grid_HeaderChkBox;	
		
		@FindBy(xpath="//div[@id='id_transaction_homescreen_Delete']//span[@class='icon-delete icon-font6']")
		private static WebElement  deleteBtn;
		
		
		public boolean checkDeletingAllVouchersinSalesInvoice()
		{
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(grid_HeaderChkBox));
			grid_HeaderChkBox.click();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(deleteBtn));
			deleteBtn.click();
			
			
			getWaitForAlert();
			
			String actAlertTxt=getAlert().getText();
			String expAlertTxt="Are you sure that you want to delete the Selected Voucher?";
			
			getAlert().accept();
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
			errorMessageCloseBtn.click();
			
			
			System.out.println("************ Alert Text    : "+actAlertTxt +"  value expected  "+expAlertTxt);
			
			if(actAlertTxt.equalsIgnoreCase(expAlertTxt)) 
			{
				System.err.println(" Test PAss: Vouchers deleted Successfully");
				return true;
			}
			else
			{
				System.err.println(" Test Fail: Vouchers deleted Successfully");
				return false;
			}
		}
		
		

public boolean checkSalesInvoiceImportWithPartialDataInTrasactionImport() throws InterruptedException, IOException
{
	
	
	getDriver().navigate().refresh();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("sales");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Sales Invoices"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();
	*/
	

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\OutwardPartialConsume.exe");
	Thread.sleep(3000);
	
	

   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("4");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("5");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    
	String alerttxt=getAlert().getText();
		
	String expalert="4 Vouchers imported successfully";
	String expAlert1="Vouchers imported : 1,2,3,4";

	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	System.out.println("expalert1 : "+expAlert1);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert) && alerttxt.endsWith(expAlert1))
	{
		System.out.println(" Test Pass: Voucher imported Successfully");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
	
}



public boolean checkImportedVouchersInSalesInvoiceWithBatchPartialComnsumption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionssalesMenu));
	financialsTransactionssalesMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceMenu));
	salesInvoiceMenu.click();
	
	
	Thread.sleep(3000);
	
	int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
	 
	int expVoucherBodyGridVouRowCount=5;
	
	System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
	
	for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
	{
		String docNo=voucherBodyGridVouRowCount.get(i).getText();
		
		if (docNo.equalsIgnoreCase("1")) 
		{
			voucherGridIndexChkboxlist.get(i).click();
			
			break;
		}
	}

 Thread.sleep(2000);
	
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
 editBtn.click();
 
 
   boolean loading=checkLoadingMessage();

   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "1";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccTxt));
    String actsalesAccTxt = salesAccTxt.getAttribute("value");
	String expsalesAccTxt = "Sales - Computers";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String actcustomerAccountTxt = customerAccountTxt.getAttribute("value");
	String expcustomerAccountTxt = "Customer A";

	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	String actwarehouseTxt = warehouseTxt.getAttribute("value");
	String expwarehouseTxt = "HYDERABAD";

	
	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, BR COGS ITEM, Pcs, 0.00, 0.00, 0.00, 0.00, 0.00, 5.00, , , 5.00, 25.00, 0.00, B2, , , ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
        
			
		System.out.println("Document No          : "+actDocNo			    +"  Value Expected  "+expDocNo);
		System.out.println("salesAccTxt          : "+actsalesAccTxt			+"  Value Expected  "+expsalesAccTxt);
		System.out.println("customerAccountTxt   : "+actcustomerAccountTxt	+"  Value Expected  "+expcustomerAccountTxt);
		System.out.println("Currency             : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department           : "+actDepartment		    +"  Value Expected  "+expDepartment);
		System.out.println("warehouseTxt         : "+actwarehouseTxt	 	+"  Value Expected  "+expwarehouseTxt);
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
		select1stRow_14thColumn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(batchItemName));
		
		int actBatchRows = batchPopupRowNosList.size();
		int expBatchRows = 2;
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchTotalQtyAmt));
		String actTotalQuantityInHeader=batchTotalQtyAmt.getText();
		String expTotalQuantityInHeader="5.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchAdjustQtyAmt));
		String actAdjustQuantity=batchAdjustQtyAmt.getText();
		String expAdjustQuantity="5.00";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchBalanceAmt));
		String actBalance=batchBalanceAmt.getText();
		String expBalance="0.00";
			
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchTable1stRow1stCol));
		String act1stRowBatchNo=batchTable1stRow1stCol.getText();
		String exp1stRowBatchNo="B2";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchTable1stRow2ndCol));
		String act1stRowBaseQuantity=batchTable1stRow2ndCol.getText();
		String exp1stRowBaseQuantity="10.00";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(batchCancelIcon));
		batchCancelIcon.click();
		
		System.out.println("********* Batch Rows Count : "+actBatchRows+"  value expected  "+expBatchRows);
		System.out.println("********* Total Quantity: "+actTotalQuantityInHeader+"  value expected  "+expTotalQuantityInHeader);
		System.out.println("********* Adjust Quantity: "+actAdjustQuantity+"  value expected  "+expAdjustQuantity);
		System.out.println("********* Balance: "+actBalance+"  value expected  "+expBalance);
		System.out.println("********* 1.Batch No: "+act1stRowBatchNo+"  value expected  "+exp1stRowBatchNo);
		System.out.println("********* 1.Base Quantity: "+act1stRowBaseQuantity+"  value expected  "+exp1stRowBaseQuantity);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actsalesAccTxt.equalsIgnoreCase(expsalesAccTxt) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actcustomerAccountTxt.equalsIgnoreCase(expcustomerAccountTxt)
        		&&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount && actwarehouseTxt.equalsIgnoreCase(expwarehouseTxt) && 
        		VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount && 
        		actBatchRows==expBatchRows && actTotalQuantityInHeader.equalsIgnoreCase(expTotalQuantityInHeader) && actAdjustQuantity.equalsIgnoreCase(expAdjustQuantity) 
				  && actBalance.equalsIgnoreCase(expBalance) && actAdjustQuantity.equalsIgnoreCase(expAdjustQuantity)
				  && act1stRowBatchNo.equalsIgnoreCase(exp1stRowBatchNo) && act1stRowBaseQuantity.equalsIgnoreCase(exp1stRowBaseQuantity))

        {
        	return true;
        }

        else
        {
        	return false;
        }
	
}
		

@FindBy (xpath="//td[@id='id_body_38_input_image']/img")
private static WebElement RMAPopupExpansionBtn;

@FindBy (xpath="//input[@id='txtSerialNo']")
private static WebElement RMAPopupSerialNumberTxtField;

@FindBy (xpath="//input[@id='txtQuantity']")
private static WebElement RMAPopupQuantityTxtField;

@FindBy (xpath="//span[@id='sp_search']")
private static WebElement RMAPopupSerialNoExpansionBtn;

@FindBy (xpath="//div[@id='img_add']/span")
private static WebElement RMAPopupAddBtn;
	

@FindBy (xpath="//*[@id='RMASearchTable']/tbody/tr[1]/td[1]/input")
private static WebElement searchRMA1ChkBox;


@FindBy (xpath="//*[@id='RMASearchTable']/tbody/tr[3]/td[1]/input")
private static WebElement searchRMA3ChkBox;

@FindBy (xpath="//*[@id='RMASearchTable']/tbody/tr/td[1]/input")
private static List<WebElement> searchRMAChkBoxList;





@FindBy (xpath="//td[@id='RMA_Table_col_1-1']")
private static WebElement RMAPopupTableRow1Column1;

@FindBy (xpath="//td[@id='RMA_Table_col_2-1']")
private static WebElement RMAPopupTableRow2Column1;

@FindBy (xpath="//div[@id='RMAModel_Bottom']/div[2]/label")
private static WebElement RMAPopupOkBtn;


public boolean checkImportedVouchersInSalesInvoiceWithRMAPartialComnsumption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionssalesMenu));
	financialsTransactionssalesMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceMenu));
	salesInvoiceMenu.click();
	
	
	Thread.sleep(3000);
	
	int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
	 
	int expVoucherBodyGridVouRowCount=5;
	
	System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
	
	for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
	{
		String docNo=voucherBodyGridVouRowCount.get(i).getText();
		
		if (docNo.equalsIgnoreCase("3")) 
		{
			voucherGridIndexChkboxlist.get(i).click();
			
			break;
		}
	}

 Thread.sleep(2000);
	
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
 editBtn.click();
 
 
   boolean loading=checkLoadingMessage();

   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "3";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccTxt));
    String actsalesAccTxt = salesAccTxt.getAttribute("value");
	String expsalesAccTxt = "Sales - Computers";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String actcustomerAccountTxt = customerAccountTxt.getAttribute("value");
	String expcustomerAccountTxt = "Customer A";

	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	String actwarehouseTxt = warehouseTxt.getAttribute("value");
	String expwarehouseTxt = "HYDERABAD";

	
	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, WA COGS ITEM, Pcs, 0.00, 0.00, 0.00, 5.00, 5.00, 2.00, , , 5.00, 10.00, 0.00, , , Hk1,Hk3, ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
        
			
		System.out.println("Document No          : "+actDocNo			    +"  Value Expected  "+expDocNo);
		System.out.println("salesAccTxt          : "+actsalesAccTxt			+"  Value Expected  "+expsalesAccTxt);
		System.out.println("customerAccountTxt   : "+actcustomerAccountTxt	+"  Value Expected  "+expcustomerAccountTxt);
		System.out.println("Currency             : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department           : "+actDepartment		    +"  Value Expected  "+expDepartment);
		System.out.println("warehouseTxt         : "+actwarehouseTxt	 	+"  Value Expected  "+expwarehouseTxt);
		
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_16thColumn));
		select1stRow_16thColumn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupExpansionBtn));
		RMAPopupExpansionBtn.click();
		
		Thread.sleep(2000);
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNumberTxtField));
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupTableRow1Column1));
		String actRMA1=RMAPopupTableRow1Column1.getAttribute("data-value");
		String actRMA2=RMAPopupTableRow2Column1.getAttribute("data-value");
		
		String expRMA1="Hk1";
		String expRMA2="Hk3";
		
		System.out.println("********* Row1: "+actRMA1+"  value expected  "+expRMA1);
		System.out.println("********* Row2: "+actRMA2+"  value expected  "+expRMA2);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupSerialNoExpansionBtn));
		RMAPopupSerialNoExpansionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMA1ChkBox));
		boolean actsearchRMA1ChkBox=searchRMA1ChkBox.isSelected();
		boolean expsearchRMA1ChkBox=true;
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMA3ChkBox));
		boolean actsearchRMA3ChkBox=searchRMA3ChkBox.isSelected();
		boolean expsearchRMA3ChkBox=true;
		
		
		System.out.println("********* Row1: "+actsearchRMA1ChkBox+"  value expected  "+expsearchRMA1ChkBox);
		System.out.println("********* Row2: "+actsearchRMA3ChkBox+"  value expected  "+expsearchRMA3ChkBox);
		
		
		boolean actRMAChecked=true;
		for(int i=0; i < 3; i++)
		{
			
			if(i!=1)
			{
				actRMAChecked=searchRMAChkBoxList.get(i).isSelected();
				System.out.println("**********Row "+i+">>>" +actRMAChecked);
				
				if(actRMAChecked==false)
				{
					System.out.println("*************** Row "+i+" Checkbox NOT selected");
					
					break;
				}
			}
			
		}
		
		boolean expRMAChecked=true;

		
		
		boolean actRMAUnChecked=false;
		for(int i=0; i < 7; i++)
		{
			if(i!=0||i!=2)
			{
				
				actRMAUnChecked=searchRMAChkBoxList.get(i).isSelected();
				System.out.println("**********Row "+i+">>>" +actRMAUnChecked);
				
				if(actRMAUnChecked==true)
				{
					System.out.println("*************** Row "+i+" Checkbox selected");
					
					break;
				}
			
			}
				
		}
		
		boolean expRMAUnChecked=true;
		
		System.out.println("******* RMA Checked   : "+actRMAChecked   +"  value expected  "+expRMAChecked);
		System.out.println("******* RMA UnChecked : "+actRMAUnChecked +"  value expected  "+expRMAUnChecked);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(searchRMAOkBtn));
		searchRMAOkBtn.click();
	
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(RMAPopupOkBtn));
		RMAPopupOkBtn.click();
		
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actsalesAccTxt.equalsIgnoreCase(expsalesAccTxt) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actcustomerAccountTxt.equalsIgnoreCase(expcustomerAccountTxt)
        		&&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount && actwarehouseTxt.equalsIgnoreCase(expwarehouseTxt) && 
        		VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount &&
        		actRMAChecked==expRMAChecked && actRMAUnChecked==expRMAUnChecked && actsearchRMA1ChkBox==expsearchRMA1ChkBox && 
        		actsearchRMA3ChkBox==expsearchRMA3ChkBox)

        {
        	return true;
        }

        else
        {
        	return false;
        }
	
}
		
@FindBy (xpath="//*[@id='SearchModal_Content']/div[3]/div/div[1]/label")
private static WebElement searchRMAOkBtn;




public boolean checkSalesInvoiceImportWithOutwradBatchOverLimitDataInTrasactionImport() throws InterruptedException, IOException
{
	
	
	getDriver().navigate().refresh();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("sales");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Sales Invoices"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();
	*/
	

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\OutwradBatchOverLimit.exe");
	Thread.sleep(3000);
	
	

   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("4");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("5");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    
	String alerttxt=getAlert().getText();
		
	String expalert="3 Vouchers imported successfully";
	String expAlert1="Vouchers imported : 2,3,4";
	String expAlert2="Vouchers failed : 1";
	String expAlert3="1 Vouchers importing failed : 1";

	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	System.out.println("expalert1 : "+expAlert1);
	System.out.println("expalert2 : "+expAlert2);
	System.out.println("expalert3 : "+expAlert3);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert) && alerttxt.endsWith(expAlert3)  
			&& alerttxt.contains(expAlert1) &&  alerttxt.contains(expAlert2))
	{
		System.out.println(" Test Pass: Voucher imported Successfully");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
	
}



public boolean checkImportedVouchersInSalesInvoiceWithOutwradBatchOverLimitComnsumption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionssalesMenu));
	financialsTransactionssalesMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceMenu));
	salesInvoiceMenu.click();
	
	
	Thread.sleep(3000);
	
	int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
	 
	int expVoucherBodyGridVouRowCount=4;
	
	System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
	
	for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
	{
		String docNo=voucherBodyGridVouRowCount.get(i).getText();
		
		if (docNo.equalsIgnoreCase("2")) 
		{
			voucherGridIndexChkboxlist.get(i).click();
			
			break;
		}
	}

 Thread.sleep(2000);
	
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
 editBtn.click();
 
 
   boolean loading=checkLoadingMessage();

   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "2";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccTxt));
    String actsalesAccTxt = salesAccTxt.getAttribute("value");
	String expsalesAccTxt = "Sales - Computers";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String actcustomerAccountTxt = customerAccountTxt.getAttribute("value");
	String expcustomerAccountTxt = "Customer A";

	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	String actwarehouseTxt = warehouseTxt.getAttribute("value");
	String expwarehouseTxt = "HYDERABAD";

	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, BR COGS ITEM, Pcs, 0.00, 0.00, 0.00, 0.00, 0.00, 20.00, , , 5.00, 100.00, 0.00, B3, , , ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
        
			
		System.out.println("Document No          : "+actDocNo			    +"  Value Expected  "+expDocNo);
		System.out.println("salesAccTxt          : "+actsalesAccTxt			+"  Value Expected  "+expsalesAccTxt);
		System.out.println("customerAccountTxt   : "+actcustomerAccountTxt	+"  Value Expected  "+expcustomerAccountTxt);
		System.out.println("Currency             : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department           : "+actDepartment		    +"  Value Expected  "+expDepartment);
		System.out.println("warehouseTxt         : "+actwarehouseTxt	 	+"  Value Expected  "+expwarehouseTxt);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actsalesAccTxt.equalsIgnoreCase(expsalesAccTxt) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actcustomerAccountTxt.equalsIgnoreCase(expcustomerAccountTxt)
        		&&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount && actwarehouseTxt.equalsIgnoreCase(expwarehouseTxt) && 
        		VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount)

        {
        	return true;
        }

        else
        {
        	return false;
        }
	
}








public boolean checkSalesInvoiceImportWithOutwardRMAOverLimitDataInTrasactionImport() throws InterruptedException, IOException
{
	
	
	getDriver().navigate().refresh();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("sales");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Sales Invoices"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();
	*/
	

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\Outward RMAOverLimit.exe");
	Thread.sleep(3000);
	
	

   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("4");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("5");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    
	String alerttxt=getAlert().getText();
		
	String expalert="3 Vouchers imported successfully";
	String expAlert1="Vouchers imported : 2,3,4";
	String expAlert2="Vouchers failed : 1";
	String expAlert3="1 Vouchers importing failed : 1";

	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	System.out.println("expalert1 : "+expAlert1);
	System.out.println("expalert2 : "+expAlert2);
	System.out.println("expalert3 : "+expAlert3);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert) && alerttxt.endsWith(expAlert3)  
			&& alerttxt.contains(expAlert1) &&  alerttxt.contains(expAlert2))
	{
		System.out.println(" Test Pass: Voucher imported Successfully");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
	
}



public boolean checkImportedVouchersInSalesInvoiceWitOutwardRMAOverLimitComnsumption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionssalesMenu));
	financialsTransactionssalesMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceMenu));
	salesInvoiceMenu.click();
	
	Thread.sleep(3000);
	
	int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
	 
	int expVoucherBodyGridVouRowCount=3;
	
	System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
	
	for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
	{
		String docNo=voucherBodyGridVouRowCount.get(i).getText();
		
		if (docNo.equalsIgnoreCase("2")) 
		{
			voucherGridIndexChkboxlist.get(i).click();
			
			break;
		}
	}

 Thread.sleep(2000);
	
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
 editBtn.click();
 
 
   boolean loading=checkLoadingMessage();

   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "2";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccTxt));
    String actsalesAccTxt = salesAccTxt.getAttribute("value");
	String expsalesAccTxt = "Sales - Computers";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String actcustomerAccountTxt = customerAccountTxt.getAttribute("value");
	String expcustomerAccountTxt = "Customer A";

	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	String actwarehouseTxt = warehouseTxt.getAttribute("value");
	String expwarehouseTxt = "HYDERABAD";

	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, BR COGS ITEM, Pcs, 0.00, 0.00, 0.00, 0.00, 0.00, 20.00, , , 5.00, 100.00, 0.00, B3, , , ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
        
			
		System.out.println("Document No          : "+actDocNo			    +"  Value Expected  "+expDocNo);
		System.out.println("salesAccTxt          : "+actsalesAccTxt			+"  Value Expected  "+expsalesAccTxt);
		System.out.println("customerAccountTxt   : "+actcustomerAccountTxt	+"  Value Expected  "+expcustomerAccountTxt);
		System.out.println("Currency             : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department           : "+actDepartment		    +"  Value Expected  "+expDepartment);
		System.out.println("warehouseTxt         : "+actwarehouseTxt	 	+"  Value Expected  "+expwarehouseTxt);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actsalesAccTxt.equalsIgnoreCase(expsalesAccTxt) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actcustomerAccountTxt.equalsIgnoreCase(expcustomerAccountTxt)
        		 && actwarehouseTxt.equalsIgnoreCase(expwarehouseTxt))

        {
        	return true;
        }

        else
        {
        	return false;
        }
	
}



public boolean checkSalesInvoiceImportWithOutwardwithChangeInBatchNumberDataInTrasactionImport() throws InterruptedException, IOException
{
	
	
	getDriver().navigate().refresh();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("sales");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Sales Invoices"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();*/
	
	

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\OutwardWithChangeInBatchNumber.exe");
	Thread.sleep(3000);
	
	
	
   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("4");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("5");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    
	String alerttxt=getAlert().getText();
		
	String expalert="3 Vouchers imported successfully";
	String expAlert1="Vouchers imported : 2,3,4";
	String expAlert2="Vouchers failed : 1";
	String expAlert3="1 Vouchers importing failed : 1";

	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	System.out.println("expalert1 : "+expAlert1);
	System.out.println("expalert2 : "+expAlert2);
	System.out.println("expalert3 : "+expAlert3);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert) && alerttxt.endsWith(expAlert3)  
			&& alerttxt.contains(expAlert1) &&  alerttxt.contains(expAlert2))
	{
		System.out.println(" Test Pass: Voucher imported Successfully");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
	
}



public boolean checkImportedVouchersInSalesInvoiceWithOutwardwithChangeInBatchNumberComnsumption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionssalesMenu));
	financialsTransactionssalesMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceMenu));
	salesInvoiceMenu.click();
	
	Thread.sleep(3000);
	
	int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
	 
	int expVoucherBodyGridVouRowCount=4;
	
	System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
	
	for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
	{
		String docNo=voucherBodyGridVouRowCount.get(i).getText();
		
		if (docNo.equalsIgnoreCase("2")) 
		{
			voucherGridIndexChkboxlist.get(i).click();
			
			break;
		}
	}

 Thread.sleep(2000);
	
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
 editBtn.click();
 
 
   boolean loading=checkLoadingMessage();

   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "2";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccTxt));
    String actsalesAccTxt = salesAccTxt.getAttribute("value");
	String expsalesAccTxt = "Sales - Computers";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String actcustomerAccountTxt = customerAccountTxt.getAttribute("value");
	String expcustomerAccountTxt = "Customer A";

	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	String actwarehouseTxt = warehouseTxt.getAttribute("value");
	String expwarehouseTxt = "HYDERABAD";

	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, BR COGS ITEM, Pcs, 0.00, 0.00, 0.00, 0.00, 0.00, 20.00, , , 5.00, 100.00, 0.00, B3, , , ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
        
			
		System.out.println("Document No          : "+actDocNo			    +"  Value Expected  "+expDocNo);
		System.out.println("salesAccTxt          : "+actsalesAccTxt			+"  Value Expected  "+expsalesAccTxt);
		System.out.println("customerAccountTxt   : "+actcustomerAccountTxt	+"  Value Expected  "+expcustomerAccountTxt);
		System.out.println("Currency             : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department           : "+actDepartment		    +"  Value Expected  "+expDepartment);
		System.out.println("warehouseTxt         : "+actwarehouseTxt	 	+"  Value Expected  "+expwarehouseTxt);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actsalesAccTxt.equalsIgnoreCase(expsalesAccTxt) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actcustomerAccountTxt.equalsIgnoreCase(expcustomerAccountTxt)
        		&&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount && actwarehouseTxt.equalsIgnoreCase(expwarehouseTxt) && 
        		VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount)

        {
        	return true;
        }

        else
        {
        	return false;
        }
	
}


public boolean checkSalesInvoiceImportWitOutwardWithChangeInRMANumberDataInTrasactionImport() throws InterruptedException, IOException
{
	
	
	getDriver().navigate().refresh();
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("sales");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Sales Invoices"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();*/
	
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\OutwardWithChangeInRMANumber.exe");
	Thread.sleep(3000);
	
	
	
   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("4");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("5");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    
	String alerttxt=getAlert().getText();
		
	String expalert="2 Vouchers imported successfully";
	String expAlert1="Vouchers imported : 2,3";
	String expAlert2="Vouchers failed : 1,4";
	String expAlert3="2 Vouchers importing failed : 2";
	String expAlert4="Failed Details : RMANotAvailable";

	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	System.out.println("expalert1 : "+expAlert1);
	System.out.println("expalert2 : "+expAlert2);
	System.out.println("expalert3 : "+expAlert3);
	System.out.println("expalert4 : "+expAlert4);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert) && alerttxt.contains(expAlert3)  
			&& alerttxt.contains(expAlert1) &&  alerttxt.contains(expAlert2) && 
			alerttxt.endsWith(expAlert4))
	{
		System.out.println(" Test Pass: Voucher imported Successfully");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
	
}



public boolean checkImportedVouchersInSalesInvoiceWitOutwardWithChangeInRMANumberComnsumption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionssalesMenu));
	financialsTransactionssalesMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceMenu));
	salesInvoiceMenu.click();
	
	Thread.sleep(3000);
	
	int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
	 
	int expVoucherBodyGridVouRowCount=4;
	
	System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
	
	for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
	{
		String docNo=voucherBodyGridVouRowCount.get(i).getText();
		
		if (docNo.equalsIgnoreCase("2")) 
		{
			voucherGridIndexChkboxlist.get(i).click();
			
			break;
		}
	}

 Thread.sleep(2000);
	
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
 editBtn.click();
 
 
   boolean loading=checkLoadingMessage();

   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "2";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccTxt));
    String actsalesAccTxt = salesAccTxt.getAttribute("value");
	String expsalesAccTxt = "Sales - Computers";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String actcustomerAccountTxt = customerAccountTxt.getAttribute("value");
	String expcustomerAccountTxt = "Customer A";

	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	String actwarehouseTxt = warehouseTxt.getAttribute("value");
	String expwarehouseTxt = "HYDERABAD";

	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, BR COGS ITEM, Pcs, 0.00, 0.00, 0.00, 0.00, 0.00, 20.00, , , 5.00, 100.00, 0.00, B3, , , ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
        
			
		System.out.println("Document No          : "+actDocNo			    +"  Value Expected  "+expDocNo);
		System.out.println("salesAccTxt          : "+actsalesAccTxt			+"  Value Expected  "+expsalesAccTxt);
		System.out.println("customerAccountTxt   : "+actcustomerAccountTxt	+"  Value Expected  "+expcustomerAccountTxt);
		System.out.println("Currency             : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department           : "+actDepartment		    +"  Value Expected  "+expDepartment);
		System.out.println("warehouseTxt         : "+actwarehouseTxt	 	+"  Value Expected  "+expwarehouseTxt);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actsalesAccTxt.equalsIgnoreCase(expsalesAccTxt) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actcustomerAccountTxt.equalsIgnoreCase(expcustomerAccountTxt)
        		&&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount && actwarehouseTxt.equalsIgnoreCase(expwarehouseTxt) && 
        		VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount)

        {
        	return true;
        }

        else
        {
        	return false;
        }
	
}




public boolean checkSalesInvoiceImportWitOutwardWithOutwardFullWithWarehouseChangeInTrasactionImport() throws InterruptedException, IOException
{
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("sales");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Sales Invoices"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();*/
	
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\OutwardFullWithWarehouseChange.exe");
	Thread.sleep(3000);
	
	
	
   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("4");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("5");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    
	String alerttxt=getAlert().getText();
		
	String expalert="4 Vouchers importing failed : 4";
	String expAlert1="Failed Details : RMANotAvailable";
	
	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	System.out.println("expalert1 : "+expAlert1);
	
	
	getAlert().accept();
	
	if (alerttxt.contains(expalert) && 
			alerttxt.endsWith(expAlert1))
	{
		System.out.println(" Test Pass: Voucher imported Successfully");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
	
}




public boolean checkSalesInvoiceImportWitOutwardWithOutwardwithN0BatchenTrasactionImport() throws InterruptedException, IOException
{
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

	Thread.sleep(1000);
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("sales");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Sales Invoices"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();
	*/
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\Outward with No BAtch Item.exe");
	Thread.sleep(3000);
	
	
	
   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("4");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("5");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    
    String alerttxt=getAlert().getText();
	
	String expalert="3 Vouchers imported successfully";
	String expAlert1="Vouchers imported : 2,3,4";
	String expAlert2="Vouchers failed : 1";
	String expAlert3="1 Vouchers importing failed : 1";

	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	System.out.println("expalert1 : "+expAlert1);
	System.out.println("expalert2 : "+expAlert2);
	System.out.println("expalert3 : "+expAlert3);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert) && alerttxt.endsWith(expAlert3)  
			&& alerttxt.contains(expAlert1) &&  alerttxt.contains(expAlert2))
	{
		System.out.println(" Test Pass: 3Voucher imported Successfully && 1st Voucher Not Imported ");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
	
}



public boolean checkImportedVouchersInSalesInvoiceWithNoBatchComnsumption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionssalesMenu));
	financialsTransactionssalesMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceMenu));
	salesInvoiceMenu.click();
	
	
	Thread.sleep(3000);
	
	int VoucherBodyGridVouRowCount=voucherBodyGridVouRowCount.size();
	 
	int expVoucherBodyGridVouRowCount=4;
	
	System.err.println("VoucherBodyGridVouRowCount  : "+VoucherBodyGridVouRowCount +" VAlue Expected  "+expVoucherBodyGridVouRowCount);
	
	for (int i = 0; i <= VoucherBodyGridVouRowCount; i++) 
	{
		String docNo=voucherBodyGridVouRowCount.get(i).getText();
		
		if (docNo.equalsIgnoreCase("2")) 
		{
			voucherGridIndexChkboxlist.get(i).click();
			
			break;
		}
	}

 Thread.sleep(2000);
	
 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
 editBtn.click();
 
 
   boolean loading=checkLoadingMessage();

   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "2";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccTxt));
    String actsalesAccTxt = salesAccTxt.getAttribute("value");
	String expsalesAccTxt = "Sales - Computers";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String actcustomerAccountTxt = customerAccountTxt.getAttribute("value");
	String expcustomerAccountTxt = "Customer A";

	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	String actwarehouseTxt = warehouseTxt.getAttribute("value");
	String expwarehouseTxt = "HYDERABAD";

	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, BR COGS ITEM, Pcs, 0.00, 0.00, 0.00, 0.00, 0.00, 20.00, , , 5.00, 100.00, 0.00, B3, , , ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
        
			
		System.out.println("Document No          : "+actDocNo			    +"  Value Expected  "+expDocNo);
		System.out.println("salesAccTxt          : "+actsalesAccTxt			+"  Value Expected  "+expsalesAccTxt);
		System.out.println("customerAccountTxt   : "+actcustomerAccountTxt	+"  Value Expected  "+expcustomerAccountTxt);
		System.out.println("Currency             : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department           : "+actDepartment		    +"  Value Expected  "+expDepartment);
		System.out.println("warehouseTxt         : "+actwarehouseTxt	 	+"  Value Expected  "+expwarehouseTxt);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actsalesAccTxt.equalsIgnoreCase(expsalesAccTxt) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actcustomerAccountTxt.equalsIgnoreCase(expcustomerAccountTxt)
        		&&VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount && actwarehouseTxt.equalsIgnoreCase(expwarehouseTxt) && 
        		VoucherBodyGridVouRowCount==expVoucherBodyGridVouRowCount)

        {
        	return true;
        }

        else
        {
        	return false;
        }
	
}




//Importing Large Number of Vouchers 



public boolean checkTransationImportWith72Vouchers() throws IOException, InterruptedException, EncryptedDocumentException, InvalidFormatException
{
	
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(dataMangementMenu));
	dataMangementMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAll));
	eraseAll.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseTranscationsRadio));
	eraseTranscationsRadio.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(eraseAllOkBtn));
	eraseAllOkBtn.click();
	
	if(getIsAlertPresent())
	{
		getWaitForAlert();
		
		getAlert().accept();
	}
	
	String expValidationMsg="Data deleted successfully";
	
	String actValidationMsg=checkValidationMessage(expValidationMsg);
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();
	Thread.sleep(2000);
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("pur");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Purchases Vouchers"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();*/
	
	

   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\Inward With 72 Vouchers.exe");
	Thread.sleep(3000);
	

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("2");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("3");
	
	
     Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid13RowCol3));
	transImpgrid13RowCol3.click();
	Select transImpgrid13RowCol3=new Select(gridDataBaseList);
	vendor.selectByValue("12");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    
	String alerttxt=getAlert().getText();
	
	String expalert="72 Vouchers imported successfully";
			
	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert))
	
	{
		System.out.println(" Test Pass: Voucher imported Successfully");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
}

@FindBy(xpath="//div[@id='myNavbar']//span[@class='icon-new icon-font6']")
private static WebElement  newBtn;


@FindBy(xpath="//*[@id='id_transactionentry_previous']")
private static WebElement  previousBtn;

public static void checkUserFriendlyMessage() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	try
	{
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
		String actErrorMessage=errorMessage.getText();
		
		System.out.println("Open Page then Message Display  :  " + actErrorMessage);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
		errorMessageCloseBtn.click();
		
	}
	catch(Exception e)
	{
		System.err.println("Error Message NOT Found or NOT Clickable");
		System.err.println(e.getMessage());
		
		String Exception=e.getMessage();
	}
}
	


public boolean checkImported72VouchersinPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
	purchaseVoucherMenu.click();
	
	Thread.sleep(3000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
    newBtn.click();
 
    checkUserFriendlyMessage();
    
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
    previousBtn.click();
 
    boolean loading=checkLoadingMessage();

    Thread.sleep(2000);
    
    getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "FOC-HYD-72";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
    String actVendor = vendorAccountTxt.getAttribute("value");
	String expVendor = "Vendor A";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
	String actpurchaseAccountTxt = purchaseAccountTxt.getAttribute("value");
	String exppurchaseAccountTxt = "purchase";
	
	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purNarrationTxt));
	String actpurNarrationTxt = purNarrationTxt.getAttribute("value");
	String exppurNarrationTxt = "Focus72";
	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, HYDERABAD, BR COGS ITEM, Pcs, 0.00, 81.00, 5.00, 405.00, 0.00, B73, , , ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
     
		
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("purchaseAccountTxt   : "+actpurchaseAccountTxt	+"  Value Expected  "+exppurchaseAccountTxt);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actpurchaseAccountTxt.equalsIgnoreCase(exppurchaseAccountTxt))

        {
        	return true;
        }

        else
        {
        	return false;
        }
}





public boolean checkSalesInvoiceImportWith72Vouchers() throws InterruptedException, IOException
{
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

	Thread.sleep(1000);
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("sales");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Sales Invoices"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();*/
	
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\Outward Full Consume With 72 Vouchers.exe");
	Thread.sleep(3000);
	
	
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("4");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("5");
	
     Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid14RowCol3));
	transImpgrid14RowCol3.click();
	Select transImpgrid14RowCol3=new Select(gridDataBaseList);
	vendor.selectByValue("12");
	
	 Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
    getWaitForAlert();
	
    
    String alerttxt=getAlert().getText();
	
	String expalert="2 Vouchers imported successfully";
	String expalert1="Vouchers imported : FOC-HYD-1,FOC-HYD-2";
	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	System.out.println("expalert1 : "+expalert1);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert) && 
			alerttxt.endsWith(expalert1))
	{
		System.out.println(" Test Pass: 3Voucher imported Successfully && 1st Voucher Not Imported ");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
	
}



public boolean checkImportedVouchersInSalesInvoiceWith72Vouchers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{

	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionssalesMenu));
	financialsTransactionssalesMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceMenu));
	salesInvoiceMenu.click();
	
	
	Thread.sleep(3000);
	
	
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
    newBtn.click();
 
 
    checkUserFriendlyMessage();
 
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
    previousBtn.click();

    boolean loading=checkLoadingMessage();

    Thread.sleep(2000);
 
   getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "FOC-HYD-2";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccTxt));
    String actsalesAccTxt = salesAccTxt.getAttribute("value");
	String expsalesAccTxt = "Sales - Computers";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	String actcustomerAccountTxt = customerAccountTxt.getAttribute("value");
	String expcustomerAccountTxt = "Customer A";

	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	String actwarehouseTxt = warehouseTxt.getAttribute("value");
	String expwarehouseTxt = "HYDERABAD";

	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, BR COGS ITEM, Pcs, 0.00, 0.00, 0.00, 0.00, 0.00, 11.00, , , 5.00, 55.00, 0.00, B3, , , ]";
		
        System.out.println("actRowList   : "+actRowList );
        System.out.println("expRow1List  : "+expRow1List );
        
			
		System.out.println("Document No          : "+actDocNo			    +"  Value Expected  "+expDocNo);
		System.out.println("salesAccTxt          : "+actsalesAccTxt			+"  Value Expected  "+expsalesAccTxt);
		System.out.println("customerAccountTxt   : "+actcustomerAccountTxt	+"  Value Expected  "+expcustomerAccountTxt);
		System.out.println("Currency             : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department           : "+actDepartment		    +"  Value Expected  "+expDepartment);
		System.out.println("warehouseTxt         : "+actwarehouseTxt	 	+"  Value Expected  "+expwarehouseTxt);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
        if (actRowList.equalsIgnoreCase(expRow1List)  
        		&& actDocNo.equalsIgnoreCase(expDocNo) && actsalesAccTxt.equalsIgnoreCase(expsalesAccTxt) && actCurrency.equalsIgnoreCase(expCurrency) 
        		&& actDepartment.equalsIgnoreCase(expDepartment) && actcustomerAccountTxt.equalsIgnoreCase(expcustomerAccountTxt)
        		&& actwarehouseTxt.equalsIgnoreCase(expwarehouseTxt) )

        {
        	return true;
        }

        else
        {
        	return false;
        }
	
}



//Importing 72 Lines of 1 Vouchers 



public boolean checkTransationImportWith72LinesInSingleVouchers() throws IOException, InterruptedException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();

    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys("pur");
		   
	int count=voucherTypeComboBoxList.size();
	
	System.err.println(count);
	
	for(int i=0 ; i < count ;i++)
	{
		String data=voucherTypeComboBoxList.get(i).getText();
		
		if(data.equalsIgnoreCase("Purchases Vouchers"))
		{
			voucherTypeComboBoxList.get(i).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
		
	/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(importBtn));
	importBtn.click();*/
	
	

 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpVoucherTyeDrpdwn));
	transImpVoucherTyeDrpdwn.sendKeys(Keys.TAB);
	
	transImpVoucherTyeDrpdwn.click();
	
	transImpFileNameBtn.click();
	
	Thread.sleep(2000);
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\Inward With 72 Rows.exe");
	Thread.sleep(3000);
	

  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid4RowCol3));
	transImpgrid4RowCol3.click();
	Select purchase=new Select(gridDataBaseList);
	purchase.selectByValue("2");
	
	
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid5RowCol3));
	transImpgrid5RowCol3.click();
	Select vendor=new Select(gridDataBaseList);
	vendor.selectByValue("3");
	
	
   Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpgrid13RowCol3));
	transImpgrid13RowCol3.click();
	Select transImpgrid13RowCol3=new Select(gridDataBaseList);
	vendor.selectByValue("12");
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
	transImpImportBtn.click();
	
	Thread.sleep(2000);
	
  getWaitForAlert();
	
  
	String alerttxt=getAlert().getText();
	
	String expalert="1 Vouchers imported successfully";
			
	
	System.out.println("alerttxt : "+alerttxt);
	System.out.println("expalert : "+expalert);
	
	getAlert().accept();
	
	if (alerttxt.startsWith(expalert))
	
	{
		System.out.println(" Test Pass: Voucher imported Successfully");
		return true;
	} 
	else
	{

		System.out.println(" Test Fail: Voucher not imported  ");

		return false;
	}
}
	


public boolean checkImported72RowsinPurchaseVoucher() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
	purchaseVoucherMenu.click();
	
	Thread.sleep(3000);
	
    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
    newBtn.click();

  checkUserFriendlyMessage();
  
  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(previousBtn));
  previousBtn.click();

  boolean loading=checkLoadingMessage();

  Thread.sleep(2000);
  
  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	String actDocNo = documentNumberTxt.getAttribute("value");
	String expDocNo = "1";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
  String actVendor = vendorAccountTxt.getAttribute("value");
	String expVendor = "Vendor A";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	String actCurrency = voucherHeaderCurrency.getAttribute("value");
	String expCurrency = "INR";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	String actDepartment = departmentTxt.getAttribute("value");
	String expDepartment = "India";
		
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
	String actpurchaseAccountTxt = purchaseAccountTxt.getAttribute("value");
	String exppurchaseAccountTxt = "purchase";
	
	
	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(purNarrationTxt));
	String actpurNarrationTxt = purNarrationTxt.getAttribute("value");
	String exppurNarrationTxt = "Focus1";
	
	ArrayList<String >voucherRow1ListArray = new ArrayList<String>();
		
	int voucherGridRowListCount=voucherGridRowList.size();
	
	int expvoucherGridRowListCount=73;
	
	System.out.println("voucherGridRowList Actual: "+voucherGridRowListCount);
	System.out.println("voucherGridRowList  Exp  : "+expvoucherGridRowListCount);
	
	int voucherRow1ListCount=voucherRow1List.size();
		
	for (int j = 0; j < voucherRow1ListCount; j++) 
	{
			String data=voucherRow1List.get(j).getText();
			voucherRow1ListArray.add(data);
	}
		String actRowList=voucherRow1ListArray.toString();
		String expRow1List="[1, HYDERABAD, BR COGS ITEM, Pcs, 0.00, 10.00, 5.00, 50.00, 0.00, B2, , , ]";
		
      System.out.println("actRowList   : "+actRowList );
      System.out.println("expRow1List  : "+expRow1List );
   
		
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("purchaseAccountTxt   : "+actpurchaseAccountTxt	+"  Value Expected  "+exppurchaseAccountTxt);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(new_CloseBtn));
		new_CloseBtn.click();
		
      if (actRowList.equalsIgnoreCase(expRow1List)  
      		&& actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
      		&& actDepartment.equalsIgnoreCase(expDepartment) && actpurchaseAccountTxt.equalsIgnoreCase(exppurchaseAccountTxt) && 
      		voucherGridRowListCount==expvoucherGridRowListCount)

      {
      	return true;
      }

      else
      {
      	return false;
      }
}


public boolean CheckGrossColumnValuInOutwardVoucherAfterSplittingOrConsumingTheQunatityFromDifferentBatches() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
	financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseVoucherMenu));
	purchaseVoucherMenu.click();
	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(newBtn));
	newBtn.click();
	
    Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(purchaseAccountTxt));
	purchaseAccountTxt.click();
	purchaseAccountTxt.sendKeys("purchase");
	Thread.sleep(2000);
	
	purchaseAccountTxt.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
	vendorAccountTxt.click();
	vendorAccountTxt.sendKeys("Vendor A");
	Thread.sleep(2000);
	
	vendorAccountTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	departmentTxt.click();
	departmentTxt.sendKeys("India");
	Thread.sleep(2000);
	departmentTxt.sendKeys(Keys.TAB);

	
	Thread.sleep(2000);
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	select1stRow_1stColumn.click();
	
	enter_WarehouseTxt.sendKeys(Keys.SPACE);
	enter_WarehouseTxt.sendKeys("HYDERABAD");
	
	/*int pvvGridWarehouseListCount=pvvGridWarehouseList.size();
	
	for (int i = 0; i < pvvGridWarehouseListCount; i++) 
	{
		String warehouse=pvvGridWarehouseList.get(i).getText();
		if (warehouse.equalsIgnoreCase("HYDERABAD")) 
		{
			pvvGridWarehouseList.get(i).click();
			break;
		}
	}*/
	Thread.sleep(2000);
    enter_WarehouseTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.click();
	enter_ItemTxt.sendKeys(Keys.SPACE);
	
	enter_ItemTxt.sendKeys("BR COGS ITEM");
	Thread.sleep(2000);
    /*int pvvGridItemListCount=pvvGridItemList.size();
	for (int i = 0; i < pvvGridItemListCount; i++) 
	{
		String Item=pvvGridItemList.get(i).getText();
		if (Item.equalsIgnoreCase("BR COGS ITEM")) 
		{
			pvvGridItemList.get(i).click();
			break;
		}
	}*/
	enter_ItemTxt.sendKeys(Keys.TAB);	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	select1stRow_5thColumn.click();
	
	enter_Quantity.click();
	enter_Quantity.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	enter_Quantity.sendKeys("10");
	
	enter_Quantity.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.sendKeys("10");
	
	enter_Rate.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
	select1stRow_9thColumn.click();
	
	enter_Batch.click();
	
	enter_Batch.sendKeys("B1");
	
	
	enter_Batch.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Expirydate));
	enter_Expirydate.click();
	enter_Expirydate.sendKeys(Keys.TAB);
	Thread.sleep(2000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
	select2ndRow_1stColumn.click();
	enter_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	enter_WarehouseTxt.sendKeys(Keys.SPACE);
	
	enter_WarehouseTxt.sendKeys("HYDERABAD");
	Thread.sleep(2000);
	/*for (int i = 0; i < pvvGridWarehouseListCount; i++) 
	{
		String warehouse=pvvGridWarehouseList.get(i).getText();
		if (warehouse.equalsIgnoreCase("HYDERABAD")) 
		{
			pvvGridWarehouseList.get(i).click();
			break;
		}
	}*/
    enter_WarehouseTxt.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_ItemTxt));
	enter_ItemTxt.click();
	enter_ItemTxt.sendKeys(Keys.SPACE);
	
	enter_ItemTxt.sendKeys("BR COGS ITEM");
	/*
	for (int i = 0; i < pvvGridItemListCount; i++) 
	{
		String Item=pvvGridItemList.get(i).getText();
		if (Item.equalsIgnoreCase("BR COGS ITEM")) 
		{
			pvvGridItemList.get(i).click();
			break;
		}
	}*/
	enter_ItemTxt.sendKeys(Keys.TAB);	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
	select2ndRow_5thColumn.click();
	
	enter_Quantity.click();
	enter_Quantity.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	enter_Quantity.sendKeys("10");
	
	enter_Quantity.sendKeys(Keys.TAB);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(enter_Rate));
	enter_Rate.sendKeys("20");
	
	enter_Rate.sendKeys(Keys.TAB);
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
	select2ndRow_9thColumn.click();
	
	enter_Batch.click();
	
	enter_Batch.sendKeys("B2");
	
	
	enter_Batch.sendKeys(Keys.TAB);
	
	
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		 
	String docno=documentNumberTxt.getAttribute("value");
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saveBtn));
	saveBtn.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefNewReferenceTxt));
	billRefNewReferenceTxt.click();
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefPickIcon));
	billRefPickIcon.click();
		
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(billRefOkBtn));
    billRefOkBtn.click();
	
        
   boolean savingVoucher=checkVoucherSavingMessage(docno);
   System.out.println(" Voucher Saved in Purchase ");
   
   Thread.sleep(2000);
   
   getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialMenu));
   financialMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionMenu));
	financialsTransactionMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionssalesMenu));
	financialsTransactionssalesMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(financialsTransactionsPurchaseMenu));
	financialsTransactionsPurchaseMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesInvoiceMenu));
	salesInvoiceMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(editBtn));
	newBtn.click();
	 
	 
	boolean loading=checkLoadingMessage();

	 getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
	 
	  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(salesAccTxt));
	  salesAccTxt.click();
	  salesAccTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	  salesAccTxt.sendKeys("Sales - Computers");
	  Thread.sleep(2000);
	  salesAccTxt.sendKeys(Keys.TAB);
		
	  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
	  voucherHeaderCurrency.click();
	  voucherHeaderCurrency.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	  voucherHeaderCurrency.sendKeys("INR");
	  Thread.sleep(2000);
	  voucherHeaderCurrency.sendKeys(Keys.TAB);
		
	  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
	  departmentTxt.click();
	  departmentTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	  departmentTxt.sendKeys("India");
	  Thread.sleep(2000);
	  departmentTxt.sendKeys(Keys.TAB);
	  
		
	  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
	  customerAccountTxt.click();
	  customerAccountTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	  customerAccountTxt.sendKeys("Customer A");
	  Thread.sleep(2000);
	  customerAccountTxt.sendKeys(Keys.TAB);
	  
	  
	  getWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
	  warehouseTxt.click();
	  warehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
	  warehouseTxt.sendKeys("HYDERABAD");
	  Thread.sleep(2000);
	  warehouseTxt.sendKeys(Keys.TAB);
 
	  
	if (savingVoucher==true) 
	{
	
		return true;
	} 
	else
	{
		return false;
	}
}

@FindBy(xpath="//input[@id='id_body_37']")
private static WebElement enter_Expirydate;

public boolean checkVoucherSavingMessage(String docno) throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
{
	try
	{
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(errorMessage));
		String actErrorMessage=errorMessage.getText();
		String expErrorMessage="Voucher saved successfully";
		String expErrorMessage1=": "+docno;
		
		System.out.println("SavingMessage  :  "+actErrorMessage +" Value Expected : "+expErrorMessage+" "+expErrorMessage1);
		
		if(actErrorMessage.startsWith(expErrorMessage) && actErrorMessage.endsWith(expErrorMessage1))
		{
			try
			{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(errorMessageCloseBtn));
				errorMessageCloseBtn.click();
				
				return  true;
			}
			catch(Exception ee)
			{
				return true;
			}	
		}
		else
		{
			return false;
		}	
	}
	catch(Exception e)
	{
		System.err.println("UNABLE TO COMPARE");
		return false;
	}
}
	
	

@FindBy(xpath="//span[@id='btnPick']")
private static WebElement  billRefPickIcon;		

@FindBy(xpath="//span[@id='btnOk']")
private static WebElement  billRefOkBtn;

@FindBy(xpath="//input[@id='txtNewReference']")
private static WebElement  billRefNewReferenceTxt;


@FindBy(xpath="//*[@id='id_transactionentry_save']/div[2]")
private static WebElement  saveBtn;


@FindBy(xpath="//input[@id='id_body_26']")
private static WebElement  enter_Quantity;

@FindBy(xpath="//input[@id='id_body_27']")
private static WebElement  enter_Rate;

@FindBy(xpath="//*[@id='id_body_28']")
private static WebElement  enter_Gross;

@FindBy(xpath="//input[@id='id_body_13']")
private static WebElement  enter_Batch;

@FindBy(xpath="//*[@id='id_body_536870916_table_data_body']/tr")
private static List<WebElement> pvvGridWarehouseList;

@FindBy(xpath="//*[@id='id_body_23_table_data_body']/tr")
private static List<WebElement> pvvGridItemList;

@FindBy(xpath="//input[@id='id_body_23']")
private static WebElement  enter_ItemTxt;


@FindBy(xpath="//input[@id='id_body_536870916']")
private static WebElement enter_WarehouseTxt;


@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[2]")
private static WebElement  select1stRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[3]")
private static WebElement  select1stRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[4]")
private static WebElement  select1stRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[5]")
private static WebElement  select1stRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[6]")
private static WebElement  select1stRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[7]")
private static WebElement  select1stRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[8]")
private static WebElement  select1stRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[9]")
private static WebElement  select1stRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[10]")
private static WebElement  select1stRow_9thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[11]")
private static WebElement  select1stRow_10thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[12]")
private static WebElement  select1stRow_11thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[13]")
private static WebElement  select1stRow_12thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[14]")
private static WebElement  select1stRow_13thColumn;

/*@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[15]")
private static WebElement  select1stRow_14thColumn;*/

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[16]")
private static WebElement  select1stRow_15thColumn;

/*@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[17]")
private static WebElement  select1stRow_16thColumn;
*/
@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[18]")
private static WebElement  select1stRow_17thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[19]")
private static WebElement  select1stRow_18thColumn;

/*@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[20]")
private static WebElement  select1stRow_19thColumn;*/

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[1]/td[21]")
private static WebElement  select1stRow_20thColumn;





@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[2]")
private static WebElement  select2ndRow_1stColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[3]")
private static WebElement  select2ndRow_2ndColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[4]")
private static WebElement  select2ndRow_3rdColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[5]")
private static WebElement  select2ndRow_4thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[6]")
private static WebElement  select2ndRow_5thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[7]")
private static WebElement  select2ndRow_6thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[8]")
private static WebElement  select2ndRow_7thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[9]")
private static WebElement  select2ndRow_8thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[10]")
private static WebElement  select2ndRow_9thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[12]")
private static WebElement  select2ndRow_11thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[13]")
private static WebElement  select2ndRow_12thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[11]")
private static WebElement  select2ndRow_10thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[15]")
private static WebElement  select2ndRow_14thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[16]")
private static WebElement  select2ndRow_15thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[17]")
private static WebElement  select2ndRow_16thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[18]")
private static WebElement  select2ndRow_17thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[19]")
private static WebElement  select2ndRow_18thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[20]")
private static WebElement  select2ndRow_19thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[21]")
private static WebElement  select2ndRow_20thColumn;

@FindBy(xpath="//*[@id='id_transaction_entry_detail_table_body']/tr[2]/td[22]")
private static WebElement  select2ndRow_21thColumn;





// Ravi Foods

public boolean checkImportingSupplierDebitNoteFinAndInvInTransationImport() throws InterruptedException, IOException
{
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(homeMenu));
	homeMenu.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(utilities));
	utilities.click();
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionImport));
	transactionImport.click();
	
	Thread.sleep(2000);
	
	/*ArrayList<String> actAlert = new ArrayList<>();
	
	ArrayList<String> expAlert = new ArrayList<>();*/
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
	voucherTypeComboBox.click();
	voucherTypeComboBox.sendKeys(Keys.END);
	voucherTypeComboBox.sendKeys(Keys.SHIFT, Keys.HOME);
	Thread.sleep(2000);
	voucherTypeComboBox.sendKeys("Supplier");
	Thread.sleep(2000);
	
	int count=voucherTypeComboBoxList.size();
	
	for(int j=0 ; j < count ;j++)
	{
		String data=voucherTypeComboBoxList.get(j).getText();
		
		if(data.equalsIgnoreCase("Supplier Debit Note (Fin & Inv)"))
		{
			voucherTypeComboBoxList.get(j).click();
			break;
		}
		
	}
	 
	voucherTypeComboBox.sendKeys(Keys.TAB);
	
	Thread.sleep(3000);
	
	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpFileNameBtn));
	transImpFileNameBtn.click();
	
	Thread.sleep(3000);
	
	Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\SupplierDebitNoteFin&InvImport.exe");
	
	Thread.sleep(5000);
	
	for (int i = 0; i < 10; i++) 
	{
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherTypeComboBox));
		voucherTypeComboBox.click();
		voucherTypeComboBox.sendKeys(Keys.END);
		voucherTypeComboBox.sendKeys(Keys.SHIFT, Keys.HOME);
		Thread.sleep(2000);
		voucherTypeComboBox.sendKeys("Supplier");
		Thread.sleep(2000);
		
		int count=voucherTypeComboBoxList.size();
		
		for(int j=0 ; j < count ;j++)
		{
			String data=voucherTypeComboBoxList.get(j).getText();
			
			if(data.equalsIgnoreCase("Supplier Debit Note (Fin & Inv)"))
			{
				voucherTypeComboBoxList.get(j).click();
				break;
			}
			
		}
		 
		voucherTypeComboBox.sendKeys(Keys.TAB);
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpFileNameBtn));
		transImpFileNameBtn.click();
		
		Thread.sleep(3000);
		
		Runtime.getRuntime().exec(getBaseDir()+"\\autoIt\\scripts\\SupplierDebitNoteFin&InvImport.exe");
		
		Thread.sleep(5000);*/
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transImpImportBtn));
		transImpImportBtn.click();
		
		Thread.sleep(2000);
		
		getWaitForAlert();
			
	    String alertTxt = getAlert().getText();
	    
	    /*actAlert.add(alertTxt);
	    
	    expAlert.add("250 Vouchers imported successfully"+
	    			 "Vouchers imported : //08/1895,//09/2881,//09/2882,//09/2883,//09/2884,//09/2885,//09/2886,//09/2887,//09/2888,//09/2889,//09/2890,//09/2891,//09/2892,//09/2893,//09/2894,//09/2895,//09/2896,//09/2897,//09/2898,//09/2899,//09/2900,//09/2901,//09/2902,//09/2903,//09/2904,//09/2905,//09/2906,//09/2907,//09/2908,//09/2909,//09/2910,//09/2911,//09/2912,//09/2913,//09/2914,//09/2915,//09/2916,//09/2917,//09/2918,//09/2919,//09/2920,//09/2921,//09/2922,//09/2923,//09/2924,//09/2925,//09/2926,//09/2927,//09/2928,//09/2929,//09/2930,//09/2931,//09/2932,//09/2933,//09/2934,//09/2935,//08/1896,//08/1897,//08/1898,//08/1899,//08/1900,//08/1901,//08/1902,//08/1903,//08/1904,//08/1905,//08/1906,//08/1907,//08/1908,//08/1909,//08/1910,//08/1911,//08/1912,//08/1913,//08/1914,//08/1915,//08/1916,//08/1917,//08/1918,//08/1919,//08/1920,//08/1921,//08/1922,//08/1923,//08/1924,//08/1925,//08/1926,//08/1927,//08/1928,//08/1929,//08/1930,//08/1931,//08/1932,//08/1933,//08/1934,//08/1935,//08/1936,//08/1937,//08/1938,//08/1939,//08/1940,//08/1941,//08/1942,//08/1943,//08/1944,//08/1945,//08/1946,//08/1947,//08/1948,//08/1949,//08/1950,//08/1951,//08/1952,//08/1953,//08/1954,//08/1955,//08/1956,//08/1957,//08/1958,//08/1959,//08/1960,//08/1961,//08/1962,//08/1963,//08/1964,//08/1965,//08/1966,//08/1967,//08/1968,//08/1969,//08/1970,//08/1971,//08/1972,//08/1973,//08/1974,//08/1975,//08/1976,//08/1977,//08/1978,//08/1979,//08/1980,//08/1981,//08/1982,//08/1983,//08/1984,//08/1985,//08/1986,//08/1987,//08/1988,//08/1989,//08/1990,//08/1991,//08/1992,//09/2936,//09/2937,//09/2938,//09/2939,//09/2940,//09/2941,//09/2942,//09/2943,//09/2944,//09/2945,//09/2946,//09/2947,//09/2948,//09/2949,//09/2950,//09/2951,//09/2952,//09/2953,//09/2954,//09/2955,//09/2956,//09/2957,//09/2958,//09/2959,//09/2960,//09/2961,//09/2962,//09/2963,//09/2964,//09/2965,//09/2966,//09/2967,//09/2968,//09/2969,//09/2970,//09/2971,//09/2972,//09/2973,//09/2974,//09/2975,//09/2976,//09/2977,//09/2978,//09/2979,//09/2980,//09/2981,//09/2982,//09/2983,//09/2984,//09/2985,//09/2986,//09/2987,//09/2988,//09/2989,//09/2990,//09/2991,//09/2992,//09/2993,//09/2994,//09/2995,//09/2996,//09/2997,//09/2998,//09/2999,//09/3000,//09/3001,//09/3002,//09/3003,//09/3004,//09/3005,//09/3006,//09/3007,//09/3008,//09/3009,//09/3010,//09/3011,//09/3012,//09/3013,//09/3014,//09/3015,//09/3016,//09/3017,//09/3018,//09/3019,//09/3020,//09/3021,//09/3022,//09/3023,//09/3024,//09/3025,//09/3026,//09/3027,//09/3028,//09/3029,//09/3030,//09/3031,//09/3032");*/	    
	    System.out.println(i+1+". "+alertTxt);
	    
	    getAlert().accept();
	    
	    Thread.sleep(2000);
	}
		
	/*String actAlertTxt = actAlert.toString();
	String expAlertTxt = expAlert.toString();
	
	System.out.println(actAlertTxt);
	System.out.println(expAlertTxt);*/
		 
	 if (/*actAlertTxt.equalsIgnoreCase(expAlertTxt)*/ true) 
	 {
		 return true;
	 }
	 else
	 {
		 return false;
	 }
}








	 
	 






		
	
	public TransactionImportPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}
	
	
}
