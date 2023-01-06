package com.focus.Pages;

import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.List;

import org.apache.http.impl.conn.LoggingSessionOutputBuffer;
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

public class Focus10InventoryReportsPage extends BaseEngine
{
	private static String resPass="Pass";
    private static String resFail="Fail";
    private static ExcelReader excelReader;
    private static String xlfile=getBaseDir()+"\\src\\main\\resources\\testdata\\FocusTestData.xlsx";
    private static String xlSheetName = "Focus10InventoryReportsPage";
    
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
	

	public boolean checkLoginTOAutomationCompany() throws InterruptedException
	{

		 getDriver().navigate().refresh();
		 Thread.sleep(1999);
		 
		// getDriver().navigate().refresh();
		 //Thread.sleep(1999);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt="su";
				      
		String pawslt="su";
				      
		LoginPage.enterUserName(unamelt);

		Thread.sleep(1000);
				
		LoginPage.enterPassword(pawslt);

		String compname="Automation Company : 08/10/2020(020)";

		 Select oSelect = new Select(companyDropDownList);
		 oSelect.selectByValue("36");
		 
		/* List <WebElement> elementCount = oSelect.getOptions();

		 int cqSize = elementCount.size();
		 
		 System.out.println("CompanyDropdownList Count :"+cqSize);

		    int i;
		  
			for(i=0; i<elementCount.size(); i++) 
			{
			
			  elementCount.get(i).getText();
			  
		      String optionName = elementCount.get(i).getText();
		      
		     if(optionName.toUpperCase().startsWith(compname.toUpperCase()))
		     {
			  System.out.println("q"+elementCount.get(i).getText());
				  elementCount.get(i).click();
			 }
		 
		    }*/
			
	        Thread.sleep(1000);
			
			LoginPage.clickOnSignInBtn();
				
			//checkRefershPopOnlogin();
			        
		//checkPopUpWindow();

		//Thread.sleep(2000);
				           	
		String userInfo=userNameDisplay.getText();
				    	
		System.out.println("User Info : "+userInfo);
				    	
		System.out.println("User Info Capture Text :"+userNameDisplay.getText());
				    	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(companyLogo));
				    	
		companyLogo.click();
				    	
		if(homeMenu.isDisplayed()==true)
		{

			System.out.println("Test Pass :Logined to Automation Company");
			return true;
			
		}
		else
		{
			System.out.println("Test Fail :Logined to BRS Company");
				return false;

			}
		
	}
	
	
	@FindBy(xpath="//*[@id='135']")
	public static WebElement inventoryMenu;
	
	@FindBy(xpath="//*[@id='200']")
	public static WebElement inventoryReportsMenu;
	
	@FindBy(xpath="//*[@id='202']")
	public static WebElement itemQuery;
	
	@FindBy(xpath="//*[@id='productQuery']")
	public static WebElement iq_ItemTxt;
	
	@FindBy(xpath="//*[@id='dv_Panelbody']/div[1]/div[2]/div/input")
	public static WebElement iq_getStockBtn;
	
	@FindBy(xpath="//span[@id='openingStock']")
	private static WebElement iq_OpeningStocks;
	
	@FindBy(xpath="//span[@id='currentStock']")
	private static WebElement iq_CurrentStock;
	
	@FindBy(xpath="//span[@id='avgStockRate']")
	private static WebElement iq_AvgStockRate;
	
	@FindBy(xpath="//span[@id='value']")
	private static WebElement iq_Value;
	
	@FindBy(xpath="//span[@id='iBaseUnit']")
	private static WebElement iq_BaseUnit;
	
	@FindBy(xpath="//span[@id='purchaseOrdered']")
	private static WebElement iq_PendingPurchaseOrders;
	
	@FindBy(xpath="//span[@id='salesOrdered']")
	private static WebElement iq_PendingSalesOrders;
	
	@FindBy(xpath="//span[@id='quantityToBeOrdered']")
	private static WebElement iq_QtyToBeOrdered;
	
	@FindBy(xpath="//*[@id='Alternatequantity']")
	private static WebElement iq_AlternateQuantity;
	
	// Purchase Table
	
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[1]")
			private static WebElement pur1strow_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[2]")
			private static WebElement pur1strow_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[3]")
			private static WebElement pur1strow_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[4]")
			private static WebElement pur1strow_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[5]")
			private static WebElement pur1strow_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[6]")
			private static WebElement pur1strow_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[7]")
			private static WebElement pur1strow_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[8]")
			private static WebElement pur1strow_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[9]")
			private static WebElement pur1strow_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[1]/td[10]")
			private static WebElement pur1strow_10thColumn;
			
			
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[1]")
			private static WebElement pur2ndrow_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[2]")
			private static WebElement pur2ndrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[3]")
			private static WebElement pur2ndrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[4]")
			private static WebElement pur2ndrow_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[5]")
			private static WebElement pur2ndrow_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[6]")
			private static WebElement pur2ndrow_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[7]")
			private static WebElement pur2ndrow_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[8]")
			private static WebElement pur2ndrow_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[9]")
			private static WebElement pur2ndrow_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[10]")
			private static WebElement pur2ndrow_10thColumn;
			
			
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[1]")
			private static WebElement pur7thRow_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[2]")
			private static WebElement pur3rdrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[3]")
			private static WebElement pur3rdrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[4]")
			private static WebElement pur3rdrow_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[5]")
			private static WebElement pur3rdrow_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[6]")
			private static WebElement pur3rdrow_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[7]")
			private static WebElement pur3rdrow_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[8]")
			private static WebElement pur3rdrow_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[9]")
			private static WebElement pur3rdrow_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[3]/td[10]")
			private static WebElement pur3rdrow_10thColumn;
			
			
			// 4TH
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[1]")
			private static WebElement pur4throw_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[2]")
			private static WebElement pur4throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[3]")
			private static WebElement pur4throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[4]")
			private static WebElement pur4throw_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[5]")
			private static WebElement pur4throw_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[6]")
			private static WebElement pur4throw_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[7]")
			private static WebElement pur4throw_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[8]")
			private static WebElement pur4throw_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[9]")
			private static WebElement pur4throw_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[4]/td[10]")
			private static WebElement pur4throw_10thColumn;
			
			// 5TH
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[1]")
			private static WebElement pur5throw_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[2]")
			private static WebElement pur5throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[3]")
			private static WebElement pur5throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[4]")
			private static WebElement pur5throw_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[5]")
			private static WebElement pur5throw_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[6]")
			private static WebElement pur5throw_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[7]")
			private static WebElement pur5throw_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[8]")
			private static WebElement pur5throw_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[9]")
			private static WebElement pur5throw_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[5]/td[10]")
			private static WebElement pur5throw_10thColumn;
			
			// 6TH
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[1]")
			private static WebElement pur6throw_1stColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[2]")
			private static WebElement pur6throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[3]")
			private static WebElement pur6throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[4]")
			private static WebElement pur6throw_4thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[5]")
			private static WebElement pur6throw_5thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[6]")
			private static WebElement pur6throw_6thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[7]")
			private static WebElement pur6throw_7thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[8]")
			private static WebElement pur6throw_8thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[9]")
			private static WebElement pur6throw_9thColumn;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[6]/td[10]")
			private static WebElement pur6throw_10thColumn;
			
			
			@FindBy(xpath="//td[contains(text(),'Quantity Purchased Till Date:')]")
			private static WebElement qtyPurchasedTillDate;
			
			@FindBy(xpath="//*[@id='purchaseTable']/tbody/tr[2]/td[2]")
			private static WebElement purTotalWa;
					
					
			// Sales Table
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[1]")
			private static WebElement sal1strow_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[2]")
			private static WebElement sal1strow_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[3]")
			private static WebElement sal1strow_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[4]")
			private static WebElement sal1strow_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[5]")
			private static WebElement sal1strow_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[6]")
			private static WebElement sal1strow_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[7]")
			private static WebElement sal1strow_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[8]")
			private static WebElement sal1strow_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[9]")
			private static WebElement sal1strow_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[1]/td[10]")
			private static WebElement sal1strow_10thColumn;
			
			// 2nd
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[1]")
			private static WebElement sal2ndrow_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[2]")
			private static WebElement sal2ndrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[3]")
			private static WebElement sal2ndrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[4]")
			private static WebElement sal2ndrow_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[5]")
			private static WebElement sal2ndrow_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[6]")
			private static WebElement sal2ndrow_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[7]")
			private static WebElement sal2ndrow_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[8]")
			private static WebElement sal2ndrow_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[9]")
			private static WebElement sal2ndrow_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[2]/td[10]")
			private static WebElement sal2ndrow_10thColumn;
			
			// 3rd
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[1]")
			private static WebElement sal3rdrow_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[2]")
			private static WebElement sal3rdrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[3]")
			private static WebElement sal3rdrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[4]")
			private static WebElement sal3rdrow_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[5]")
			private static WebElement sal3rdrow_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[6]")
			private static WebElement sal3rdrow_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[7]")
			private static WebElement sal3rdrow_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[8]")
			private static WebElement sal3rdrow_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[9]")
			private static WebElement sal3rdrow_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[10]")
			private static WebElement sal3rdrow_10thColumn;
			
			// 4th
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[1]")
			private static WebElement sal4throw_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[2]")
			private static WebElement sal4throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[3]")
			private static WebElement sal4throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[4]")
			private static WebElement sal4throw_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[5]")
			private static WebElement sal4throw_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[6]")
			private static WebElement sal4throw_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[7]")
			private static WebElement sal4throw_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[8]")
			private static WebElement sal4throw_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[9]")
			private static WebElement sal4throw_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[4]/td[10]")
			private static WebElement sal4throw_10thColumn;
			
			//5th
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[1]")
			private static WebElement sal5throw_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[2]")
			private static WebElement sal5throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[3]")
			private static WebElement sal5throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[4]")
			private static WebElement sal5throw_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[5]")
			private static WebElement sal5throw_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[6]")
			private static WebElement sal5throw_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[7]")
			private static WebElement sal5throw_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[8]")
			private static WebElement sal5throw_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[9]")
			private static WebElement sal5throw_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[5]/td[10]")
			private static WebElement sal5throw_10thColumn;
			
			//6th
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[1]")
			private static WebElement sal6throw_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[2]")
			private static WebElement sal6throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[3]")
			private static WebElement sal6throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[4]")
			private static WebElement sal6throw_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[5]")
			private static WebElement sal6throw_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[6]")
			private static WebElement sal6throw_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[7]")
			private static WebElement sal6throw_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[8]")
			private static WebElement sal6throw_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[9]")
			private static WebElement sal6throw_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[6]/td[10]")
			private static WebElement sal6throw_10thColumn;
			
			// 7th
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[1]")
			private static WebElement sal7throw_1stColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[2]")
			private static WebElement sal7throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[3]")
			private static WebElement sal7throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[4]")
			private static WebElement sal7throw_4thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[5]")
			private static WebElement sal7throw_5thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[6]")
			private static WebElement sal7throw_6thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[7]")
			private static WebElement sal7throw_7thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[8]")
			private static WebElement sal7throw_8thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[9]")
			private static WebElement sal7throw_9thColumn;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[7]/td[10]")
			private static WebElement sal7throw_10thColumn;

			
			@FindBy(xpath="//td[contains(text(),'Quantity Sold Till Date:')]")
			private static WebElement qtySoldTillDate;
			
			@FindBy(xpath="//*[@id='salesTable']/tbody/tr[3]/td[2]")
			private static WebElement salTotalWa;
					
			
			
			// Warehouse Table
			
			@FindBy(xpath="//div[contains(text(),'Warehouse')]")
			private static WebElement BalancebyWareHouselable;
			
			@FindBy(xpath="/html/body/section/div[2]/div[3]/div[1]/div[1]/div[2]/div[8]/div[1]/div[1]/div[1]/span")
			private static WebElement BalancebyWareHouseExpandBtn;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[1]/td[2]")
			private static WebElement wh1strow_2ndColumn;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[1]/td[3]")
			private static WebElement wh1strow_3rdColumn;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[1]/td[4]")
			private static WebElement wh1strow_4thColumn;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[1]/td[5]")
			private static WebElement wh1strow_5thColumn;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[1]/td[6]")
			private static WebElement wh1strow_6thColumn;
			
			
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[2]/td[2]")
			private static WebElement wh2ndrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[2]/td[3]")
			private static WebElement wh2ndrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[2]/td[4]")
			private static WebElement wh2ndrow_4thColumn;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[2]/td[5]")
			private static WebElement wh2ndrow_5thColumn;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[2]/td[6]")
			private static WebElement wh2ndrow_6thColumn;
			
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[3]/td[2]")
			private static WebElement whStockTotal;
			
			@FindBy(xpath="//*[@id='tagWiseData']/tbody/tr[3]/td[4]")
			private static WebElement whValueTotal;
			
			
			// Selling rate and buying rates
			
			@FindBy(xpath="//div[contains(text(),'Selling rates and Buying rates')]")
			private static WebElement sellingAndBuyingRateslable;
			
			@FindBy(xpath="//span[@id='ProdQuery_ShowSellandBuyRatesPopup']")
			private static WebElement sellingAndBuyingRatesExpandBtn;
			
			// 1st 
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[1]/td[2]")
			private static WebElement sbr1strow_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[1]/td[3]")
			private static WebElement sbr1strow_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[1]/td[4]")
			private static WebElement sbr1strow_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[1]/td[5]")
			private static WebElement sbr1strow_5thColumn;
			
			// 2nd 
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[2]/td[2]")
			private static WebElement sbr2ndrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[2]/td[3]")
			private static WebElement sbr2ndrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[2]/td[4]")
			private static WebElement sbr2ndrow_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[2]/td[5]")
			private static WebElement sbr2ndrow_5thColumn;
			
			//3rd
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[3]/td[2]")
			private static WebElement sbr3rdrow_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[3]/td[3]")
			private static WebElement sbr3rdrow_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[3]/td[4]")
			private static WebElement sbr3rdrow_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[3]/td[5]")
			private static WebElement sbr3rdrow_5thColumn;
			
			// 4th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[4]/td[2]")
			private static WebElement sbr4throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[4]/td[3]")
			private static WebElement sbr4throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[4]/td[4]")
			private static WebElement sbr4throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[4]/td[5]")
			private static WebElement sbr4throw_5thColumn;
			
			
			
			// 5th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[5]/td[2]")
			private static WebElement sbr5throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[5]/td[3]")
			private static WebElement sbr5throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[5]/td[4]")
			private static WebElement sbr5throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[5]/td[5]")
			private static WebElement sbr5throw_5thColumn;
				
				
			// 6th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[6]/td[2]")
			private static WebElement sbr6throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[6]/td[3]")
			private static WebElement sbr6throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[6]/td[4]")
			private static WebElement sbr6throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[6]/td[5]")				
			private static WebElement sbr6throw_5thColumn;
				
			
			
			// 7th

			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[7]/td[2]")
			private static WebElement sbr7throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[7]/td[3]")
			private static WebElement sbr7throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[7]/td[4]")
			private static WebElement sbr7throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[7]/td[5]")
			private static WebElement sbr7throw_5thColumn;
				
				
			// 8th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[8]/td[2]")
			private static WebElement sbr8throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[8]/td[3]")
			private static WebElement sbr8throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[8]/td[4]")
			private static WebElement sbr8throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[8]/td[5]")
			private static WebElement sbr8throw_5thColumn;
				
				
			
			// 9th
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[9]/td[2]")
			private static WebElement sbr9throw_2ndColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[9]/td[3]")
			private static WebElement sbr9throw_3rdColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[9]/td[4]")
			private static WebElement sbr9throw_4thColumn;
			
			@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[9]/td[5]")
			private static WebElement sbr9throw_5thColumn;
				
				
			// 10th
			
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[10]/td[2]")
				private static WebElement sbr10throw_2ndColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[10]/td[3]")
				private static WebElement sbr10throw_3rdColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[10]/td[4]")
				private static WebElement sbr10throw_4thColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[10]/td[5]")
				private static WebElement sbr10throw_5thColumn;
				
			// 11th
			
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[11]/td[2]")
				private static WebElement sbr11throw_2ndColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[11]/td[3]")
				private static WebElement sbr11throw_3rdColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[11]/td[4]")
				private static WebElement sbr11throw_4thColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[11]/td[5]")
				private static WebElement sbr11throw_5thColumn;
				
			// 12th
			
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[12]/td[2]")
				private static WebElement sbr12throw_2ndColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[12]/td[3]")
				private static WebElement sbr12throw_3rdColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[12]/td[4]")
				private static WebElement sbr12throw_4thColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[12]/td[5]")
				private static WebElement sbr12throw_5thColumn;
				
			// 13th
			
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[13]/td[2]")
				private static WebElement sbr13throw_2ndColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[13]/td[3]")
				private static WebElement sbr13throw_3rdColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[13]/td[4]")
				private static WebElement sbr13throw_4thColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[13]/td[5]")
				private static WebElement sbr13throw_5thColumn;
			
			// 14th
			
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[14]/td[2]")
				private static WebElement sbr14throw_2ndColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[14]/td[3]")
				private static WebElement sbr14throw_3rdColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[14]/td[4]")
				private static WebElement sbr14throw_4thColumn;
				
				@FindBy(xpath="//*[@id='ratesWiseData']/tbody/tr[14]/td[5]")
				private static WebElement sbr14throw_5thColumn;
	       	
				@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[2]/a")
				private static WebElement iq_CloseBtn;
				
				@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[1]/a")
				private static WebElement iq_ClearBtn;

	public boolean checkItemQueryReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		// LogoutandLoginwithSU();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(itemQuery));
		itemQuery.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(iq_ItemTxt));
		iq_ItemTxt.click();
		iq_ItemTxt.sendKeys(excelReader.getCellData(xlSheetName, 8, 5));
		Thread.sleep(2000);
		iq_ItemTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(iq_getStockBtn));
		iq_getStockBtn.click();
		
		Thread.sleep(5000);
		
		String actiq_OpeningStocks				=iq_OpeningStocks.getText();
		String actiq_PendingPurchaseOrders		=iq_PendingPurchaseOrders.getText();
		String actiq_CurrentStock				=iq_CurrentStock.getText();
		String actiq_PendingSalesOrders			=iq_PendingSalesOrders.getText();
		String actiq_AvgStockRate				=iq_AvgStockRate.getText();
		String actiq_QtyToBeOrdered				=iq_QtyToBeOrdered.getText();
		String actiq_Value						=iq_Value.getText();
		String actiq_BaseUnit     				=iq_BaseUnit.getText();
		String actiq_AlternativeQuantity     	=iq_AlternateQuantity.getText();
		
		String expiq_OpeningStocks				=excelReader.getCellData(xlSheetName, 8, 6);
		String expiq_PendingPurchaseOrders		=excelReader.getCellData(xlSheetName, 9, 6);
		String expiq_CurrentStock				=excelReader.getCellData(xlSheetName, 10, 6);
		String expiq_PendingSalesOrders			=excelReader.getCellData(xlSheetName, 11, 6);
		String expiq_AvgStockRate				=excelReader.getCellData(xlSheetName, 12, 6);
		String expiq_QtyToBeOrdered				=excelReader.getCellData(xlSheetName, 13, 6);
		String expiq_Value						=excelReader.getCellData(xlSheetName, 14, 6);
		String expiq_BaseUnit     				=excelReader.getCellData(xlSheetName, 15, 6);
		//String expiq_AlternativeQuantity		="0.00";
		
		excelReader.setCellData(xlfile, xlSheetName, 8, 7, actiq_OpeningStocks);
		excelReader.setCellData(xlfile, xlSheetName, 9, 7, actiq_PendingPurchaseOrders);
		excelReader.setCellData(xlfile, xlSheetName, 10, 7, actiq_CurrentStock);
		excelReader.setCellData(xlfile, xlSheetName, 11, 7, actiq_PendingSalesOrders);
		excelReader.setCellData(xlfile, xlSheetName, 12, 7, actiq_AvgStockRate);
		excelReader.setCellData(xlfile, xlSheetName, 13, 7, actiq_QtyToBeOrdered);
		excelReader.setCellData(xlfile, xlSheetName, 14, 7, actiq_Value);
		excelReader.setCellData(xlfile, xlSheetName, 15, 7, actiq_BaseUnit);
		
		Thread.sleep(3000);
		// Recent Stock Transaction -Purchases
		
		String actpurVoucherNoR1	=pur1strow_2ndColumn.getText();
		String actpurVendorR1		=pur1strow_3rdColumn.getText();
		String actpurQtyR1			=pur1strow_4thColumn.getText();
		String actpurUnitR1			=pur1strow_5thColumn.getText();
		String actpurRateR1			=pur1strow_6thColumn.getText();
		String actpurCurrencyR1		=pur1strow_7thColumn.getText();
		String actpurQtyInBaseR1	=pur1strow_8thColumn.getText();
		
		String actpurVoucherNoR2	=pur2ndrow_2ndColumn.getText();
		String actpurVendorR2		=pur2ndrow_3rdColumn.getText();
		String actpurQtyR2			=pur2ndrow_4thColumn.getText();
		String actpurUnitR2			=pur2ndrow_5thColumn.getText();
		String actpurRateR2			=pur2ndrow_6thColumn.getText();
		String actpurCurrencyR2		=pur2ndrow_7thColumn.getText();
		String actpurQtyInBaseR2	=pur2ndrow_8thColumn.getText();
	
		String actpurTotalWa		=purTotalWa.getText();
		
		// Expected
		
		String exppurVoucherNoR1	=excelReader.getCellData(xlSheetName, 16, 6);
		String exppurVendorR1		=excelReader.getCellData(xlSheetName, 17, 6);
		String exppurQtyR1			=excelReader.getCellData(xlSheetName, 18, 6);
		String exppurUnitR1			=excelReader.getCellData(xlSheetName, 19, 6);
		String exppurRateR1			=excelReader.getCellData(xlSheetName, 20, 6);
		String exppurCurrencyR1		=excelReader.getCellData(xlSheetName, 21, 6);
		String exppurQtyInBaseR1	=excelReader.getCellData(xlSheetName, 22, 6);
		
		String exppurVoucherNoR2	=excelReader.getCellData(xlSheetName, 23, 6);
		String exppurVendorR2		=excelReader.getCellData(xlSheetName, 24, 6);
		String exppurQtyR2			=excelReader.getCellData(xlSheetName, 25, 6);
		String exppurUnitR2			=excelReader.getCellData(xlSheetName, 26, 6);
		String exppurRateR2			=excelReader.getCellData(xlSheetName, 27, 6);
		String exppurCurrencyR2		=excelReader.getCellData(xlSheetName, 28, 6);
		String exppurQtyInBaseR2	=excelReader.getCellData(xlSheetName, 29, 6);
		
		String exppurTotalWa		=excelReader.getCellData(xlSheetName, 30, 6);//NDT57:SU/IND/TEXT2
		
		excelReader.setCellData(xlfile, xlSheetName, 16, 7, actpurVoucherNoR1);
		excelReader.setCellData(xlfile, xlSheetName, 17, 7, actpurVendorR1);
		excelReader.setCellData(xlfile, xlSheetName, 18, 7, actpurQtyR1);
		excelReader.setCellData(xlfile, xlSheetName, 19, 7, actpurUnitR1);
		excelReader.setCellData(xlfile, xlSheetName, 20, 7, actpurRateR1);
		excelReader.setCellData(xlfile, xlSheetName, 21, 7, actpurCurrencyR1);
		excelReader.setCellData(xlfile, xlSheetName, 22, 7, actpurQtyInBaseR1);
		

		excelReader.setCellData(xlfile, xlSheetName, 23, 7, actpurVoucherNoR2);
		excelReader.setCellData(xlfile, xlSheetName, 24, 7, actpurVendorR2);
		excelReader.setCellData(xlfile, xlSheetName, 25, 7, actpurQtyR2);
		excelReader.setCellData(xlfile, xlSheetName, 26, 7, actpurUnitR2);
		excelReader.setCellData(xlfile, xlSheetName, 27, 7, actpurRateR2);
		excelReader.setCellData(xlfile, xlSheetName, 28, 7, actpurCurrencyR2);
		excelReader.setCellData(xlfile, xlSheetName, 29, 7, actpurQtyInBaseR2);

		excelReader.setCellData(xlfile, xlSheetName, 30, 7, actpurTotalWa);

		
		
		getAction().sendKeys(Keys.CONTROL,Keys.END).build().perform();
		
		
		// Recent Stock Transaction -Sales
		
		String actsalVoucherNoR1	=sal1strow_2ndColumn.getText();
		String actsalCustomerR1		=sal1strow_3rdColumn.getText();
		String actsalQtyR1			=sal1strow_4thColumn.getText();
		String actsalUnitR1			=sal1strow_5thColumn.getText();
		String actsalRateR1			=sal1strow_6thColumn.getText();
		String actsalCurrencyR1		=sal1strow_7thColumn.getText();
		String actsalQtyInBaseR1	=sal1strow_8thColumn.getText();
		
		String actsalVoucherNoR2	=sal2ndrow_2ndColumn.getText();
		String actsalCustomerR2		=sal2ndrow_3rdColumn.getText();
		String actsalQtyR2			=sal2ndrow_4thColumn.getText();
		String actsalUnitR2			=sal2ndrow_5thColumn.getText();
		String actsalRateR2			=sal2ndrow_6thColumn.getText();
		String actsalCurrencyR2		=sal2ndrow_7thColumn.getText();
		String actsalQtyInBaseR2	=sal2ndrow_8thColumn.getText();
		
		String actsalTotalWa		=salTotalWa.getText();

		// Expected
		String expsalVoucherNoR1	=excelReader.getCellData(xlSheetName, 31, 6);
		String expsalCustomerR1		=excelReader.getCellData(xlSheetName, 32, 6);
		String expsalQtyR1			=excelReader.getCellData(xlSheetName, 33, 6);
		String expsalUnitR1			=excelReader.getCellData(xlSheetName, 34, 6);
		String expsalRateR1			=excelReader.getCellData(xlSheetName, 35, 6);
		String expsalCurrencyR1		=excelReader.getCellData(xlSheetName, 36, 6);
		String expsalQtyInBaseR1	=excelReader.getCellData(xlSheetName, 37, 6);
		
		String expsalVoucherNoR2	=excelReader.getCellData(xlSheetName, 38, 6);
		String expsalCustomerR2 	=excelReader.getCellData(xlSheetName, 39, 6);
		String expsalQtyR2			=excelReader.getCellData(xlSheetName, 40, 6);
		String expsalUnitR2			=excelReader.getCellData(xlSheetName, 41, 6);
		String expsalRateR2			=excelReader.getCellData(xlSheetName, 42, 6);
		String expsalCurrencyR2		=excelReader.getCellData(xlSheetName, 43, 6);
		String expsalQtyInBaseR2	=excelReader.getCellData(xlSheetName, 44, 6);
		
		String expsalTotalWa		=excelReader.getCellData(xlSheetName, 45, 6);
		
		excelReader.setCellData(xlfile, xlSheetName, 31, 7, actsalVoucherNoR1);
		excelReader.setCellData(xlfile, xlSheetName, 32, 7, actsalCustomerR1);
		excelReader.setCellData(xlfile, xlSheetName, 33, 7, actsalQtyR1);
		excelReader.setCellData(xlfile, xlSheetName, 34, 7, actsalUnitR1);
		excelReader.setCellData(xlfile, xlSheetName, 35, 7, actsalRateR1);
		excelReader.setCellData(xlfile, xlSheetName, 36, 7, actsalCurrencyR1);
		excelReader.setCellData(xlfile, xlSheetName, 37, 7, actsalQtyInBaseR1);
		
		excelReader.setCellData(xlfile, xlSheetName, 38, 7, actsalVoucherNoR2);
		excelReader.setCellData(xlfile, xlSheetName, 39, 7, actsalCustomerR2);
		excelReader.setCellData(xlfile, xlSheetName, 40, 7, actsalQtyR2);
		excelReader.setCellData(xlfile, xlSheetName, 41, 7, actsalUnitR2);
		excelReader.setCellData(xlfile, xlSheetName, 42, 7, actsalRateR2);
		excelReader.setCellData(xlfile, xlSheetName, 43, 7, actsalCurrencyR2);
		excelReader.setCellData(xlfile, xlSheetName, 44, 7, actsalQtyInBaseR2);
		
		excelReader.setCellData(xlfile, xlSheetName, 45, 7, actsalTotalWa);

		
		// WareHouse Table
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BalancebyWareHouseExpandBtn));
		BalancebyWareHouseExpandBtn.click();
		
		Thread.sleep(3000);
		
		getAction().sendKeys(Keys.CONTROL,Keys.END).build().perform();
		
		
		String actwareHouseR1		=wh1strow_2ndColumn.getText();
		String actstockR1			=wh1strow_3rdColumn.getText();
		String actrateR1			=wh1strow_4thColumn.getText();
		String actvalueR1			=wh1strow_5thColumn.getText();
		
		String acttotalStock		=wh2ndrow_2ndColumn.getText();
		String acttotalValue		=wh2ndrow_4thColumn.getText();
		
		// Expected
		String expwareHouseR1		=excelReader.getCellData(xlSheetName, 46, 6);
		String expstockR1			=excelReader.getCellData(xlSheetName, 47, 6);
		String exprateR1			=excelReader.getCellData(xlSheetName, 48, 6);
		String expvalueR1			=excelReader.getCellData(xlSheetName, 49, 6);
		
		String exptotalStock		=excelReader.getCellData(xlSheetName, 50, 6);
		String exptotalValue		=excelReader.getCellData(xlSheetName, 51, 6);
		
		
		
		excelReader.setCellData(xlfile, xlSheetName, 46, 7, actwareHouseR1);
		excelReader.setCellData(xlfile, xlSheetName, 47, 7, actstockR1);
		excelReader.setCellData(xlfile, xlSheetName, 48, 7, actrateR1);
		excelReader.setCellData(xlfile, xlSheetName, 49, 7, actvalueR1);
		
		excelReader.setCellData(xlfile, xlSheetName, 50, 7, acttotalStock);
		excelReader.setCellData(xlfile, xlSheetName, 51, 7, acttotalValue);

		
		
		// Seller And Buying rates Table
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sellingAndBuyingRatesExpandBtn));
		sellingAndBuyingRatesExpandBtn.click();
		
		Thread.sleep(2000);
		
		getAction().sendKeys(Keys.CONTROL,Keys.END).build().perform();
		
		Thread.sleep(2000);
		
		String actsellingRateR1		=sbr1strow_2ndColumn.getText();
		String actsellingValueR1	=sbr1strow_3rdColumn.getText();
		String actbuyingRateR1		=sbr1strow_4thColumn.getText();
		String actbuyingValueR1		=sbr1strow_5thColumn.getText();
		
		String actsellingRateR2		=sbr2ndrow_2ndColumn.getText();
		String actsellingValueR2	=sbr2ndrow_3rdColumn.getText();
		String actbuyingRateR2		=sbr2ndrow_4thColumn.getText();
		String actbuyingValueR2		=sbr2ndrow_5thColumn.getText();
		
		String actsellingRateR3		=sbr3rdrow_2ndColumn.getText();
		String actsellingValueR3	=sbr3rdrow_3rdColumn.getText();
		String actbuyingRateR3		=sbr3rdrow_4thColumn.getText();
		String actbuyingValueR3		=sbr3rdrow_5thColumn.getText();
		
		String actsellingRateR4		=sbr4throw_2ndColumn.getText();
		String actsellingValueR4	=sbr4throw_3rdColumn.getText();
		String actbuyingRateR4		=sbr4throw_4thColumn.getText();
		String actbuyingValueR4		=sbr4throw_5thColumn.getText();
		
		String actsellingRateR5		=sbr5throw_2ndColumn.getText();
		String actsellingValueR5	=sbr5throw_3rdColumn.getText();
		String actbuyingRateR5		=sbr5throw_4thColumn.getText();
		String actbuyingValueR5		=sbr5throw_5thColumn.getText();
		
		String actsellingRateR6		=sbr6throw_2ndColumn.getText();
		String actsellingValueR6	=sbr6throw_3rdColumn.getText();
		String actbuyingRateR6		=sbr6throw_4thColumn.getText();
		String actbuyingValueR6		=sbr6throw_5thColumn.getText();
		
		String actsellingRateR7		=sbr7throw_2ndColumn.getText();
		String actsellingValueR7	=sbr7throw_3rdColumn.getText();
		String actbuyingRateR7		=sbr7throw_4thColumn.getText();
		String actbuyingValueR7		=sbr7throw_5thColumn.getText();
		
		String actsellingRateR8		=sbr8throw_2ndColumn.getText();
		String actsellingValueR8	=sbr8throw_3rdColumn.getText();
		String actbuyingRateR8		=sbr8throw_4thColumn.getText();
		String actbuyingValueR8		=sbr8throw_5thColumn.getText();
		
		String actsellingRateR9		=sbr9throw_2ndColumn.getText();
		String actsellingValueR9	=sbr9throw_3rdColumn.getText();
		String actbuyingRateR9		=sbr9throw_4thColumn.getText();
		String actbuyingValueR9		=sbr9throw_5thColumn.getText();
		
		String actsellingRateR10	=sbr10throw_2ndColumn.getText();
		String actsellingValueR10	=sbr10throw_3rdColumn.getText();
		String actbuyingRateR10		=sbr10throw_4thColumn.getText();
		String actbuyingValueR10	=sbr10throw_5thColumn.getText();
		
		String actsellingRateR11	=sbr11throw_2ndColumn.getText();
		String actsellingValueR11	=sbr11throw_3rdColumn.getText();
		String actbuyingRateR11		=sbr11throw_4thColumn.getText();
		String actbuyingValueR11	=sbr11throw_5thColumn.getText();
		
		String actsellingRateR12	=sbr12throw_2ndColumn.getText();
		String actsellingValueR12	=sbr12throw_3rdColumn.getText();
		String actbuyingRateR12		=sbr12throw_4thColumn.getText();
		String actbuyingValueR12	=sbr12throw_5thColumn.getText();
		
		String actsellingRateR13	=sbr13throw_2ndColumn.getText();
		String actsellingValueR13	=sbr13throw_3rdColumn.getText();
		String actbuyingRateR13		=sbr13throw_4thColumn.getText();
		String actbuyingValueR13	=sbr13throw_5thColumn.getText();
		
		String actsellingRateR14	=sbr14throw_2ndColumn.getText();
		String actsellingValueR14	=sbr14throw_3rdColumn.getText();
		String actbuyingRateR14		=sbr14throw_4thColumn.getText();
		String actbuyingValueR14	=sbr14throw_5thColumn.getText();
		
		
		// Expected
		
		String expsellingRateR1		=excelReader.getCellData(xlSheetName, 52, 6);
		String expsellingValueR1	=excelReader.getCellData(xlSheetName, 53, 6);
		String expbuyingRateR1		=excelReader.getCellData(xlSheetName, 54, 6);
		String expbuyingValueR1		=excelReader.getCellData(xlSheetName, 55, 6);

		String expsellingRateR2		=excelReader.getCellData(xlSheetName, 56, 6);
		String expsellingValueR2	=excelReader.getCellData(xlSheetName, 57, 6);
		String expbuyingRateR2		=excelReader.getCellData(xlSheetName, 58, 6);
		String expbuyingValueR2		=excelReader.getCellData(xlSheetName, 59, 6);
		
		String expsellingRateR3		=excelReader.getCellData(xlSheetName, 60, 6);
		String expsellingValueR3	=excelReader.getCellData(xlSheetName, 61, 6);
		String expbuyingRateR3		=excelReader.getCellData(xlSheetName, 62, 6);
		String expbuyingValueR3		=excelReader.getCellData(xlSheetName, 63, 6);
		
		String expsellingRateR4		=excelReader.getCellData(xlSheetName, 64, 6);
		String expsellingValueR4	=excelReader.getCellData(xlSheetName, 65, 6);
		String expbuyingRateR4		=excelReader.getCellData(xlSheetName, 66, 6);
		String expbuyingValueR4		=excelReader.getCellData(xlSheetName, 67, 6);
		
		String expsellingRateR5		=excelReader.getCellData(xlSheetName, 68, 6);
		String expsellingValueR5	=excelReader.getCellData(xlSheetName, 69, 6);
		String expbuyingRateR5		=excelReader.getCellData(xlSheetName, 70, 6);
		String expbuyingValueR5		=excelReader.getCellData(xlSheetName, 71, 6);
		
		String expsellingRateR6		=excelReader.getCellData(xlSheetName, 72, 6);
		String expsellingValueR6	=excelReader.getCellData(xlSheetName, 73, 6);
		String expbuyingRateR6		=excelReader.getCellData(xlSheetName, 74, 6);
		String expbuyingValueR6		=excelReader.getCellData(xlSheetName, 75, 6);
		
		String expsellingRateR7		=excelReader.getCellData(xlSheetName, 76, 6);
		String expsellingValueR7	=excelReader.getCellData(xlSheetName, 77, 6);
		String expbuyingRateR7		=excelReader.getCellData(xlSheetName, 78, 6);
		String expbuyingValueR7		=excelReader.getCellData(xlSheetName, 79, 6);
		
		String expsellingRateR8		=excelReader.getCellData(xlSheetName, 80, 6);
		String expsellingValueR8	=excelReader.getCellData(xlSheetName, 81, 6);
		String expbuyingRateR8		=excelReader.getCellData(xlSheetName, 82, 6);
		String expbuyingValueR8		=excelReader.getCellData(xlSheetName, 83, 6);

		String expsellingRateR9		=excelReader.getCellData(xlSheetName, 84, 6);
		String expsellingValueR9	=excelReader.getCellData(xlSheetName, 85, 6);
		String expbuyingRateR9		=excelReader.getCellData(xlSheetName, 86, 6);
		String expbuyingValueR9		=excelReader.getCellData(xlSheetName, 87, 6);
		
		String expsellingRateR10	=excelReader.getCellData(xlSheetName, 88, 6);
		String expsellingValueR10	=excelReader.getCellData(xlSheetName, 89, 6);
		String expbuyingRateR10		=excelReader.getCellData(xlSheetName, 90, 6);
		String expbuyingValueR10	=excelReader.getCellData(xlSheetName, 91, 6);
		
		String expsellingRateR11	=excelReader.getCellData(xlSheetName, 92, 6);
		String expsellingValueR11	=excelReader.getCellData(xlSheetName, 93, 6);
		String expbuyingRateR11		=excelReader.getCellData(xlSheetName, 94, 6);
		String expbuyingValueR11	=excelReader.getCellData(xlSheetName, 95, 6);
		
		String expsellingRateR12	=excelReader.getCellData(xlSheetName, 96, 6);
		String expsellingValueR12	=excelReader.getCellData(xlSheetName, 97, 6);
		String expbuyingRateR12		=excelReader.getCellData(xlSheetName, 98, 6);
		String expbuyingValueR12	=excelReader.getCellData(xlSheetName, 99, 6);
		
		String expsellingRateR13	=excelReader.getCellData(xlSheetName, 100, 6);
		String expsellingValueR13	=excelReader.getCellData(xlSheetName, 101, 6);
		String expbuyingRateR13		=excelReader.getCellData(xlSheetName, 102, 6);
		String expbuyingValueR13	=excelReader.getCellData(xlSheetName, 103, 6);
		
		String expsellingRateR14	=excelReader.getCellData(xlSheetName, 104, 6);
		String expsellingValueR14	=excelReader.getCellData(xlSheetName, 105, 6);
		String expbuyingRateR14		=excelReader.getCellData(xlSheetName, 106, 6);
		String expbuyingValueR14	=excelReader.getCellData(xlSheetName, 107, 6);
		
		
		excelReader.setCellData(xlfile, xlSheetName, 52, 7, actsellingRateR1);
		excelReader.setCellData(xlfile, xlSheetName, 53, 7, actsellingValueR1);
		excelReader.setCellData(xlfile, xlSheetName, 54, 7, actbuyingRateR1);
		excelReader.setCellData(xlfile, xlSheetName, 55, 7, actbuyingValueR1);
		
		
		excelReader.setCellData(xlfile, xlSheetName, 56, 7, actsellingRateR2);
		excelReader.setCellData(xlfile, xlSheetName, 57, 7, actsellingValueR2);
		excelReader.setCellData(xlfile, xlSheetName, 58, 7, actbuyingRateR2);
		excelReader.setCellData(xlfile, xlSheetName, 59, 7, actbuyingValueR2);
		
		excelReader.setCellData(xlfile, xlSheetName, 60, 7, actsellingRateR3);
		excelReader.setCellData(xlfile, xlSheetName, 61, 7, actsellingValueR3);
		excelReader.setCellData(xlfile, xlSheetName, 62, 7, actbuyingRateR3);
		excelReader.setCellData(xlfile, xlSheetName, 63, 7, actbuyingValueR3);
		
		excelReader.setCellData(xlfile, xlSheetName, 64, 7, actsellingRateR4);
		excelReader.setCellData(xlfile, xlSheetName, 65, 7, actsellingValueR4);
		excelReader.setCellData(xlfile, xlSheetName, 66, 7, actbuyingRateR4);
		excelReader.setCellData(xlfile, xlSheetName, 67, 7, actbuyingValueR4);
		
		excelReader.setCellData(xlfile, xlSheetName, 68, 7, actsellingRateR5);
		excelReader.setCellData(xlfile, xlSheetName, 69, 7, actsellingValueR5);
		excelReader.setCellData(xlfile, xlSheetName, 70, 7, actbuyingRateR5);
		excelReader.setCellData(xlfile, xlSheetName, 71, 7, actbuyingValueR5);
		
		excelReader.setCellData(xlfile, xlSheetName, 72, 7, actsellingRateR6);
		excelReader.setCellData(xlfile, xlSheetName, 73, 7, actsellingValueR6);
		excelReader.setCellData(xlfile, xlSheetName, 74, 7, actbuyingRateR6);
		excelReader.setCellData(xlfile, xlSheetName, 75, 7, actbuyingValueR6);
		
		excelReader.setCellData(xlfile, xlSheetName, 76, 7, actsellingRateR7);
		excelReader.setCellData(xlfile, xlSheetName, 77, 7, actsellingValueR7);
		excelReader.setCellData(xlfile, xlSheetName, 78, 7, actbuyingRateR7);
		excelReader.setCellData(xlfile, xlSheetName, 79, 7, actbuyingValueR7);
		
		excelReader.setCellData(xlfile, xlSheetName, 80, 7, actsellingRateR8);
		excelReader.setCellData(xlfile, xlSheetName, 81, 7, actsellingValueR8);
		excelReader.setCellData(xlfile, xlSheetName, 82, 7, actbuyingRateR8);
		excelReader.setCellData(xlfile, xlSheetName, 83, 7, actbuyingValueR8);
		
		excelReader.setCellData(xlfile, xlSheetName, 84, 7, actsellingRateR9);
		excelReader.setCellData(xlfile, xlSheetName, 85, 7, actsellingValueR9);
		excelReader.setCellData(xlfile, xlSheetName, 86, 7, actbuyingRateR9);
		excelReader.setCellData(xlfile, xlSheetName, 87, 7, actbuyingValueR9);
		
		excelReader.setCellData(xlfile, xlSheetName, 88, 7, actsellingRateR10);
		excelReader.setCellData(xlfile, xlSheetName, 89, 7, actsellingValueR10);
		excelReader.setCellData(xlfile, xlSheetName, 90, 7, actbuyingRateR10);
		excelReader.setCellData(xlfile, xlSheetName, 91, 7, actbuyingValueR10);
		
		excelReader.setCellData(xlfile, xlSheetName, 92, 7, actsellingRateR11);
		excelReader.setCellData(xlfile, xlSheetName, 93, 7, actsellingValueR11);
		excelReader.setCellData(xlfile, xlSheetName, 94, 7, actbuyingRateR11);
		excelReader.setCellData(xlfile, xlSheetName, 95, 7, actbuyingValueR11);
		
		
		excelReader.setCellData(xlfile, xlSheetName, 96, 7, actsellingRateR12);
		excelReader.setCellData(xlfile, xlSheetName, 97, 7, actsellingValueR12);
		excelReader.setCellData(xlfile, xlSheetName, 98, 7, actbuyingRateR12);
		excelReader.setCellData(xlfile, xlSheetName, 99, 7, actbuyingValueR12);
		
		
		excelReader.setCellData(xlfile, xlSheetName, 100, 7, actsellingRateR13);
		excelReader.setCellData(xlfile, xlSheetName, 101, 7, actsellingValueR13);
		excelReader.setCellData(xlfile, xlSheetName, 102, 7, actbuyingRateR13);
		excelReader.setCellData(xlfile, xlSheetName, 103, 7, actbuyingValueR13);
		
		
		
		excelReader.setCellData(xlfile, xlSheetName, 104, 7, actsellingRateR14);
		excelReader.setCellData(xlfile, xlSheetName, 105, 7, actsellingValueR14);
		excelReader.setCellData(xlfile, xlSheetName, 106, 7, actbuyingRateR14);
		excelReader.setCellData(xlfile, xlSheetName, 107, 7, actbuyingValueR14);
		


		
		
		
		
		System.out.println("************************************checkItemQueryOfBatchWAItem*************************************");
		
		System.out.println("iq_OpeningStocks     		 :   "+actiq_OpeningStocks					+" Value Expected : "+expiq_OpeningStocks);
		System.out.println("iq_PendingPurchaseOrders     :   "+actiq_PendingPurchaseOrders			+" Value Expected : "+expiq_PendingPurchaseOrders);
		System.out.println("iq_CurrentStock     		 :   "+actiq_CurrentStock					+" Value Expected : "+expiq_CurrentStock);
		System.out.println("iq_PendingSalesOrders     	 :   "+actiq_PendingSalesOrders				+" Value Expected : "+expiq_PendingSalesOrders);
		System.out.println("iq_AvgStockRate     		 :   "+actiq_AvgStockRate					+" Value Expected : "+expiq_AvgStockRate);
		System.out.println("iq_QtyToBeOrdered     		 :   "+actiq_QtyToBeOrdered					+" Value Expected : "+expiq_QtyToBeOrdered);
		System.out.println("iq_Value     				 :   "+actiq_Value							+" Value Expected : "+expiq_Value);
		System.out.println("iq_BaseUnit     			 :   "+actiq_BaseUnit						+" Value Expected : "+expiq_BaseUnit);
		
		System.out.println("********************* Purchase ******************");
		
		System.out.println("purVoucherNoR1     			 :   "+actpurVoucherNoR1					+" Value Expected : "+exppurVoucherNoR1);
		System.out.println("purVendorR1     			 :   "+actpurVendorR1						+" Value Expected : "+exppurVendorR1);
		System.out.println("purQtyR1     				 :   "+actpurQtyR1							+" Value Expected : "+exppurQtyR1);
		System.out.println("purUnitR1     				 :   "+actpurUnitR1							+" Value Expected : "+exppurUnitR1);
		System.out.println("purRateR1     				 :   "+actpurRateR1							+" Value Expected : "+exppurRateR1);
		System.out.println("purCurrencyR1     			 :   "+actpurCurrencyR1						+" Value Expected : "+exppurCurrencyR1);
		System.out.println("purQtyInBaseR1     			 :   "+actpurQtyInBaseR1					+" Value Expected : "+exppurQtyInBaseR1);
		
		System.out.println("purVoucherNoR2     			 :   "+actpurVoucherNoR2					+" Value Expected : "+exppurVoucherNoR2);
		System.out.println("purVendorR2     			 :   "+actpurVendorR2						+" Value Expected : "+exppurVendorR2);
		System.out.println("purQtyR2     				 :   "+actpurQtyR2							+" Value Expected : "+exppurQtyR2);
		System.out.println("purUnitR2     				 :   "+actpurUnitR2							+" Value Expected : "+exppurUnitR2);
		System.out.println("purRateR2     				 :   "+actpurRateR2							+" Value Expected : "+exppurRateR2);
		System.out.println("purCurrencyR2     			 :   "+actpurCurrencyR2						+" Value Expected : "+exppurCurrencyR2);
		System.out.println("purQtyInBaseR2     			 :   "+actpurQtyInBaseR2					+" Value Expected : "+exppurQtyInBaseR2);
		
		System.out.println("purTotalWa                   :   "+actpurTotalWa                        +" Value expected : "+exppurTotalWa);
		
		
		
		System.out.println("********************* Sales ******************");
		
		System.out.println("salVoucherNoR1     			 :   "+actsalVoucherNoR1					+" Value Expected : "+expsalVoucherNoR1);
		System.out.println("salCustomerR1     			 :   "+actsalCustomerR1						+" Value Expected : "+expsalCustomerR1);
		System.out.println("salQtyR1     				 :   "+actsalQtyR1							+" Value Expected : "+expsalQtyR1);
		System.out.println("salUnitR1     				 :   "+actsalUnitR1							+" Value Expected : "+expsalUnitR1);
		System.out.println("salRateR1     				 :   "+actsalRateR1							+" Value Expected : "+expsalRateR1);
		System.out.println("salCurrencyR1     			 :   "+actsalCurrencyR1						+" Value Expected : "+expsalCurrencyR1);
		System.out.println("salQtyInBaseR1     			 :   "+actsalQtyInBaseR1					+" Value Expected : "+expsalQtyInBaseR1);
		
		System.out.println("salVoucherNoR2     			 :   "+actsalVoucherNoR2					+" Value Expected : "+expsalVoucherNoR2);
		System.out.println("salCustomerR2     			 :   "+actsalCustomerR2						+" Value Expected : "+expsalCustomerR2);
		System.out.println("salQtyR2     				 :   "+actsalQtyR2							+" Value Expected : "+expsalQtyR2);
		System.out.println("salUnitR2     				 :   "+actsalUnitR2							+" Value Expected : "+expsalUnitR2);
		System.out.println("salRateR2     				 :   "+actsalRateR2							+" Value Expected : "+expsalRateR2);
		System.out.println("salCurrencyR2     			 :   "+actsalCurrencyR2						+" Value Expected : "+expsalCurrencyR2);
		System.out.println("salQtyInBaseR2     			 :   "+actsalQtyInBaseR2					+" Value Expected : "+expsalQtyInBaseR2);
		
		System.out.println("salTotalWa                   :   "+actsalTotalWa                        +" Value expected : "+expsalTotalWa);
		
		
		System.out.println("********************* WareHouse ******************");
		
		System.out.println("wareHouseR1     			 :   "+actwareHouseR1						+" Value Expected : "+expwareHouseR1);
		System.out.println("stockR1     				 :   "+actstockR1							+" Value Expected : "+expstockR1);
		System.out.println("rateR1     				 	 :   "+actrateR1							+" Value Expected : "+exprateR1);
		System.out.println("valueR1     				 :   "+actvalueR1							+" Value Expected : "+expvalueR1);
	/*	
		System.out.println("wareHouseR2     			 :   "+actwareHouseR2						+" Value Expected : "+expwareHouseR2);
		System.out.println("stockR2     				 :   "+actstockR2							+" Value Expected : "+expstockR2);
		System.out.println("rateR2     				 	 :   "+actrateR2							+" Value Expected : "+exprateR2);
		System.out.println("valueR2     				 :   "+actvalueR2							+" Value Expected : "+expvalueR2);
	*/	
		System.out.println("totalStock     				 :   "+acttotalStock						+" Value Expected : "+exptotalStock);
		System.out.println("totalValue     				 :   "+acttotalValue						+" Value Expected : "+exptotalValue);
		
		
		System.out.println("*************************Selling and Buying rates********************");
		
		System.out.println("sellingRateR1     			 :   "+actsellingRateR1						+" Value Expected : "+expsellingRateR1);
		System.out.println("sellingValueR1  			 :   "+actsellingValueR1					+" Value Expected : "+expsellingValueR1);
		System.out.println("buyingRateR1     			 :   "+actbuyingRateR1						+" Value Expected : "+expbuyingRateR1);
		System.out.println("buyingValueR1     			 :   "+actbuyingValueR1						+" Value Expected : "+expbuyingValueR1);
		
		System.out.println("sellingRateR2     			 :   "+actsellingRateR2						+" Value Expected : "+expsellingRateR2);
		System.out.println("sellingValueR2  			 :   "+actsellingValueR2					+" Value Expected : "+expsellingValueR2);
		System.out.println("buyingRateR2     			 :   "+actbuyingRateR2						+" Value Expected : "+expbuyingRateR2);
		System.out.println("buyingValueR2     			 :   "+actbuyingValueR2						+" Value Expected : "+expbuyingValueR2);
		
		System.out.println("sellingRateR3     			 :   "+actsellingRateR3						+" Value Expected : "+expsellingRateR3);
		System.out.println("sellingValueR3  			 :   "+actsellingValueR3					+" Value Expected : "+expsellingValueR3);
		System.out.println("buyingRateR3     			 :   "+actbuyingRateR3						+" Value Expected : "+expbuyingRateR3);
		System.out.println("buyingValueR3     			 :   "+actbuyingValueR3						+" Value Expected : "+expbuyingValueR3);
		
		System.out.println("sellingRateR4     			 :   "+actsellingRateR4						+" Value Expected : "+expsellingRateR4);
		System.out.println("sellingValueR4  			 :   "+actsellingValueR4					+" Value Expected : "+expsellingValueR4);
		System.out.println("buyingRateR4     			 :   "+actbuyingRateR4						+" Value Expected : "+expbuyingRateR4);
		System.out.println("buyingValueR4     			 :   "+actbuyingValueR4						+" Value Expected : "+expbuyingValueR4);
		
		System.out.println("sellingRateR5     			 :   "+actsellingRateR5						+" Value Expected : "+expsellingRateR5);
		System.out.println("sellingValueR5  			 :   "+actsellingValueR5					+" Value Expected : "+expsellingValueR5);
		System.out.println("buyingRateR5     			 :   "+actbuyingRateR5						+" Value Expected : "+expbuyingRateR5);
		System.out.println("buyingValueR5     			 :   "+actbuyingValueR5						+" Value Expected : "+expbuyingValueR5);
		
		System.out.println("sellingRateR6     			 :   "+actsellingRateR6						+" Value Expected : "+expsellingRateR6);
		System.out.println("sellingValueR6  			 :   "+actsellingValueR6					+" Value Expected : "+expsellingValueR6);
		System.out.println("buyingRateR6     			 :   "+actbuyingRateR6						+" Value Expected : "+expbuyingRateR6);
		System.out.println("buyingValueR6     			 :   "+actbuyingValueR6						+" Value Expected : "+expbuyingValueR6);
		
		System.out.println("sellingRateR7     			 :   "+actsellingRateR7						+" Value Expected : "+expsellingRateR7);
		System.out.println("sellingValueR7  			 :   "+actsellingValueR7					+" Value Expected : "+expsellingValueR7);
		System.out.println("buyingRateR7     			 :   "+actbuyingRateR7						+" Value Expected : "+expbuyingRateR7);
		System.out.println("buyingValueR7     			 :   "+actbuyingValueR7						+" Value Expected : "+expbuyingValueR7);
		
		System.out.println("sellingRateR8     			 :   "+actsellingRateR8						+" Value Expected : "+expsellingRateR8);
		System.out.println("sellingValueR8  			 :   "+actsellingValueR8					+" Value Expected : "+expsellingValueR8);
		System.out.println("buyingRateR8     			 :   "+actbuyingRateR8						+" Value Expected : "+expbuyingRateR8);
		System.out.println("buyingValueR8     			 :   "+actbuyingValueR8						+" Value Expected : "+expbuyingValueR8);
		
		System.out.println("sellingRateR9     			 :   "+actsellingRateR9						+" Value Expected : "+expsellingRateR9);
		System.out.println("sellingValueR9  			 :   "+actsellingValueR9					+" Value Expected : "+expsellingValueR9);
		System.out.println("buyingRateR9     			 :   "+actbuyingRateR9						+" Value Expected : "+expbuyingRateR9);
		System.out.println("buyingValueR9     			 :   "+actbuyingValueR9						+" Value Expected : "+expbuyingValueR9);
		
		System.out.println("sellingRateR10     			 :   "+actsellingRateR10					+" Value Expected : "+expsellingRateR10);
		System.out.println("sellingValueR10  			 :   "+actsellingValueR10					+" Value Expected : "+expsellingValueR10);
		System.out.println("buyingRateR10     			 :   "+actbuyingRateR10						+" Value Expected : "+expbuyingRateR10);
		System.out.println("buyingValueR10     			 :   "+actbuyingValueR10					+" Value Expected : "+expbuyingValueR10);
		
		System.out.println("sellingRateR11     			 :   "+actsellingRateR11					+" Value Expected : "+expsellingRateR11);
		System.out.println("sellingValueR11  			 :   "+actsellingValueR11					+" Value Expected : "+expsellingValueR11);
		System.out.println("buyingRateR11     			 :   "+actbuyingRateR11						+" Value Expected : "+expbuyingRateR11);
		System.out.println("buyingValueR11     			 :   "+actbuyingValueR11					+" Value Expected : "+expbuyingValueR11);
		
		System.out.println("sellingRateR12     			 :   "+actsellingRateR12					+" Value Expected : "+expsellingRateR12);
		System.out.println("sellingValueR12  			 :   "+actsellingValueR12					+" Value Expected : "+expsellingValueR12);
		System.out.println("buyingRateR12     			 :   "+actbuyingRateR12						+" Value Expected : "+expbuyingRateR12);
		System.out.println("buyingValueR12     			 :   "+actbuyingValueR12					+" Value Expected : "+expbuyingValueR12);
		
		System.out.println("sellingRateR13     			 :   "+actsellingRateR13					+" Value Expected : "+expsellingRateR13);
		System.out.println("sellingValueR13  			 :   "+actsellingValueR13					+" Value Expected : "+expsellingValueR13);
		System.out.println("buyingRateR13     			 :   "+actbuyingRateR13						+" Value Expected : "+expbuyingRateR13);
		System.out.println("buyingValueR13     			 :   "+actbuyingValueR13					+" Value Expected : "+expbuyingValueR13);
		
		System.out.println("sellingRateR14     			 :   "+actsellingRateR14					+" Value Expected : "+expsellingRateR14);
		System.out.println("sellingValueR14  			 :   "+actsellingValueR14					+" Value Expected : "+expsellingValueR14);
		System.out.println("buyingRateR14     			 :   "+actbuyingRateR14						+" Value Expected : "+expbuyingRateR14);
		System.out.println("buyingValueR14     			 :   "+actbuyingValueR14					+" Value Expected : "+expbuyingValueR14);
		
		if(actiq_OpeningStocks.equalsIgnoreCase(expiq_OpeningStocks) && actiq_PendingPurchaseOrders.equalsIgnoreCase(expiq_PendingPurchaseOrders)
				&& actiq_CurrentStock.equalsIgnoreCase(expiq_CurrentStock) && actiq_PendingSalesOrders.equalsIgnoreCase(expiq_PendingSalesOrders)
				&& actiq_AvgStockRate.equalsIgnoreCase(expiq_AvgStockRate) && actiq_QtyToBeOrdered.equalsIgnoreCase(expiq_QtyToBeOrdered)
				&& actiq_Value.equalsIgnoreCase(expiq_Value) && actiq_BaseUnit.equalsIgnoreCase(expiq_BaseUnit)
				
				&& actpurVoucherNoR1.equalsIgnoreCase(exppurVoucherNoR1) && actpurVendorR1.equalsIgnoreCase(exppurVendorR1) && actpurQtyR1.equalsIgnoreCase(exppurQtyR1)
				&& actpurUnitR1.equalsIgnoreCase(exppurUnitR1) && actpurRateR1.equalsIgnoreCase(exppurRateR1) && actpurCurrencyR1.equalsIgnoreCase(exppurCurrencyR1)
				&& actpurQtyInBaseR1.equalsIgnoreCase(exppurQtyInBaseR1)
				
				&& actpurVoucherNoR2.equalsIgnoreCase(exppurVoucherNoR2) && actpurVendorR2.equalsIgnoreCase(exppurVendorR2) && actpurQtyR2.equalsIgnoreCase(exppurQtyR2)
				&& actpurUnitR2.equalsIgnoreCase(exppurUnitR2) && actpurRateR2.equalsIgnoreCase(exppurRateR2) && actpurCurrencyR2.equalsIgnoreCase(exppurCurrencyR2)
				&& actpurQtyInBaseR2.equalsIgnoreCase(exppurQtyInBaseR2)
			
				&& actpurTotalWa.equalsIgnoreCase(exppurTotalWa)
				
				&& actsalVoucherNoR1.equalsIgnoreCase(expsalVoucherNoR1) && actsalCustomerR1.equalsIgnoreCase(expsalCustomerR1) && actsalQtyR1.equalsIgnoreCase(expsalQtyR1)
				&& actsalUnitR1.equalsIgnoreCase(expsalUnitR1) && actsalRateR1.equalsIgnoreCase(expsalRateR1) && actsalCurrencyR1.equalsIgnoreCase(expsalCurrencyR1)
				&& actsalQtyInBaseR1.equalsIgnoreCase(expsalQtyInBaseR1)
				
				&& actsalVoucherNoR2.equalsIgnoreCase(expsalVoucherNoR2) && actsalCustomerR2.equalsIgnoreCase(expsalCustomerR2) && actsalQtyR2.equalsIgnoreCase(expsalQtyR2)
				&& actsalUnitR2.equalsIgnoreCase(expsalUnitR2) && actsalRateR2.equalsIgnoreCase(expsalRateR2) && actsalCurrencyR2.equalsIgnoreCase(expsalCurrencyR2)
				&& actsalQtyInBaseR2.equalsIgnoreCase(expsalQtyInBaseR2)
				
				&& actsalTotalWa.equalsIgnoreCase(expsalTotalWa)
				
				&& actwareHouseR1.equalsIgnoreCase(expwareHouseR1) && actstockR1.equalsIgnoreCase(expstockR1) 
				&& actrateR1.equalsIgnoreCase(exprateR1) && actvalueR1.equalsIgnoreCase(expvalueR1)
				
			/*	&& actwareHouseR2.equalsIgnoreCase(expwareHouseR2) && actstockR2.equalsIgnoreCase(expstockR2) 
				&& actrateR2.equalsIgnoreCase(exprateR2) && actvalueR2.equalsIgnoreCase(expvalueR2)*/
				
				&& acttotalStock.equalsIgnoreCase(exptotalStock) && acttotalValue.equalsIgnoreCase(exptotalValue)
				
				&& actsellingRateR1.equalsIgnoreCase(expsellingRateR1) && actsellingValueR1.equalsIgnoreCase(expsellingValueR1)
				&& actbuyingRateR1.equalsIgnoreCase(expbuyingRateR1) && actbuyingValueR1.equalsIgnoreCase(expbuyingValueR1)
				
				&& actsellingRateR2.equalsIgnoreCase(expsellingRateR2) && actsellingValueR2.equalsIgnoreCase(expsellingValueR2)
				&& actbuyingRateR2.equalsIgnoreCase(expbuyingRateR2) && actbuyingValueR2.equalsIgnoreCase(expbuyingValueR2)
				
				&& actsellingRateR3.equalsIgnoreCase(expsellingRateR3) && actsellingValueR3.equalsIgnoreCase(expsellingValueR3)
				&& actbuyingRateR3.equalsIgnoreCase(expbuyingRateR3) && actbuyingValueR3.equalsIgnoreCase(expbuyingValueR3)
				
				&& actsellingRateR4.equalsIgnoreCase(expsellingRateR4) && actsellingValueR4.equalsIgnoreCase(expsellingValueR4)
				&& actbuyingRateR4.equalsIgnoreCase(expbuyingRateR4) && actbuyingValueR4.equalsIgnoreCase(expbuyingValueR4)
				
				&& actsellingRateR5.equalsIgnoreCase(expsellingRateR5) && actsellingValueR4.equalsIgnoreCase(expsellingValueR5)
				&& actbuyingRateR5.equalsIgnoreCase(expbuyingRateR5) && actbuyingValueR5.equalsIgnoreCase(expbuyingValueR5)
				
				&& actsellingRateR6.equalsIgnoreCase(expsellingRateR6) && actsellingValueR6.equalsIgnoreCase(expsellingValueR6)
				&& actbuyingRateR6.equalsIgnoreCase(expbuyingRateR6) && actbuyingValueR6.equalsIgnoreCase(expbuyingValueR6)
				
				&& actsellingRateR7.equalsIgnoreCase(expsellingRateR7) && actsellingValueR7.equalsIgnoreCase(expsellingValueR7)
				&& actbuyingRateR7.equalsIgnoreCase(expbuyingRateR7) && actbuyingValueR7.equalsIgnoreCase(expbuyingValueR7)
				
				&& actsellingRateR8.equalsIgnoreCase(expsellingRateR8) && actsellingValueR8.equalsIgnoreCase(expsellingValueR8)
				&& actbuyingRateR8.equalsIgnoreCase(expbuyingRateR8) && actbuyingValueR8.equalsIgnoreCase(expbuyingValueR8)
				
				&& actsellingRateR9.equalsIgnoreCase(expsellingRateR9) && actsellingValueR9.equalsIgnoreCase(expsellingValueR9)
				&& actbuyingRateR9.equalsIgnoreCase(expbuyingRateR9) && actbuyingValueR9.equalsIgnoreCase(expbuyingValueR9)
				
				&& actsellingRateR10.equalsIgnoreCase(expsellingRateR10) && actsellingValueR10.equalsIgnoreCase(expsellingValueR10)
				&& actbuyingRateR10.equalsIgnoreCase(expbuyingRateR10) && actbuyingValueR10.equalsIgnoreCase(expbuyingValueR10)
				
				&& actsellingRateR11.equalsIgnoreCase(expsellingRateR11) && actsellingValueR11.equalsIgnoreCase(expsellingValueR11)
				&& actbuyingRateR11.equalsIgnoreCase(expbuyingRateR11) && actbuyingValueR11.equalsIgnoreCase(expbuyingValueR11)
				
				&& actsellingRateR12.equalsIgnoreCase(expsellingRateR12) && actsellingValueR12.equalsIgnoreCase(expsellingValueR12)
				&& actbuyingRateR12.equalsIgnoreCase(expbuyingRateR12) && actbuyingValueR12.equalsIgnoreCase(expbuyingValueR12)
				
				&& actsellingRateR13.equalsIgnoreCase(expsellingRateR13) && actsellingValueR13.equalsIgnoreCase(expsellingValueR13)
				&& actbuyingRateR13.equalsIgnoreCase(expbuyingRateR13) && actbuyingValueR13.equalsIgnoreCase(expbuyingValueR13)
				
				&& actsellingRateR14.equalsIgnoreCase(expsellingRateR14) && actsellingValueR14.equalsIgnoreCase(expsellingValueR14)
				&& actbuyingRateR14.equalsIgnoreCase(expbuyingRateR14) && actbuyingValueR14.equalsIgnoreCase(expbuyingValueR14))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(iq_CloseBtn));
			iq_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(iq_CloseBtn));
			iq_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 8, 8, resFail);

			return false;
		}
	}
	

	// Stock Ledger Elements 
	
			@FindBy(xpath="//*[@id='551']")
			private static WebElement stockLedger;
			
			@FindBy(xpath="//*[@id='dvReportInputs']/div[1]/div/label/span")
			private static WebElement sl_SelectAllChkBox;
			
			@FindBy(xpath="//div[@id='spnhiererchy']//select[@id='iTreeId']")
			private static WebElement sl_Dropdown1;
			
			@FindBy(xpath="//div[@id='spnhiererchy']//select[@id='iTreeViewId']")
			private static WebElement sl_Dropdown2;
			
			@FindBy(xpath="//label[@id='oncheckaccorderan']")
			private static WebElement sl_oncheckBtn;
			
			@FindBy(xpath="//input[@id='txtsrch-term']")
			private static WebElement sl_SearchItemTxt;
			
			@FindBy(xpath="//i[contains(@class,'icon-search theme_color-inverse')]")
			private static WebElement sl_SearchBtn;
			
			@FindBy(xpath="//input[@id='cmbUserTypeMaster']")
			private static WebElement sl_MasterTypeTxt;
			
			@FindBy(xpath="//span[contains(@class,'icon-font6 icon-external-module theme_icon-color')]")
			private static WebElement sl_AutoAdjustColumnsBtn;
			
			/*@FindBy(xpath="//a[contains(@class,'clsMasterName')][contains(text(),'ITEMS GROUP')]")
			private static WebElement sl_SearchItemTxt;*/
			
			@FindBy(xpath="//a[contains(@class,'clsMasterName')][contains(text(),'ITEMS GROUP')]")
			private static WebElement sl_ItemGroupsBtn;
			
			//@FindBy(xpath="//input[@id='liSelectAllMasters']")
			//private static WebElement sl_HeaderSelectChkBox;
			
			@FindBy(xpath="//*[@id='dvReportInputs']/div[1]/div/label/span")
			private static WebElement sl_HeaderSelectChkBox;
			
			@FindBy(xpath="//i[contains(@class,'icon-font6 icon-schedule')]")
			private static WebElement sl_ScheduleBtn;
			
			@FindBy(xpath="//i[contains(@class,'icon-font6 icon-custamize')]")
			private static WebElement sl_CustomizeBtn;
			
			@FindBy(xpath="//*[@id='id_reportmenudisplay']/ul/li[9]")
			private static WebElement sl_FilterExpandBtn;
			
			@FindBy(xpath="//*[@id='id_reportmenudisplay']/ul/li[9]/ul//li[1]")
			private static WebElement sl_FilterBtn;
			
			@FindBy(xpath="//*[@id='id_reportmenudisplay']/ul/li[9]/ul//li[2]")
			private static WebElement sl_RefineBtn;
			
			@FindBy(xpath="//*[@id='id_mainreportmenuheadings']/ul/li[9]")
			private static WebElement sl_OkBtn;
			
			@FindBy(xpath="//*[@id='id_mainreportmenuheadings']/ul/li[10]")
			private static WebElement sl_CloseBtn;
			
			@FindBy(xpath="//select[@id='DateOptions_']")
			private static WebElement sl_DateOptionDropdown;
			
			@FindBy(xpath="//input[@id='id_starting_date_']")
			private static WebElement sl_StartDateTxt;
			
			@FindBy(xpath="//td[@id='id_starting_date__input_image']//span[contains(@class,'icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow')]")
			private static WebElement sl_StartDateCalanderBtn;
			
			@FindBy(xpath="//input[@id='id_ending_date_']")
			private static WebElement sl_EndDateTxt;
			
			@FindBy(xpath="//td[@id='id_ending_date__input_image']//span[contains(@class,'icon-calender theme_color-inverse datecontrol_arrow_margin datecontrol_arrow')]")
			private static WebElement sl_EndDateCalanderBtn;
			
			@FindBy(xpath="//*[@id='MasterSingle__101']")
			private static WebElement sl_WarehouseTxt;
			
			@FindBy(xpath="//*[@id='RITCheckbox__6']")
			private static WebElement sl_IgnoreInternalTransferForAgeingChkbox;
			
			@FindBy(xpath="//*[@id='RITCombobox__2']")
			private static WebElement sl_ReportingLevelDropdown;
			
			@FindBy(xpath="//*[@id='RITCheckbox__4']")
			private static WebElement sl_DisplayLinearReportChkbox;
			
			@FindBy(xpath="//*[@id='RITCombobox__1']")
			private static WebElement sl_StockValuationDropdown;
			
			//@FindBy(xpath="//input[@id='RITCheckbox__3']")
			@FindBy(xpath="(//span[@class='checkmark'])[3]")
			private static WebElement sl_IncludeServiceTypeItemChkBox;
			
			@FindBy(xpath="(//span[@class='checkmark'])[2]")
			private static WebElement sl_IncludeServiceTypeProductItemChkBox;
			
			@FindBy(xpath="(//span[@class='checkmark'])[4]")
			
			private static WebElement stockLedger_IncludeServiceTypeItemChkBox;

			
			//@FindBy(xpath="//*[@id='dvReportInputs']/div[6]/label/span")
			//private static WebElement sm_IncludeServiceTypeItemChkBox;
			
			@FindBy(xpath="//select[@id='RITLayout_']")
			private static WebElement sl_LayoutDropdown;
			
			@FindBy(xpath="//select[@id='RITOutput_']")
			private static WebElement sl_OutputDropdown;
			
			@FindBy(xpath="//*[@id='dvReportInputs']/div[1]/div/label/span")
			private static WebElement sl_SelectAllItemsChkBox;
			
			@FindBy(xpath="//div[contains(@class,'first')]//input[contains(@class,'btn btn-xs Fbutton')]")
			private static WebElement sl_FirstBtn;
			
			@FindBy(xpath="//div[contains(@class,'prev')]//input[contains(@class,'btn btn-xs Fbutton')]")
			private static WebElement sl_PreviousBtn;
			
			@FindBy(xpath="//input[@id='btn1']")
			private static WebElement sl_PageBtn;
			
			@FindBy(xpath="//div[contains(@class,'next')]//input[contains(@class,'btn btn-xs Fbutton')]")
			private static WebElement sl_NextBtn;
			
			@FindBy(xpath="//div[contains(@class,'End')]//input[contains(@class,'btn btn-xs Fbutton')]")
			private static WebElement sl_EndBtn;
			
			@FindBy(xpath="//tr[1]//td[8]//div[1]//label[1]//input")
			private static WebElement sl_1stRowChkBox;
			
			@FindBy(xpath="//tr[2]//td[8]//div[1]//label[1]//input")
			private static WebElement sl_2ndRowChkBox;
			
			@FindBy(xpath="//tr[3]//td[8]//div[1]//label[1]//input")
			private static WebElement sl_3rdRowChkBox;
			
			
			
			////////
			
			@FindBy(xpath="//span[@id='reportSort']")
			private static WebElement report_sortingBtn;
					
			@FindBy(xpath="//span[@id='reportRefresh']")
			private static WebElement report_RefreshBtn;
					
			@FindBy(xpath="//span[@id='print_report_']")
			private static WebElement report_PrintBtn;
					
			@FindBy(xpath="//*[@id='reportRenderControls']/ul/li/span[4]")
			private static WebElement report_ExportBtn;
			
			@FindBy(xpath="//a[@id='pdf_report_']")
			private static WebElement report_ExportPDFBtn;
			
			
					
			@FindBy(xpath="//span[@id='GraphIcon']")
			private static WebElement report_GraphBtn;
					
			/*@FindBy(xpath="//span[@id='BackTrackIcon']")
			private static WebElement report_BackTrackBtn;*/
			
			
					
			/*@FindBy(xpath="//span[@id='reportCustomize_']")
			private static WebElement report_ReportCustomizeBtn;*/
			
			@FindBy(xpath="//*[@id='id_reportmenudisplay']/ul/li[8]")
			private static WebElement report_ReportCustomizeBtn;	
				
			/*@FindBy(xpath="//span[@id='filterIcon_']")
			private static WebElement report_FilterBtn;*/
					
			@FindBy(xpath="//*[@id='id_reportmenudisplay']/ul/li[9]/ul//li[1]")
			private static WebElement report_FilterBtn;
			
			
			@FindBy(xpath="//div[@id='dvfilter']//span[contains(@class,'icon-down-arrow icon-font8 dropdown-toggle theme_button_color')]")
			private static WebElement report_FilterDownBtn;
					
			@FindBy(xpath="//span[@id='analyzeIcon_']")
			private static WebElement report_AnalyzeBtn;
					
			@FindBy(xpath="//span[@id='CrossrefrenceIcon_']")
			private static WebElement report_CrossReferenceBtn;
					
			@FindBy(xpath="//span[contains(text(),'Options')]")
			private static WebElement report_OptionsBtn;
					
			@FindBy(xpath="//div[@id='dvoptions']//span[contains(@class,'icon-down-arrow icon-font8 dropdown-toggle theme_button_color')]")
			private static WebElement report_OptionsDownBtn;
					
			/*@FindBy(xpath="//span[@id='reportClose_']")
			private static WebElement report_CloseBtn;*/
			
			@FindBy(xpath="//*[@id='id_reportmenudisplay']/ul/li[13]")
			private static WebElement report_CloseBtn;
					
			@FindBy(xpath="//div[@id='tblFooterReportRender']/div[2]/button[1]")
			private static WebElement report_FirstBtn;
					
			@FindBy(xpath="//div[@id='tblFooterReportRender']/div[2]/button[2]")
			private static WebElement report_PreviousBtn;
					
			@FindBy(xpath="//div[@id='tblFooterReportRender']/div[2]/button[3]")
			private static WebElement report_PageNoBtn;
					
			@FindBy(xpath="//div[@id='tblFooterReportRender']/div[2]/button[4]")
			private static WebElement report_NextBtn;
					
			@FindBy(xpath="//div[@id='tblFooterReportRender']/div[2]/button[5]")
			private static WebElement report_LastBtn;
			
			@FindBy(xpath="//input[@id='txtSearchReport']")
			private static WebElement report_SearchTxt;
			
			
			@FindBy(xpath="//*[@id='trRender_0']/td[2]")
			private static WebElement sl_1stRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_0']/td[3]")
			private static WebElement sl_1stRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[4]")
			private static WebElement sl_1stRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[5]")
			private static WebElement sl_1stRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[6]")
			private static WebElement sl_1stRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[7]")
			private static WebElement sl_1stRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[8]")
			private static WebElement sl_1stRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[9]")
			private static WebElement sl_1stRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[10]")
			private static WebElement sl_1stRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[11]")
			private static WebElement sl_1stRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[12]")
			private static WebElement sl_1stRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[13]")
			private static WebElement sl_1stRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[14]")
			private static WebElement sl_1stRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[15]")
			private static WebElement sl_1stRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[16]")
			private static WebElement sl_1stRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[17]")
			private static WebElement sl_1stRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[18]")
			private static WebElement sl_1stRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[19]")
			private static WebElement sl_1stRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[20]")
			private static WebElement sl_1stRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[21]")
			private static WebElement sl_1stRow20thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[22]")
			private static WebElement sl_1stRow21thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[23]")
			private static WebElement sl_1stRow22thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[24]")
			private static WebElement sl_1stRow23thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[25]")
			private static WebElement sl_1stRow24thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[26]")
			private static WebElement sl_1stRow25thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[27]")
			private static WebElement sl_1stRow26thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[28]")
			private static WebElement sl_1stRow27thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[29]")
			private static WebElement sl_1stRow28thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[30]")
			private static WebElement sl_1stRow29thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[31]")
			private static WebElement sl_1stRow30thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[32]")
			private static WebElement sl_1stRow31thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[33]")
			private static WebElement sl_1stRow32thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[34]")
			private static WebElement sl_1stRow33thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[35]")
			private static WebElement sl_1stRow34thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[36]")
			private static WebElement sl_1stRow35thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[37]")
			private static WebElement sl_1stRow36thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[38]")
			private static WebElement sl_1stRow37thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[39]")
			private static WebElement sl_1stRow38thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[40]")
			private static WebElement sl_1stRow39thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[41]")
			private static WebElement sl_1stRow40thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[42]")
			private static WebElement sl_1stRow41thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[43]")
			private static WebElement sl_1stRow42thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[44]")
			private static WebElement sl_1stRow43thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[45]")
			private static WebElement sl_1stRow44thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[46]")
			private static WebElement sl_1stRow45thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[47]")
			private static WebElement sl_1stRow46thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[48]")
			private static WebElement sl_1stRow47thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[49]")
			private static WebElement sl_1stRow48thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[50]")
			private static WebElement sl_1stRow49thCol;
			
			@FindBy(xpath="//*[@id='trRender_0']/td[51]")
			private static WebElement sl_1stRow50thCol;
			
			
			@FindBy(xpath="//*[@id='trRender_1']/td[2]")
			private static WebElement sl_2ndRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_1']/td[3]")
			private static WebElement sl_2ndRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[4]")
			private static WebElement sl_2ndRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[5]")
			private static WebElement sl_2ndRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[6]")
			private static WebElement sl_2ndRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[7]")
			private static WebElement sl_2ndRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[8]")
			private static WebElement sl_2ndRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[9]")
			private static WebElement sl_2ndRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[10]")
			private static WebElement sl_2ndRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[11]")
			private static WebElement sl_2ndRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[12]")
			private static WebElement sl_2ndRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[13]")
			private static WebElement sl_2ndRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[14]")
			private static WebElement sl_2ndRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[15]")
			private static WebElement sl_2ndRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[16]")
			private static WebElement sl_2ndRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[17]")
			private static WebElement sl_2ndRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[18]")
			private static WebElement sl_2ndRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[19]")
			private static WebElement sl_2ndRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[20]")
			private static WebElement sl_2ndRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[21]")
			private static WebElement sl_2ndRow20thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[22]")
			private static WebElement sl_2ndRow21thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[23]")
			private static WebElement sl_2ndRow22thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[24]")
			private static WebElement sl_2ndRow23thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[25]")
			private static WebElement sl_2ndRow24thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[26]")
			private static WebElement sl_2ndRow25thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[27]")
			private static WebElement sl_2ndRow26thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[28]")
			private static WebElement sl_2ndRow27thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[29]")
			private static WebElement sl_2ndRow28thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[30]")
			private static WebElement sl_2ndRow29thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[31]")
			private static WebElement sl_2ndRow30thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[32]")
			private static WebElement sl_2ndRow31thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[33]")
			private static WebElement sl_2ndRow32thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[34]")
			private static WebElement sl_2ndRow33thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[35]")
			private static WebElement sl_2ndRow34thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[36]")
			private static WebElement sl_2ndRow35thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[37]")
			private static WebElement sl_2ndRow36thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[38]")
			private static WebElement sl_2ndRow37thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[39]")
			private static WebElement sl_2ndRow38thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[40]")
			private static WebElement sl_2ndRow39thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[41]")
			private static WebElement sl_2ndRow40thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[42]")
			private static WebElement sl_2ndRow41thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[43]")
			private static WebElement sl_2ndRow42thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[44]")
			private static WebElement sl_2ndRow43thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[45]")
			private static WebElement sl_2ndRow44thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[46]")
			private static WebElement sl_2ndRow45thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[47]")
			private static WebElement sl_2ndRow46thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[48]")
			private static WebElement sl_2ndRow47thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[49]")
			private static WebElement sl_2ndRow48thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[50]")
			private static WebElement sl_2ndRow49thCol;
			
			@FindBy(xpath="//*[@id='trRender_1']/td[51]")
			private static WebElement sl_2ndRow50thCol;
			
			
			
			
			
			@FindBy(xpath="//*[@id='trRender_2']/td[2]")
			private static WebElement sl_3rdRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_2']/td[3]")
			private static WebElement sl_3rdRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[4]")
			private static WebElement sl_3rdRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[5]")
			private static WebElement sl_3rdRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[6]")
			private static WebElement sl_3rdRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[7]")
			private static WebElement sl_3rdRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[8]")
			private static WebElement sl_3rdRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[9]")
			private static WebElement sl_3rdRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[10]")
			private static WebElement sl_3rdRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[11]")
			private static WebElement sl_3rdRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[12]")
			private static WebElement sl_3rdRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[13]")
			private static WebElement sl_3rdRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[14]")
			private static WebElement sl_3rdRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[15]")
			private static WebElement sl_3rdRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[16]")
			private static WebElement sl_3rdRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[17]")
			private static WebElement sl_3rdRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[18]")
			private static WebElement sl_3rdRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[19]")
			private static WebElement sl_3rdRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[20]")
			private static WebElement sl_3rdRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[21]")
			private static WebElement sl_3rdRow20thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[22]")
			private static WebElement sl_3rdRow21thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[23]")
			private static WebElement sl_3rdRow22thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[24]")
			private static WebElement sl_3rdRow23thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[25]")
			private static WebElement sl_3rdRow24thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[26]")
			private static WebElement sl_3rdRow25thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[27]")
			private static WebElement sl_3rdRow26thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[28]")
			private static WebElement sl_3rdRow27thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[29]")
			private static WebElement sl_3rdRow28thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[30]")
			private static WebElement sl_3rdRow29thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[31]")
			private static WebElement sl_3rdRow30thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[32]")
			private static WebElement sl_3rdRow31thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[33]")
			private static WebElement sl_3rdRow32thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[34]")
			private static WebElement sl_3rdRow33thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[35]")
			private static WebElement sl_3rdRow34thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[36]")
			private static WebElement sl_3rdRow35thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[37]")
			private static WebElement sl_3rdRow36thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[38]")
			private static WebElement sl_3rdRow37thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[39]")
			private static WebElement sl_3rdRow38thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[40]")
			private static WebElement sl_3rdRow39thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[41]")
			private static WebElement sl_3rdRow40thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[42]")
			private static WebElement sl_3rdRow41thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[43]")
			private static WebElement sl_3rdRow42thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[44]")
			private static WebElement sl_3rdRow43thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[45]")
			private static WebElement sl_3rdRow44thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[46]")
			private static WebElement sl_3rdRow45thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[47]")
			private static WebElement sl_3rdRow46thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[48]")
			private static WebElement sl_3rdRow47thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[49]")
			private static WebElement sl_3rdRow48thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[50]")
			private static WebElement sl_3rdRow49thCol;
			
			@FindBy(xpath="//*[@id='trRender_2']/td[51]")
			private static WebElement sl_3rdRow50thCol;
			
			
			
			
			
			@FindBy(xpath="//*[@id='trRender_3']/td[2]")
			private static WebElement sl_4thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_3']/td[3]")
			private static WebElement sl_4thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[4]")
			private static WebElement sl_4thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[5]")
			private static WebElement sl_4thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[6]")
			private static WebElement sl_4thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[7]")
			private static WebElement sl_4thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[8]")
			private static WebElement sl_4thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[9]")
			private static WebElement sl_4thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[10]")
			private static WebElement sl_4thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[11]")
			private static WebElement sl_4thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[12]")
			private static WebElement sl_4thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[13]")
			private static WebElement sl_4thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[14]")
			private static WebElement sl_4thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[15]")
			private static WebElement sl_4thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[16]")
			private static WebElement sl_4thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[17]")
			private static WebElement sl_4thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[18]")
			private static WebElement sl_4thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[19]")
			private static WebElement sl_4thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[20]")
			private static WebElement sl_4thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[21]")
			private static WebElement sl_4thRow20thCol;
			

			@FindBy(xpath="//*[@id='trRender_3']/td[22]")
			private static WebElement sl_4thRow21thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[23]")
			private static WebElement sl_4thRow22thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[24]")
			private static WebElement sl_4thRow23thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[25]")
			private static WebElement sl_4thRow24thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[26]")
			private static WebElement sl_4thRow25thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[27]")
			private static WebElement sl_4thRow26thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[28]")
			private static WebElement sl_4thRow27thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[29]")
			private static WebElement sl_4thRow28thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[30]")
			private static WebElement sl_4thRow29thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[31]")
			private static WebElement sl_4thRow30thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[32]")
			private static WebElement sl_4thRow31thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[33]")
			private static WebElement sl_4thRow32thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[34]")
			private static WebElement sl_4thRow33thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[35]")
			private static WebElement sl_4thRow34thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[36]")
			private static WebElement sl_4thRow35thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[37]")
			private static WebElement sl_4thRow36thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[38]")
			private static WebElement sl_4thRow37thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[39]")
			private static WebElement sl_4thRow38thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[40]")
			private static WebElement sl_4thRow39thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[41]")
			private static WebElement sl_4thRow40thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[42]")
			private static WebElement sl_4thRow41thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[43]")
			private static WebElement sl_4thRow42thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[44]")
			private static WebElement sl_4thRow43thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[45]")
			private static WebElement sl_4thRow44thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[46]")
			private static WebElement sl_4thRow45thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[47]")
			private static WebElement sl_4thRow46thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[48]")
			private static WebElement sl_4thRow47thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[49]")
			private static WebElement sl_4thRow48thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[50]")
			private static WebElement sl_4thRow49thCol;
			
			@FindBy(xpath="//*[@id='trRender_3']/td[51]")
			private static WebElement sl_4thRow50thCol;
			
			

			
			
			
			
			
			@FindBy(xpath="//*[@id='trRender_4']/td[2]")
			private static WebElement sl_5thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_4']/td[3]")
			private static WebElement sl_5thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[4]")
			private static WebElement sl_5thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[5]")
			private static WebElement sl_5thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[6]")
			private static WebElement sl_5thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[7]")
			private static WebElement sl_5thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[8]")
			private static WebElement sl_5thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[9]")
			private static WebElement sl_5thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[10]")
			private static WebElement sl_5thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[11]")
			private static WebElement sl_5thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[12]")
			private static WebElement sl_5thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[13]")
			private static WebElement sl_5thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[14]")
			private static WebElement sl_5thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[15]")
			private static WebElement sl_5thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[16]")
			private static WebElement sl_5thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[17]")
			private static WebElement sl_5thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[18]")
			private static WebElement sl_5thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[19]")
			private static WebElement sl_5thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[20]")
			private static WebElement sl_5thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[21]")
			private static WebElement sl_5thRow20thCol;
			

			@FindBy(xpath="//*[@id='trRender_4']/td[22]")
			private static WebElement sl_5thRow21thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[23]")
			private static WebElement sl_5thRow22thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[24]")
			private static WebElement sl_5thRow23thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[25]")
			private static WebElement sl_5thRow24thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[26]")
			private static WebElement sl_5thRow25thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[27]")
			private static WebElement sl_5thRow26thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[28]")
			private static WebElement sl_5thRow27thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[29]")
			private static WebElement sl_5thRow28thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[30]")
			private static WebElement sl_5thRow29thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[31]")
			private static WebElement sl_5thRow30thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[32]")
			private static WebElement sl_5thRow31thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[33]")
			private static WebElement sl_5thRow32thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[34]")
			private static WebElement sl_5thRow33thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[35]")
			private static WebElement sl_5thRow34thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[36]")
			private static WebElement sl_5thRow35thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[37]")
			private static WebElement sl_5thRow36thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[38]")
			private static WebElement sl_5thRow37thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[39]")
			private static WebElement sl_5thRow38thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[40]")
			private static WebElement sl_5thRow39thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[41]")
			private static WebElement sl_5thRow40thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[42]")
			private static WebElement sl_5thRow41thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[43]")
			private static WebElement sl_5thRow42thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[44]")
			private static WebElement sl_5thRow43thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[45]")
			private static WebElement sl_5thRow44thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[46]")
			private static WebElement sl_5thRow45thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[47]")
			private static WebElement sl_5thRow46thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[48]")
			private static WebElement sl_5thRow47thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[49]")
			private static WebElement sl_5thRow48thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[50]")
			private static WebElement sl_5thRow49thCol;
			
			@FindBy(xpath="//*[@id='trRender_4']/td[51]")
			private static WebElement sl_5thRow50thCol;
			
			
			
			

			
			@FindBy(xpath="//*[@id='trRender_5']/td[2]")
			private static WebElement sl_6thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_5']/td[3]")
			private static WebElement sl_6thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[4]")
			private static WebElement sl_6thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[5]")
			private static WebElement sl_6thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[6]")
			private static WebElement sl_6thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[7]")
			private static WebElement sl_6thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[8]")
			private static WebElement sl_6thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[9]")
			private static WebElement sl_6thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[10]")
			private static WebElement sl_6thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[11]")
			private static WebElement sl_6thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[12]")
			private static WebElement sl_6thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[13]")
			private static WebElement sl_6thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[14]")
			private static WebElement sl_6thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[15]")
			private static WebElement sl_6thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[16]")
			private static WebElement sl_6thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[17]")
			private static WebElement sl_6thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[18]")
			private static WebElement sl_6thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[19]")
			private static WebElement sl_6thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[20]")
			private static WebElement sl_6thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[21]")
			private static WebElement sl_6thRow20thCol;
			

			@FindBy(xpath="//*[@id='trRender_5']/td[22]")
			private static WebElement sl_6thRow21thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[23]")
			private static WebElement sl_6thRow22thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[24]")
			private static WebElement sl_6thRow23thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[25]")
			private static WebElement sl_6thRow24thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[26]")
			private static WebElement sl_6thRow25thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[27]")
			private static WebElement sl_6thRow26thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[28]")
			private static WebElement sl_6thRow27thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[29]")
			private static WebElement sl_6thRow28thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[30]")
			private static WebElement sl_6thRow29thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[31]")
			private static WebElement sl_6thRow30thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[32]")
			private static WebElement sl_6thRow31thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[33]")
			private static WebElement sl_6thRow32thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[34]")
			private static WebElement sl_6thRow33thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[35]")
			private static WebElement sl_6thRow34thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[36]")
			private static WebElement sl_6thRow35thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[37]")
			private static WebElement sl_6thRow36thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[38]")
			private static WebElement sl_6thRow37thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[39]")
			private static WebElement sl_6thRow38thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[40]")
			private static WebElement sl_6thRow39thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[41]")
			private static WebElement sl_6thRow40thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[42]")
			private static WebElement sl_6thRow41thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[43]")
			private static WebElement sl_6thRow42thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[44]")
			private static WebElement sl_6thRow43thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[45]")
			private static WebElement sl_6thRow44thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[46]")
			private static WebElement sl_6thRow45thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[47]")
			private static WebElement sl_6thRow46thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[48]")
			private static WebElement sl_6thRow47thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[49]")
			private static WebElement sl_6thRow48thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[50]")
			private static WebElement sl_6thRow49thCol;
			
			@FindBy(xpath="//*[@id='trRender_5']/td[51]")
			private static WebElement sl_6thRow50thCol;
			
			
			
			
			
			
			
			
			@FindBy(xpath="//*[@id='trRender_6']/td[2]")
			private static WebElement sl_7thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_6']/td[3]")
			private static WebElement sl_7thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[4]")
			private static WebElement sl_7thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[5]")
			private static WebElement sl_7thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[6]")
			private static WebElement sl_7thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[7]")
			private static WebElement sl_7thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[8]")
			private static WebElement sl_7thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[9]")
			private static WebElement sl_7thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[10]")
			private static WebElement sl_7thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[11]")
			private static WebElement sl_7thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[12]")
			private static WebElement sl_7thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[13]")
			private static WebElement sl_7thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[14]")
			private static WebElement sl_7thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[15]")
			private static WebElement sl_7thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[16]")
			private static WebElement sl_7thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[17]")
			private static WebElement sl_7thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[18]")
			private static WebElement sl_7thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[19]")
			private static WebElement sl_7thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[20]")
			private static WebElement sl_7thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_6']/td[21]")
			private static WebElement sl_7thRow20thCol;
			
			
			
			
			
			@FindBy(xpath="//*[@id='trRender_7']/td[2]")
			private static WebElement sl_8thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_7']/td[3]")
			private static WebElement sl_8thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[4]")
			private static WebElement sl_8thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[5]")
			private static WebElement sl_8thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[6]")
			private static WebElement sl_8thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[7]")
			private static WebElement sl_8thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[8]")
			private static WebElement sl_8thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[9]")
			private static WebElement sl_8thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[10]")
			private static WebElement sl_8thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[11]")
			private static WebElement sl_8thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[12]")
			private static WebElement sl_8thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[13]")
			private static WebElement sl_8thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[14]")
			private static WebElement sl_8thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[15]")
			private static WebElement sl_8thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[16]")
			private static WebElement sl_8thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[17]")
			private static WebElement sl_8thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[18]")
			private static WebElement sl_8thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[19]")
			private static WebElement sl_8thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[20]")
			private static WebElement sl_8thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_7']/td[21]")
			private static WebElement sl_8thRow20thCol;
			
			
			
			@FindBy(xpath="//*[@id='trRender_8']/td[2]")
			private static WebElement sl_9thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_8']/td[3]")
			private static WebElement sl_9thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[4]")
			private static WebElement sl_9thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[5]")
			private static WebElement sl_9thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[6]")
			private static WebElement sl_9thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[7]")
			private static WebElement sl_9thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[8]")
			private static WebElement sl_9thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[9]")
			private static WebElement sl_9thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[10]")
			private static WebElement sl_9thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[11]")
			private static WebElement sl_9thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[12]")
			private static WebElement sl_9thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[13]")
			private static WebElement sl_9thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[14]")
			private static WebElement sl_9thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[15]")
			private static WebElement sl_9thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[16]")
			private static WebElement sl_9thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[17]")
			private static WebElement sl_9thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[18]")
			private static WebElement sl_9thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[19]")
			private static WebElement sl_9thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[20]")
			private static WebElement sl_9thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_8']/td[21]")
			private static WebElement sl_9thRow20thCol;
			
			
			
			@FindBy(xpath="//*[@id='trRender_9']/td[2]")
			private static WebElement sl_10thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_9']/td[3]")
			private static WebElement sl_10thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[4]")
			private static WebElement sl_10thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[5]")
			private static WebElement sl_10thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[6]")
			private static WebElement sl_10thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[7]")
			private static WebElement sl_10thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[8]")
			private static WebElement sl_10thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[9]")
			private static WebElement sl_10thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[10]")
			private static WebElement sl_10thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[11]")
			private static WebElement sl_10thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[12]")
			private static WebElement sl_10thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[13]")
			private static WebElement sl_10thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[14]")
			private static WebElement sl_10thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[15]")
			private static WebElement sl_10thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[16]")
			private static WebElement sl_10thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[17]")
			private static WebElement sl_10thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[18]")
			private static WebElement sl_10thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[19]")
			private static WebElement sl_10thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[20]")
			private static WebElement sl_10thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_9']/td[21]")
			private static WebElement sl_10thRow20thCol;
			
			
			@FindBy(xpath="//*[@id='trRender_10']/td[2]")
			private static WebElement sl_11thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_10']/td[3]")
			private static WebElement sl_11thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[4]")
			private static WebElement sl_11thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[5]")
			private static WebElement sl_11thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[6]")
			private static WebElement sl_11thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[7]")
			private static WebElement sl_11thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[8]")
			private static WebElement sl_11thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[9]")
			private static WebElement sl_11thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[10]")
			private static WebElement sl_11thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[11]")
			private static WebElement sl_11thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[12]")
			private static WebElement sl_11thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[13]")
			private static WebElement sl_11thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[14]")
			private static WebElement sl_11thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[15]")
			private static WebElement sl_11thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[16]")
			private static WebElement sl_11thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[17]")
			private static WebElement sl_11thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[18]")
			private static WebElement sl_11thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[19]")
			private static WebElement sl_11thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[20]")
			private static WebElement sl_11thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_10']/td[21]")
			private static WebElement sl_11thRow20thCol;
			
			
			@FindBy(xpath="//*[@id='trRender_11']/td[2]")
			private static WebElement sl_12thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_11']/td[3]")
			private static WebElement sl_12thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[4]")
			private static WebElement sl_12thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[5]")
			private static WebElement sl_12thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[6]")
			private static WebElement sl_12thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[7]")
			private static WebElement sl_12thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[8]")
			private static WebElement sl_12thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[9]")
			private static WebElement sl_12thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[10]")
			private static WebElement sl_12thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[11]")
			private static WebElement sl_12thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[12]")
			private static WebElement sl_12thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[13]")
			private static WebElement sl_12thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[14]")
			private static WebElement sl_12thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[15]")
			private static WebElement sl_12thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[16]")
			private static WebElement sl_12thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[17]")
			private static WebElement sl_12thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[18]")
			private static WebElement sl_12thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[19]")
			private static WebElement sl_12thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[20]")
			private static WebElement sl_12thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_11']/td[21]")
			private static WebElement sl_12thRow20thCol;
			
			
			
			@FindBy(xpath="//*[@id='trRender_12']/td[2]")
			private static WebElement sl_13thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_12']/td[3]")
			private static WebElement sl_13thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[4]")
			private static WebElement sl_13thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[5]")
			private static WebElement sl_13thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[6]")
			private static WebElement sl_13thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[7]")
			private static WebElement sl_13thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[8]")
			private static WebElement sl_13thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[9]")
			private static WebElement sl_13thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[10]")
			private static WebElement sl_13thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[11]")
			private static WebElement sl_13thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[12]")
			private static WebElement sl_13thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[13]")
			private static WebElement sl_13thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[14]")
			private static WebElement sl_13thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[15]")
			private static WebElement sl_13thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[16]")
			private static WebElement sl_13thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[17]")
			private static WebElement sl_13thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[18]")
			private static WebElement sl_13thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[19]")
			private static WebElement sl_13thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[20]")
			private static WebElement sl_13thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_12']/td[21]")
			private static WebElement sl_13thRow20thCol;
			
			
			
			@FindBy(xpath="//*[@id='trRender_13']/td[2]")
			private static WebElement sl_14thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_13']/td[3]")
			private static WebElement sl_14thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[4]")
			private static WebElement sl_14thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[5]")
			private static WebElement sl_14thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[6]")
			private static WebElement sl_14thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[7]")
			private static WebElement sl_14thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[8]")
			private static WebElement sl_14thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[9]")
			private static WebElement sl_14thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[10]")
			private static WebElement sl_14thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[11]")
			private static WebElement sl_14thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[12]")
			private static WebElement sl_14thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[13]")
			private static WebElement sl_14thRow12thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[14]")
			private static WebElement sl_14thRow13thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[15]")
			private static WebElement sl_14thRow14thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[16]")
			private static WebElement sl_14thRow15thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[17]")
			private static WebElement sl_14thRow16thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[18]")
			private static WebElement sl_14thRow17thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[19]")
			private static WebElement sl_14thRow18thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[20]")
			private static WebElement sl_14thRow19thCol;
			
			@FindBy(xpath="//*[@id='trRender_13']/td[21]")
			private static WebElement sl_14thRow20thCol;
			
			
			
			@FindBy(xpath="//*[@id='trRender_14']/td[2]")
			private static WebElement sl_15thRow1stCol;
					
			@FindBy(xpath="//*[@id='trRender_14']/td[3]")
			private static WebElement sl_15thRow2ndCol;
			
			@FindBy(xpath="//*[@id='trRender_14']/td[4]")
			private static WebElement sl_15thRow3rdCol;
			
			@FindBy(xpath="//*[@id='trRender_14']/td[5]")
			private static WebElement sl_15thRow4thCol;
			
			@FindBy(xpath="//*[@id='trRender_14']/td[6]")
			private static WebElement sl_15thRow5thCol;
			
			@FindBy(xpath="//*[@id='trRender_14']/td[7]")
			private static WebElement sl_15thRow6thCol;
			
			@FindBy(xpath="//*[@id='trRender_14']/td[8]")
			private static WebElement sl_15thRow7thCol;
			
			@FindBy(xpath="//*[@id='trRender_14']/td[9]")
			private static WebElement sl_15thRow8thCol;
			
			@FindBy(xpath="//*[@id='trRender_14']/td[10]")
			private static WebElement sl_15thRow9thCol;
			
			@FindBy(xpath="//*[@id='trRender_14']/td[11]")
			private static WebElement sl_15thRow10thCol;
			
			@FindBy(xpath="//*[@id='trRender_14']/td[12]")
			private static WebElement sl_15thRow11thCol;
			
			@FindBy(xpath="//*[@id='trRender_14']/td[13]")
			private static WebElement sl_15thRow12thCol;
			
			
			
			@FindBy(xpath="//tbody[@id='LandingGridBody']/tr")
			private static List<WebElement> stockLedgerHometableRowCount;
			
			
			 // Reports Table Row List  
		  	@FindBy(xpath="(//tr[@id='trRender_0'])[1]/td")
			private static List<WebElement> reportsRow1List;
			
			@FindBy(xpath="(//tr[@id='trRender_1'])[1]/td")
			private static List<WebElement> reportsRow2List;
			
			@FindBy(xpath="(//tr[@id='trRender_2'])[1]/td")
			private static List<WebElement> reportsRow3List;
			
			@FindBy(xpath="(//tr[@id='trRender_3'])[1]/td")
			private static List<WebElement> reportsRow4List;
			
			@FindBy(xpath="(//tr[@id='trRender_4'])[1]/td")
			private static List<WebElement> reportsRow5List;
			
			@FindBy(xpath="(//tr[@id='trRender_5'])[1]/td")
			private static List<WebElement> reportsRow6List;
			
			@FindBy(xpath="(//tr[@id='trRender_6'])[1]/td")
			private static List<WebElement> reportsRow7List;
			
			@FindBy(xpath="(//tr[@id='trRender_7'])[1]/td")
			private static List<WebElement> reportsRow8List;
			
			@FindBy(xpath="(//tr[@id='trRender_8'])[1]/td")
			private static List<WebElement> reportsRow9List;
			
			@FindBy(xpath="(//tr[@id='trRender_9'])[1]/td")
			private static List<WebElement> reportsRow10List;
			
			@FindBy(xpath="(//tr[@id='trRender_10'])[1]/td")
			private static List<WebElement> reportsRow11List;
			
			@FindBy(xpath="(//tr[@id='trRender_11'])[1]/td")
			private static List<WebElement> reportsRow12List;
			
			@FindBy(xpath="(//tr[@id='trRender_12'])[1]/td")
			private static List<WebElement> reportsRow13List;
			
			@FindBy(xpath="(//tr[@id='trRender_13'])[1]/td")
			private static List<WebElement> reportsRow14List;
			
			
			
			//@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[12]")
			
			@FindBy(xpath="//*[@id='landgridData']/tbody/tr/td[12]")
			private static List<WebElement> stockLedgerHometableItemNamesList;
			
			@FindBy(xpath="//tbody[@id='LandingGridBody']/tr/td[8]/div/label/input")
			private static List<WebElement> stockLedgerHometableItemChkboxList;
			
			@FindBy(xpath="(//div[@id='dvReportDetails']/div/table/thead)[1]/tr/th")
			private static List<WebElement> reportsHeaderList;
			
			
			
	       	
			public boolean checkStockLedgerReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				 excelReader=new ExcelReader(POJOUtility.getExcelPath());
				 
				 LogoutandLoginwithSU();
				 
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger));
				stockLedger.click();
				
				Thread.sleep(2000);
				
				int rowcount=stockLedgerHometableRowCount.size();
				
				System.out.println(rowcount);
				
				for (int i = 1; i < rowcount; i++) 
				{
					String actName = stockLedgerHometableItemNamesList.get(i).getText();
					
					System.out.println(actName);
					
								
					if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 109, 5)))
					{
						stockLedgerHometableItemChkboxList.get(i).click();
						break;
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(2000);
				
				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List = excelReader.getCellData(xlSheetName, 109, 6);
				excelReader.setCellData(xlfile, xlSheetName, 109, 7, actRow1List);

				

				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List = excelReader.getCellData(xlSheetName, 110, 6);
				excelReader.setCellData(xlfile, xlSheetName, 110, 7, actRow2List);

				
				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List = excelReader.getCellData(xlSheetName, 111, 6);
				excelReader.setCellData(xlfile, xlSheetName, 111, 7, actRow3List);

				
				
				int reportsRow4ListCount = reportsRow4List.size();
				ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow4ListCount;i++)
				{
					String data = reportsRow4List.get(i).getText();
					reportsRow4ListArray.add(data);
				}
				String actRow4List = reportsRow4ListArray.toString();
				String expRow4List = excelReader.getCellData(xlSheetName, 112, 6);
				excelReader.setCellData(xlfile, xlSheetName, 112, 7, actRow4List);

				
				
				int reportsRow5ListCount = reportsRow5List.size();
				ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow5ListCount;i++)
				{
					String data = reportsRow5List.get(i).getText();
					reportsRow5ListArray.add(data);
				}
				String actRow5List = reportsRow5ListArray.toString();
				String expRow5List = excelReader.getCellData(xlSheetName, 113, 6);
				excelReader.setCellData(xlfile, xlSheetName, 113, 7, actRow5List);

				
				
				int reportsRow6ListCount = reportsRow6List.size();
				ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow6ListCount;i++)
				{
					String data = reportsRow6List.get(i).getText();
					reportsRow6ListArray.add(data);
				}
				String actRow6List = reportsRow6ListArray.toString();
				String expRow6List = excelReader.getCellData(xlSheetName, 114, 6);
				excelReader.setCellData(xlfile, xlSheetName, 114, 7, actRow6List);

				
				
				int reportsRow7ListCount = reportsRow7List.size();
				ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow7ListCount;i++)
				{
					String data = reportsRow7List.get(i).getText();
					reportsRow7ListArray.add(data);
				}
				String actRow7List = reportsRow7ListArray.toString();
				String expRow7List = excelReader.getCellData(xlSheetName, 115, 6);
				excelReader.setCellData(xlfile, xlSheetName, 115, 7, actRow7List);

				
				
				int reportsRow8ListCount = reportsRow8List.size();
				ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow8ListCount;i++)
				{
					String data = reportsRow8List.get(i).getText();
					reportsRow8ListArray.add(data);
				}
				String actRow8List = reportsRow8ListArray.toString();
				String expRow8List = excelReader.getCellData(xlSheetName, 116, 6);
				excelReader.setCellData(xlfile, xlSheetName, 117, 7, actRow8List);

				
				
				int reportsRow9ListCount = reportsRow9List.size();
				ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow9ListCount;i++)
				{
					String data = reportsRow9List.get(i).getText();
					reportsRow9ListArray.add(data);
				}
				String actRow9List = reportsRow9ListArray.toString();
				String expRow9List = excelReader.getCellData(xlSheetName, 117, 6);
				excelReader.setCellData(xlfile, xlSheetName, 117, 7, actRow9List);

				
				
				int reportsRow10ListCount = reportsRow10List.size();
				ArrayList<String> reportsRow10ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow10ListCount;i++)
				{
					String data = reportsRow10List.get(i).getText();
					reportsRow10ListArray.add(data);
				}
				String actRow10List = reportsRow10ListArray.toString();
				String expRow10List = excelReader.getCellData(xlSheetName, 118, 6);
				excelReader.setCellData(xlfile, xlSheetName, 118, 7, actRow10List);

				
			
				int reportsRow11ListCount = reportsRow11List.size();
				ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow11ListCount;i++)
				{
					String data = reportsRow11List.get(i).getText();
					reportsRow11ListArray.add(data);
				}
				String actRow11List = reportsRow11ListArray.toString();
				String expRow11List = excelReader.getCellData(xlSheetName, 119, 6);
				excelReader.setCellData(xlfile, xlSheetName, 119, 7, actRow11List);

				
				
				/*int reportsRow12ListCount = reportsRow12List.size();
				ArrayList<String> reportsRow12ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow12ListCount;i++)
				{
					String data = reportsRow12List.get(i).getText();
					reportsRow12ListArray.add(data);
				}
				String actRow12List = reportsRow12ListArray.toString();
				String expRow12List = "[Grand Total, , 66.00, 38.41, 66.00, 22.23, , 405.11, 170.45, 318.66, 53.23, , , ]";
				
				*/
				
				
				
				System.out.println("************************************checkStockLedgerReport********************************************");
			
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
			
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				System.out.println(actRow4List);
				System.out.println(expRow4List);
				
				System.out.println(actRow5List);
				System.out.println(expRow5List);
				
				System.out.println(actRow6List);
				System.out.println(expRow6List);
				
				System.out.println(actRow7List);
				System.out.println(expRow7List);
				
				System.out.println(actRow8List);
				System.out.println(expRow8List);
				
				System.out.println(actRow9List);
				System.out.println(expRow9List);
				
				System.out.println(actRow10List);
				System.out.println(expRow10List);
				
				System.out.println(actRow11List);
				System.out.println(expRow11List);
				
				/*System.out.println(actRow12List);
				System.out.println(expRow12List);*/
				
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
					&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
					&& actRow10List.equalsIgnoreCase(expRow10List) &&actRow11List.equalsIgnoreCase(expRow11List) /*&& actRow12List.equalsIgnoreCase(expRow12List)*/)
				{
					excelReader.setCellData(xlfile, xlSheetName, 109, 8, resPass);
					return true;
				}
				else
				{
					excelReader.setCellData(xlfile, xlSheetName, 109, 8, resFail);

					return false;
				}
			}
				
			
			// With Customizing Warehouse
			
			@FindBy(xpath="//*[@id='CustomizeTransTree']/ul/li/span")
			private static WebElement  osr_customizeFieldsTransactionExpansion; 
			
			@FindBy(xpath="//*[@id='rd_customization_tree2']/span")
			private static WebElement  osr_customizeFieldsTransExtraFieldsExpansion; 
			
			@FindBy(xpath="//*[@id='rd_customization_tree2_43']/span/span")
			private static WebElement  osr_customizeFieldsTransExtraFieldsWarehouseExpansion; 
			
			@FindBy(xpath="//*[@id='rd_customization_tree2_43_0']")
			private static WebElement  osr_customizeFieldsTransExtraFieldsWarehouseName; 
			
			@FindBy(xpath="//li[@id='rd_customization_tree2_15']")
			private static WebElement  osr_customizeFieldsTransExtraFieldsNarration;
			
			@FindBy(xpath="//*[@id='id_rc_columnheadertable']/thead/tr/th")
			private static List<WebElement> osr_customizeHeaderList;
			
			@FindBy(xpath="//table[@id='tblRDRender']/thead/tr/th")
			private static List<WebElement> osr_reportHeaderList;
			
			
			@FindBy(xpath="//ul[@id='id_Item_main']/li/span")
			private static WebElement  osr_customizeFieldsItemExpansion; 
			
			@FindBy(xpath="//li[@id='id_Item_main_0']")
			private static WebElement  osr_customizeFieldsItemName; 
			
			
			@FindBy(xpath="//input[@id='HidethisColumn']")
			private static WebElement osr_customizeTabHideThisColumnChkBox;
			
			@FindBy(xpath="//input[@id='ScrollTexttonextLine']")
			private static WebElement osr_customizeTabScrollTextToNextLineChkBox;
			
			@FindBy(xpath="//input[@id='Printunderpreviouscolumn']")
			private static WebElement osr_customizeTabPrintUnderPreviousColumnChkBox;
			
			@FindBy(xpath="//input[@id='Mandatorycolumn']")
			private static WebElement osr_customizeTabMandatoryColumnChkBox;
			
			@FindBy(xpath="//*[@id='plnCustomizecolumn']/div/span[4]/a")
			private static WebElement osr_customizeTabRemoveColumnBtn;
			
			//@FindBy(xpath="//*[@id='Save']")
			@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[7]")
			private static WebElement  osr_customizeSaveBtn;
			
			@FindBy(xpath="//*[@id='Deleteayout']/div[2]")
			private static WebElement  osr_customizeDeleteLayoutBtn;
			
			@FindBy(xpath="//div[@id='btnCustomizeClose']")
			private static WebElement  osr_customizeCloseBtn;
			
			@FindBy(xpath="//*[@id='navbarSupportedContent2']/ul/li[4]")
			private static WebElement report_CustomizationDeleteLayoutBtn;
			
			//delete column in layout
			@FindBy(xpath="//*[@id='id_rc_columnheadertable']//tr[1]/th")
			private static List<WebElement> Listcolumnnames;
		
			@FindBy(xpath="//*[@id='id_rc_columnheadertable']//tr[1]/th[11]")
			private static WebElement warehouseColumn;
			@FindBy(xpath="//*[@id='plnCustomizecolumn']/div/span[4]")
			private static WebElement removeColumn;
			
			public boolean checkStockLedgerCustomizationReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				 excelReader=new ExcelReader(POJOUtility.getExcelPath());

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger));
				stockLedger.click();
				int rowcount=stockLedgerHometableRowCount.size();
				
				System.out.println(rowcount);
				
				for (int i = 1; i <= rowcount; i++) 
				{
					String actName = stockLedgerHometableItemNamesList.get(i).getText();
					Thread.sleep(1000);
					
					System.out.println(actName);
					
					if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 121, 5)))
					{
						stockLedgerHometableItemChkboxList.get(i).click();
						break;
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_ReportCustomizeBtn));
				report_ReportCustomizeBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeFieldsTransactionExpansion));
				osr_customizeFieldsTransactionExpansion.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeFieldsTransExtraFieldsExpansion));
				osr_customizeFieldsTransExtraFieldsExpansion.click();
				
				getAction().moveToElement(osr_customizeFieldsTransExtraFieldsWarehouseExpansion).build().perform();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeFieldsTransExtraFieldsWarehouseExpansion));
				osr_customizeFieldsTransExtraFieldsWarehouseExpansion.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeFieldsTransExtraFieldsWarehouseName));
				getAction().doubleClick(osr_customizeFieldsTransExtraFieldsWarehouseName).build().perform();
				
				Thread.sleep(2000);	
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeSaveBtn));
				osr_customizeSaveBtn.click();
				
				String expMessage = excelReader.getCellData(xlSheetName, 121, 6);
			    
			    String actMessage = checkValidationMessage(expMessage);
				excelReader.setCellData(xlfile, xlSheetName, 121, 7, actMessage);

				
				Thread.sleep(2000);
				
				/*if (sl_OkBtn.isDisplayed()==true) 
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
					sl_OkBtn.click();
					
					Thread.sleep(4000);
				}
			    else
			    {
			    	if (osr_customizeCloseBtn.isDisplayed()==true) 
			    	{
						getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeCloseBtn));
						osr_customizeCloseBtn.click();
						
						Thread.sleep(4000);
					}
			    }*/
				
				/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
				report_CloseBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(2000);*/
				
				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List = excelReader.getCellData(xlSheetName, 122, 6);
				excelReader.setCellData(xlfile, xlSheetName, 122, 7, actRow1List);

				

				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List = excelReader.getCellData(xlSheetName, 123, 6);
				excelReader.setCellData(xlfile, xlSheetName, 123, 7, actRow2List);
				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List = excelReader.getCellData(xlSheetName, 124, 6);
				excelReader.setCellData(xlfile, xlSheetName, 124, 7, actRow3List);
				
				
				
				int reportsRow4ListCount = reportsRow4List.size();
				ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow4ListCount;i++)
				{
					String data = reportsRow4List.get(i).getText();
					reportsRow4ListArray.add(data);
				}
				String actRow4List = reportsRow4ListArray.toString();
				String expRow4List = excelReader.getCellData(xlSheetName, 125, 6);
				excelReader.setCellData(xlfile, xlSheetName, 125, 7, actRow4List);
				
				
				
				int reportsRow5ListCount = reportsRow5List.size();
				ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow5ListCount;i++)
				{
					String data = reportsRow5List.get(i).getText();
					reportsRow5ListArray.add(data);
				}
				String actRow5List = reportsRow5ListArray.toString();
				String expRow5List = excelReader.getCellData(xlSheetName, 126, 6);
				excelReader.setCellData(xlfile, xlSheetName, 126, 7, actRow5List);
				
				
				
				int reportsRow6ListCount = reportsRow6List.size();
				ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow6ListCount;i++)
				{
					String data = reportsRow6List.get(i).getText();
					reportsRow6ListArray.add(data);
				}
				String actRow6List = reportsRow6ListArray.toString();
				String expRow6List = excelReader.getCellData(xlSheetName, 127, 6);
				excelReader.setCellData(xlfile, xlSheetName, 127, 7, actRow6List);
				
				
				
				int reportsRow7ListCount = reportsRow7List.size();
				ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow7ListCount;i++)
				{
					String data = reportsRow7List.get(i).getText();
					reportsRow7ListArray.add(data);
				}
				String actRow7List = reportsRow7ListArray.toString();
				String expRow7List = excelReader.getCellData(xlSheetName, 128, 6);
				excelReader.setCellData(xlfile, xlSheetName, 128, 7, actRow7List);
				
				
				
				int reportsRow8ListCount = reportsRow8List.size();
				ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow8ListCount;i++)
				{
					String data = reportsRow8List.get(i).getText();
					reportsRow8ListArray.add(data);
				}
				String actRow8List = reportsRow8ListArray.toString();
				String expRow8List = excelReader.getCellData(xlSheetName, 129, 6);
				excelReader.setCellData(xlfile, xlSheetName, 129, 7, actRow8List);
				
				
				
				int reportsRow9ListCount = reportsRow9List.size();
				ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow9ListCount;i++)
				{
					String data = reportsRow9List.get(i).getText();
					reportsRow9ListArray.add(data);
				}
				String actRow9List = reportsRow9ListArray.toString();
				String expRow9List = excelReader.getCellData(xlSheetName, 130, 6);
				excelReader.setCellData(xlfile, xlSheetName, 130, 7, actRow9List);
				
				

				int reportsRow10ListCount = reportsRow10List.size();
				ArrayList<String> reportsRow10ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow10ListCount;i++)
				{
					String data = reportsRow10List.get(i).getText();
					reportsRow10ListArray.add(data);
				}
				String actRow10List = reportsRow10ListArray.toString();
				String expRow10List = excelReader.getCellData(xlSheetName, 131, 6);
				excelReader.setCellData(xlfile, xlSheetName, 131, 7, actRow10List);
				
				
				
				int reportsRow11ListCount = reportsRow11List.size();
				ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow11ListCount;i++)
				{
					String data = reportsRow11List.get(i).getText();
					reportsRow11ListArray.add(data);
				}
				String actRow11List = reportsRow11ListArray.toString();
				String expRow11List = excelReader.getCellData(xlSheetName, 132, 6);
				excelReader.setCellData(xlfile, xlSheetName, 132, 7, actRow11List);
				
				
				/*int reportsRow12ListCount = reportsRow12List.size();
				ArrayList<String> reportsRow12ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow12ListCount;i++)
				{
					String data = reportsRow12List.get(i).getText();
					reportsRow12ListArray.add(data);
				}
				String actRow12List = reportsRow12ListArray.toString();
				String expRow12List = "[Grand Total, , 66.00, 38.41, 66.00, 22.23, , 405.11, 170.45, 318.66, 53.23, , , , ]";
				*/
			
				
				System.out.println("************************************checkStockLedgerCustomizationReport********************************************");
			
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				System.out.println(actRow4List);
				System.out.println(expRow4List);
				
				System.out.println(actRow5List);
				System.out.println(expRow5List);
				
				System.out.println(actRow6List);
				System.out.println(expRow6List);
				
				System.out.println(actRow7List);
				System.out.println(expRow7List);
				
				System.out.println(actRow8List);
				System.out.println(expRow8List);
				
				System.out.println(actRow9List);
				System.out.println(expRow9List);
				
				System.out.println(actRow10List);
				System.out.println(expRow10List);
				
				System.out.println(actRow11List);
				System.out.println(expRow11List);
				
				/*System.out.println(actRow12List);
				System.out.println(expRow12List);*/
				
				Thread.sleep(2000);
			//delete  layout	
			
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_ReportCustomizeBtn));
				report_ReportCustomizeBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CustomizationDeleteLayoutBtn));
				report_CustomizationDeleteLayoutBtn.click();
			
				getWaitForAlert();
				getAlert().accept();
				
				String expDeleteMessage =excelReader.getCellData(xlSheetName, 133, 6) ;
			    
			    String actDeleteMessage = checkValidationMessage(expDeleteMessage);
			    excelReader.setCellData(xlfile, xlSheetName, 133, 7, actDeleteMessage);
			    
			  
			  //  Thread.sleep(4000);
				
			/*	int count=Listcolumnnames.size();
				for(int i=0;i<count;i++)
				{
					String data=Listcolumnnames.get(i).getText();
					if(data.equals("Warehouse Name"))
					{
						 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseColumn));
						    warehouseColumn.click();
						   Thread.sleep(1000);
					}
				}
			   
			    
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(removeColumn));
			    removeColumn.click();
			    */
			    Thread.sleep(3000);
			   
			   /* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeSaveBtn));
				osr_customizeSaveBtn.click();
				
				String expMessageAfterColumnRemove = "Data saved successfully";
			    
			    String actMessageAfterColumnRemove = checkValidationMessage(expMessageAfterColumnRemove);
				
				Thread.sleep(4000);
				*/
			    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
			    
			    int reportsHeaderListCount = reportsHeaderList.size();
				ArrayList<String> reportsHeaderListArray = new ArrayList<String>();
				for(int i=0;i<reportsHeaderListCount;i++)
				{
					String data = reportsHeaderList.get(i).getText();
					reportsHeaderListArray.add(data);
				}
				String actHeaderList = reportsHeaderListArray.toString();
				String expHeaderList = excelReader.getCellData(xlSheetName, 134, 6);
				excelReader.setCellData(xlfile, xlSheetName, 134, 7, actHeaderList);
				
				
			    System.out.println(actHeaderList);
			    System.out.println(expHeaderList);
			    
				if(actMessage.equalsIgnoreCase(expMessage)
					&& actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
					&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
					&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List) /*&& actRow12List.equalsIgnoreCase(expRow12List)*/ 
				/*&& actMessageAfterColumnRemove.equalsIgnoreCase(expMessageAfterColumnRemove)*/
					&& actHeaderList.equalsIgnoreCase(expHeaderList))
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 121, 8, resPass);
					return true;
				}
				else
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 121, 8, resFail);
					return false;
				}
			}
				

			
			@FindBy(xpath="//*[@id='dvReportInputs']/div[6]/label/span")
			private static WebElement sl_MovedItemsOnlyChkBox;
			
			
			
			
			
			public boolean checkStockLedgerReportWithMovedItemsOnlyAndIncludeServiceTypeItemsCheckBoxs() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				 excelReader=new ExcelReader(POJOUtility.getExcelPath());

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger));
				stockLedger.click();
				Thread.sleep(2000);
				int rowcount=stockLedgerHometableRowCount.size();
				
				System.out.println(rowcount);
				
				for (int i = 0; i < rowcount; i++) 
				{
					String actName = stockLedgerHometableItemNamesList.get(i).getText();
					
					System.out.println(actName);
					
					if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 136, 5)))
					{
						stockLedgerHometableItemChkboxList.get(i).click();
						break;
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				Thread.sleep(2000);
				
				getAction().moveToElement(sl_MovedItemsOnlyChkBox).build().perform();
				//Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_MovedItemsOnlyChkBox));
				sl_MovedItemsOnlyChkBox.click();
				Thread.sleep(2000);
				
				getAction().moveToElement(sl_IncludeServiceTypeItemChkBox).build().perform();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger_IncludeServiceTypeItemChkBox));
				stockLedger_IncludeServiceTypeItemChkBox.click();
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);
				
				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List = excelReader.getCellData(xlSheetName, 136, 6);
				
				excelReader.setCellData(xlfile, xlSheetName, 136, 7, actRow1List);
				

				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List = excelReader.getCellData(xlSheetName, 137, 6);
				
				excelReader.setCellData(xlfile, xlSheetName, 137, 7, actRow2List);				
				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List = excelReader.getCellData(xlSheetName, 138, 6);
				excelReader.setCellData(xlfile, xlSheetName, 138, 7, actRow3List);

				
				
			
				int reportsRow4ListCount = reportsRow4List.size();
				ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow4ListCount;i++)
				{
					String data = reportsRow4List.get(i).getText();
					reportsRow4ListArray.add(data);
				}
				String actRow4List = reportsRow4ListArray.toString();
				String expRow4List = excelReader.getCellData(xlSheetName, 139, 6);
				excelReader.setCellData(xlfile, xlSheetName, 139, 7, actRow4List);

				
				
				
				int reportsRow5ListCount = reportsRow5List.size();
				ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow5ListCount;i++)
				{
					String data = reportsRow5List.get(i).getText();
					reportsRow5ListArray.add(data);
				}
				String actRow5List = reportsRow5ListArray.toString();
				String expRow5List = excelReader.getCellData(xlSheetName, 140, 6);
				excelReader.setCellData(xlfile, xlSheetName, 140, 7, actRow5List);

				
				
				
				int reportsRow6ListCount = reportsRow6List.size();
				ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow6ListCount;i++)
				{
					String data = reportsRow6List.get(i).getText();
					reportsRow6ListArray.add(data);
				}
				String actRow6List = reportsRow6ListArray.toString();
				String expRow6List = excelReader.getCellData(xlSheetName, 141, 6);
				excelReader.setCellData(xlfile, xlSheetName, 141, 7, actRow6List);

				
				
				
				int reportsRow7ListCount = reportsRow7List.size();
				ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow7ListCount;i++)
				{
					String data = reportsRow7List.get(i).getText();
					reportsRow7ListArray.add(data);
				}
				String actRow7List = reportsRow7ListArray.toString();
				String expRow7List = excelReader.getCellData(xlSheetName, 142, 6);
				excelReader.setCellData(xlfile, xlSheetName, 142, 7, actRow7List);

				
				
				
				int reportsRow8ListCount = reportsRow8List.size();
				ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow8ListCount;i++)
				{
					String data = reportsRow8List.get(i).getText();
					reportsRow8ListArray.add(data);
				}
				String actRow8List = reportsRow8ListArray.toString();
				String expRow8List = excelReader.getCellData(xlSheetName, 143, 6);
				excelReader.setCellData(xlfile, xlSheetName, 143, 7, actRow8List);

				
				
				
				int reportsRow9ListCount = reportsRow9List.size();
				ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow9ListCount;i++)
				{
					String data = reportsRow9List.get(i).getText();
					reportsRow9ListArray.add(data);
				}
				String actRow9List = reportsRow9ListArray.toString();
				String expRow9List = excelReader.getCellData(xlSheetName, 144, 6);
				excelReader.setCellData(xlfile, xlSheetName, 144, 7, actRow9List);

				
				

				int reportsRow10ListCount = reportsRow10List.size();
				ArrayList<String> reportsRow10ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow10ListCount;i++)
				{
					String data = reportsRow10List.get(i).getText();
					reportsRow10ListArray.add(data);
				}
				String actRow10List = reportsRow10ListArray.toString();
				String expRow10List = excelReader.getCellData(xlSheetName, 145, 6);
				excelReader.setCellData(xlfile, xlSheetName, 145, 7, actRow10List);

				
				
				
				int reportsRow11ListCount = reportsRow11List.size();
				ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow11ListCount;i++)
				{
					String data = reportsRow11List.get(i).getText();
					reportsRow11ListArray.add(data);
				}
				String actRow11List = reportsRow11ListArray.toString();
				String expRow11List = excelReader.getCellData(xlSheetName, 146, 6);
				excelReader.setCellData(xlfile, xlSheetName, 146, 7, actRow11List);

			
				
				/*
				int reportsRow12ListCount = reportsRow12List.size();
				ArrayList<String> reportsRow12ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow12ListCount;i++)
				{
					String data = reportsRow12List.get(i).getText();
					reportsRow12ListArray.add(data);
				}
				String actRow12List = reportsRow12ListArray.toString();
				String expRow12List = "[Grand Total, , 66.00, 38.41, 66.00, 22.23, , 405.11, 170.45, 318.66, 53.23, , , ]";
				*/
				
				System.out.println("************************checkStockLedgerReportWithMovedItemsOnlyAndIncludeServiceTypeItemsCheckBoxs********************************");
			
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				System.out.println(actRow4List);
				System.out.println(expRow4List);
				
				System.out.println(actRow5List);
				System.out.println(expRow5List);
				
				System.out.println(actRow6List);
				System.out.println(expRow6List);
				
				System.out.println(actRow7List);
				System.out.println(expRow7List);
				
				System.out.println(actRow8List);
				System.out.println(expRow8List);
				
				System.out.println(actRow9List);
				System.out.println(expRow9List);
				
				System.out.println(actRow10List);
				System.out.println(expRow10List);
				
				System.out.println(actRow11List);
				System.out.println(expRow11List);
				
				/*System.out.println(actRow12List);
				System.out.println(expRow12List);*/
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
					&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
					&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List) /*&& actRow12List.equalsIgnoreCase(expRow12List)*/)
				{
					excelReader.setCellData(xlfile, xlSheetName, 136, 8, resPass);
					return true;
				}
				else
				{
					excelReader.setCellData(xlfile, xlSheetName, 136, 8, resFail);

					return false;
				}
			}

			
		/*	@FindBy (xpath="/html/body/section/div[2]/div[3]/div[1]/div[3]/div[1]/div[1]/div/nav/div/div/ul/li[7]/a/i")
			private static WebElement sl_BackTrackBtn;*/
			
			@FindBy (xpath="//*[@id='id_reportmenudisplay']/ul/li[7]")
			private static WebElement sl_BackTrackBtn;
			
			
			@FindBy (xpath="(//*[@id='navbarSupportedContent2']/ul/li/a/i)[3]")
			private static WebElement BackTrackItemDetailsBackBtn;
			
			@FindBy (xpath="(//*[@id='navbarSupportedContent2']/ul/li/a/i)[2]")
			private static WebElement BackTrackMonthDetailsBackBtn;
			
			@FindBy (xpath="(//*[@id='navbarSupportedContent2']/ul/li/a/i)[1]")
			private static WebElement BackTrackDateDetailsBackBtn;
			
			
			
			@FindBy (xpath="//*[@id='dvReportDetails']/div/table/tbody/tr/td[3]")
			private static List<WebElement> reportTableColumn2RowsList;
			
			public boolean checkStockLedgerReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				 excelReader=new ExcelReader(POJOUtility.getExcelPath());
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger));
				stockLedger.click();
				int rowcount=stockLedgerHometableRowCount.size();
				
				System.out.println(rowcount);
				
				for (int i = 0; i < rowcount; i++) 
				{
					String actName = stockLedgerHometableItemNamesList.get(i).getText();
					Thread.sleep(1000);
					
					System.out.println(actName);
					
					if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 148, 5)))
					{
						stockLedgerHometableItemChkboxList.get(i).click();
						break;
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);

				int reportItemGridListCount=reportTableColumn2RowsList.size();
				System.err.println(reportItemGridListCount);
			    for (int i = 0; i < reportItemGridListCount; i++) 
			    {
		             String data=reportTableColumn2RowsList.get(i).getText();
		             
		             System.err.println(data);
		             
		             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 148, 6)))
		             {
		            	 reportTableColumn2RowsList.get(i).click();
	            		 break;
	                 }
			    }
				
			    
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
				sl_BackTrackBtn.click();
				
				Thread.sleep(2000);
				
				ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
				int actOpenWindowsCount = getDriver().getWindowHandles().size();
				int expOpenWindowsCount = 2;

			 	getDriver().switchTo().window(openTabs.get(1));
			 	
			 	Thread.sleep(2000);
			 	
			 	checkValidationMessage("");
			 	
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
				String actDocNo = documentNumberTxt.getAttribute("value");
				String expDocNo =  excelReader.getCellData(xlSheetName, 149, 6);
				excelReader.setCellData(xlfile, xlSheetName, 149, 7, actDocNo);


				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
				String actVendor = vendorAccountTxt.getAttribute("value");
				String expVendor = excelReader.getCellData(xlSheetName, 150, 6);
				excelReader.setCellData(xlfile, xlSheetName, 150, 7, actVendor);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
				String actCurrency = voucherHeaderCurrency.getAttribute("value");
				String expCurrency = excelReader.getCellData(xlSheetName, 151, 6);
				excelReader.setCellData(xlfile, xlSheetName, 151, 7, actCurrency);

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
				String actDepartment = departmentTxt.getAttribute("value");
				String expDepartment = excelReader.getCellData(xlSheetName, 152, 6);
				excelReader.setCellData(xlfile, xlSheetName, 152, 7, actDepartment);

				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
				String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
				String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 153, 6);
				excelReader.setCellData(xlfile, xlSheetName, 153, 7, actPlaceOfSupply);


				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
				String actJurisdiction = jurisdictionTxt.getAttribute("value");
				String expJurisdiction = excelReader.getCellData(xlSheetName, 154, 6);
				excelReader.setCellData(xlfile, xlSheetName, 154, 7, actJurisdiction);


				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				String actR1Warehouse = select1stRow_1stColumn.getText();
				String expR1Warehouse = excelReader.getCellData(xlSheetName, 155, 6);
				excelReader.setCellData(xlfile, xlSheetName, 155, 7, actR1Warehouse);

				
				String actR1Item = select1stRow_2ndColumn.getText();
				String expR1Item = excelReader.getCellData(xlSheetName, 156, 6);
				excelReader.setCellData(xlfile, xlSheetName, 156, 7, actR1Item);

				
				String actR1TaxCode = select1stRow_3rdColumn.getText();
				String expR1TaxCode = excelReader.getCellData(xlSheetName, 157, 6);
				excelReader.setCellData(xlfile, xlSheetName, 157, 7, actR1TaxCode);

				
				String actR1PurchaseAccount = select1stRow_4thColumn.getText();
				String expR1PurchaseAccount = excelReader.getCellData(xlSheetName, 158, 6);
				excelReader.setCellData(xlfile, xlSheetName, 158, 7, actR1PurchaseAccount);

				
				String actR1Units = select1stRow_5thColumn.getText();
				String expR1Units = excelReader.getCellData(xlSheetName, 159, 6);
				excelReader.setCellData(xlfile, xlSheetName, 159, 7, actR1Units);

				
				String actR1Quantity = select1stRow_9thColumn.getText();
				String expR1Quantity = excelReader.getCellData(xlSheetName, 160, 6);
				excelReader.setCellData(xlfile, xlSheetName, 160, 7, actR1Quantity);

				
				String actR1Rate = select1stRow_11thColumn.getText();
				String expR1Rate = excelReader.getCellData(xlSheetName, 161, 6);
				excelReader.setCellData(xlfile, xlSheetName, 161, 7, actR1Rate);

				
				String actR1Gross = select1stRow_12thColumn.getText();
				String expR1Gross = excelReader.getCellData(xlSheetName, 162, 6);
				excelReader.setCellData(xlfile, xlSheetName, 162, 7, actR1Gross);

				
				String actR1RMA = select1stRow_19thColumn.getText();
				String expR1RMA = "";

				
				
			/*	
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
				String actR2Warehouse = select2ndRow_1stColumn.getText();
				String expR2Warehouse = "HYDERABAD";
				
				String actR2Item = select2ndRow_2ndColumn.getText();
				String expR2Item = "WA COGS ITEM";
				
				String actR2TaxCode = select2ndRow_3rdColumn.getText();
				String expR2TaxCode = "Standard Rated Purchase - Recoverable";
				
				String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
				String expR2PurchaseAccount = "WA COGS ACC INV";
				
				String actR2Units = select2ndRow_5thColumn.getText();
				String expR2Units = "Pcs";
				
				String actR2Quantity = select2ndRow_9thColumn.getText();
				String expR2Quantity = "10.00";
				
				String actR2Rate = select2ndRow_11thColumn.getText();
				String expR2Rate = "10.00";
				
				String actR2Gross = select2ndRow_12thColumn.getText();
				String expR2Gross = "100.00";
				
				String actR2RMA = select2ndRow_19thColumn.getText();
				String expR2RMA = "RMA#R2PVVAT1,RMA#R2PVVAT2,RMA#R2PVVAT3,RMA#R2PVVAT4,RMA#R2PVVAT5,RMA#R2PVVAT6,RMA#R2PVVAT7,RMA#R2PVVAT8,RMA#R2PVVAT9,RMA#R2PVVAT10";
*/
				System.out.println("*******************************checkStockLedgerReportBackTrackOption*********************************");
				
				System.out.println("OpenWindowsCount: "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
				
				System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
				System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
				System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
				System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
				System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
				System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);
				
				System.out.println("*******************************ROW1**********************************");
				
				System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
				System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
				System.out.println("TaxCode         : "+actR1TaxCode			+"  Value Expected  "+expR1TaxCode);
				System.out.println("PurchaseAccount : "+actR1PurchaseAccount	+"  Value Expected  "+expR1PurchaseAccount);
				System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
				System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
				System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
				System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
				System.out.println("RMA             : "+actR1RMA				+"  Value Expected  "+expR1RMA);
				
			/*	System.out.println("*******************************ROW2**********************************");
				
				System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
				System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
				System.out.println("TaxCode         : "+actR2TaxCode			+"  Value Expected  "+expR2TaxCode);
				System.out.println("PurchaseAccount : "+actR2PurchaseAccount	+"  Value Expected  "+expR2PurchaseAccount);
				System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
				System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
				System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
				System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
				System.out.println("RMA             : "+actR2RMA				+"  Value Expected  "+expR2RMA);
				*/
				
				getDriver().switchTo().window(openTabs.get(1)).close();
				
				Thread.sleep(1000);
				
				getDriver().switchTo().window(openTabs.get(0));
				
				if(actOpenWindowsCount==expOpenWindowsCount 
					&& actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
					&& actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
					&& actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
					&& actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actR1PurchaseAccount.equalsIgnoreCase(expR1PurchaseAccount) 
					&& actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) 
					&& actR1Gross.equalsIgnoreCase(expR1Gross) && actR1RMA.equalsIgnoreCase(expR1RMA) 
				/*	&& actR2Warehouse.equalsIgnoreCase(expR2Warehouse) 
					&& actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
					&& actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
					&& actR2Gross.equalsIgnoreCase(expR2Gross) && actR2RMA.equalsIgnoreCase(expR2RMA)*/)
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 148, 8, resPass);
					return true;
				}
				else
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 148, 8, resFail);

					return false;
				}
			}
			
			
			@FindBy(xpath="//*[@id='552']")
			private static WebElement openingStocksRegister;
			

			@FindBy(xpath="//*[@id='id_reportmenudisplay']/ul/li[13]")
			private static WebElement osr_CloseBtn;
			public boolean checkOpeningStockRegisterReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
                excelReader=new ExcelReader(POJOUtility.getExcelPath());
                LogoutandLoginwithSU();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksRegister));
				openingStocksRegister.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);
				
				//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
				
				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List = excelReader.getCellData(xlSheetName, 164, 6);

                excelReader.setCellData(xlfile, xlSheetName, 164, 7, actRow1List);

				
				

				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List = excelReader.getCellData(xlSheetName, 165, 6);

                excelReader.setCellData(xlfile, xlSheetName, 165, 7, actRow2List);

			

				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List = excelReader.getCellData(xlSheetName, 166, 6);

                excelReader.setCellData(xlfile, xlSheetName, 166, 7, actRow3List);

				
				
				
				int reportsRow4ListCount = reportsRow4List.size();
				ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow4ListCount;i++)
				{
					String data = reportsRow4List.get(i).getText();
					reportsRow4ListArray.add(data);
				}
				String actRow4List = reportsRow4ListArray.toString();
				String expRow4List = excelReader.getCellData(xlSheetName, 167, 6);

                excelReader.setCellData(xlfile, xlSheetName, 167, 7, actRow4List);

				
				
				
				int reportsRow5ListCount = reportsRow5List.size();
				ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow5ListCount;i++)
				{
					String data = reportsRow5List.get(i).getText();
					reportsRow5ListArray.add(data);
				}
				String actRow5List = reportsRow5ListArray.toString();
				String expRow5List = excelReader.getCellData(xlSheetName, 168, 6);

                excelReader.setCellData(xlfile, xlSheetName, 168, 7, actRow5List);

				
				
				
				int reportsRow6ListCount = reportsRow6List.size();
				ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow6ListCount;i++)
				{
					String data = reportsRow6List.get(i).getText();
					reportsRow6ListArray.add(data);
				}
				String actRow6List = reportsRow6ListArray.toString();
				String expRow6List = excelReader.getCellData(xlSheetName, 169, 6);

                excelReader.setCellData(xlfile, xlSheetName, 169, 7, actRow6List);

				
				
				
				int reportsRow7ListCount = reportsRow7List.size();
				ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow7ListCount;i++)
				{
					String data = reportsRow7List.get(i).getText();
					reportsRow7ListArray.add(data);
				}
				String actRow7List = reportsRow7ListArray.toString();
				String expRow7List = excelReader.getCellData(xlSheetName, 170, 6);

                excelReader.setCellData(xlfile, xlSheetName, 170, 7, actRow7List);

				
				
				
				int reportsRow8ListCount = reportsRow8List.size();
				ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow8ListCount;i++)
				{
					String data = reportsRow8List.get(i).getText();
					reportsRow8ListArray.add(data);
				}
				String actRow8List = reportsRow8ListArray.toString();
				String expRow8List = excelReader.getCellData(xlSheetName, 171, 6);

                excelReader.setCellData(xlfile, xlSheetName, 171, 7, actRow8List);

				
				
				int reportsRow9ListCount = reportsRow9List.size();
				ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow9ListCount;i++)
				{
					String data = reportsRow9List.get(i).getText();
					reportsRow9ListArray.add(data);
				}
				String actRow9List = reportsRow9ListArray.toString();
				String expRow9List = excelReader.getCellData(xlSheetName, 172, 6);

                excelReader.setCellData(xlfile, xlSheetName, 172, 7, actRow9List);

				
				
				int reportsRow10ListCount = reportsRow10List.size();
				ArrayList<String> reportsRow10ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow10ListCount;i++)
				{
					String data = reportsRow10List.get(i).getText();
					reportsRow10ListArray.add(data);
				}
				String actRow10List = reportsRow10ListArray.toString();
				String expRow10List = excelReader.getCellData(xlSheetName, 173, 6);

                excelReader.setCellData(xlfile, xlSheetName, 173, 7, actRow10List);

				
				
				System.out.println("***************************checkOpenigStockRegisterReport*********************************");
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				System.out.println(actRow4List);
				System.out.println(expRow4List);
				
				System.out.println(actRow5List);
				System.out.println(expRow5List);
				
				System.out.println(actRow6List);
				System.out.println(expRow6List);
				
				System.out.println(actRow7List);
				System.out.println(expRow7List);
				
				System.out.println(actRow8List);
				System.out.println(expRow8List);
				
				System.out.println(actRow9List);
				System.out.println(expRow9List);
				
				System.out.println(actRow10List);
				System.out.println(expRow10List);
				
				Thread.sleep(1000);
				osr_CloseBtn.click();
				
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
						&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List)
						&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List)
						&& actRow10List.equalsIgnoreCase(expRow10List))
				{
					excelReader.setCellData(xlfile, xlSheetName, 164, 8, resPass);
					return true;
				}
				else
				{
					excelReader.setCellData(xlfile, xlSheetName, 164, 8, resFail);

					if (actRow10List.equalsIgnoreCase(expRow10List)) 
					{
						return true;
					}
					else 
					{
						return false;
					}
				}
			}


			
			@FindBy(xpath="//*[@id='idFilterCustomizeIcon']")
			private static WebElement  report_FilterCustomizeBtn;
			
			@FindBy(xpath="//a[contains(text(),'Item')]//i")
			private static WebElement  report_FilterItemExpansion;
			
			@FindBy(xpath="(//*[@id='5021'])[1]")
			private static WebElement  report_FilterItemNameChkbox;
			
			@FindBy(xpath="(//input[@class='FButton-Primary'])[1]")
			private static WebElement  reports_Filter_Customize_OkBtn;
			
			@FindBy(xpath="//input[@id='FOption_552_0_DefaultFilter_0']")
			private static WebElement  osr_DefaultFilterTxt;
			
			@FindBy(xpath="//span[@id='filterRefresh']")
			private static WebElement  report_Filter_RefreshBtn;
			
			@FindBy(xpath="//*[@id='filterRefresh']")
			private static WebElement  report_Filter_RefreshConditionBtn;
			
			
			
			@FindBy(xpath="//*[@id='filter_Okbtn_']")
			private static WebElement  report_FilterOkBtn;
			
			@FindBy(xpath="//i[contains(@class,'icon icon-close')]")
			private static WebElement  report_FilterCancelBtn;
			
			
			@FindBy(xpath="(//table[@class='CommonReportTable']/tbody)[1]/tr")
			private static List<WebElement> reportsRowsCount;
			
			@FindBy(xpath="//*[@id='FilterFields_552_0']/ul[2]/li[1]/div/label/span")
			private static WebElement osr_report_FilterItemNameChkbox;
			
			@FindBy(xpath="//*[@id='FilterFieldCust_552_0']/div/div[3]/input[1]")
			private static WebElement osr_reports_Filter_Customize_OkBtn;
			
			
			public boolean checkOpeningStockRegisterReportFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				 excelReader=new ExcelReader(POJOUtility.getExcelPath());
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksRegister));
				openingStocksRegister.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
				report_FilterExpandBtn.click();
				Thread.sleep(1000);
				
				Thread.sleep(3000);getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
				report_FilterBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
				report_FilterCustomizeBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterItemExpansion));
				report_FilterItemExpansion.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_report_FilterItemNameChkbox));
				
				if (report_FilterItemNameChkbox.isSelected()==false) 
				{
					osr_report_FilterItemNameChkbox.click();
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_reports_Filter_Customize_OkBtn));
				osr_reports_Filter_Customize_OkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_DefaultFilterTxt));
				osr_DefaultFilterTxt.click();
				
				osr_DefaultFilterTxt.sendKeys(excelReader.getCellData(xlSheetName, 175, 5));
				System.err.println(excelReader.getCellData(xlSheetName, 175, 5));
				
				Thread.sleep(2000);
				
				osr_DefaultFilterTxt.sendKeys(Keys.TAB);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
				report_FilterOkBtn.click();
				
				Thread.sleep(3000);
				
				//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
				
				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List =  excelReader.getCellData(xlSheetName, 175, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 175, 7,  actRow1List);
				

				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List =  excelReader.getCellData(xlSheetName, 176, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 176, 7,  actRow2List);
			

				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List =  excelReader.getCellData(xlSheetName, 177, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 177, 7,  actRow3List);
				
				
				int reportsRow4ListCount = reportsRow4List.size();
				ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow4ListCount;i++)
				{
					String data = reportsRow4List.get(i).getText();
					reportsRow4ListArray.add(data);
				}
				String actRow4List = reportsRow4ListArray.toString();
				String expRow4List =  excelReader.getCellData(xlSheetName, 178, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 178, 7,  actRow4List);
				
				
				System.out.println("***************************checkOpeningStockRegisterReportFilterOption*********************************");
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				System.out.println(actRow4List);
				System.out.println(expRow4List);
				
				Thread.sleep(2000);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
				report_FilterExpandBtn.click();
				Thread.sleep(1000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
				report_FilterBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
				report_FilterCustomizeBtn.click();
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_report_FilterItemNameChkbox));
				
				if (osr_report_FilterItemNameChkbox.isSelected()==true) 
				{
					osr_report_FilterItemNameChkbox.click();
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_reports_Filter_Customize_OkBtn));
				osr_reports_Filter_Customize_OkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
				report_Filter_RefreshConditionBtn.click();

				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
				report_FilterOkBtn.click();
				
				Thread.sleep(2000);
				
				int actRowCount = reportsRowsCount.size();
				int  expRowCount =  10;
				System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
						&& actRow4List.equalsIgnoreCase(expRow4List) && actRowCount==expRowCount)
				{
					excelReader.setCellData(xlfile, xlSheetName, 175, 8, resPass);
					return true;
				}
				else
				{
					excelReader.setCellData(xlfile, xlSheetName, 175, 8, resFail);
					return false;
				}
			}

	    	
			
			
			
			@FindBy (xpath="(//div[@id='dvReportDetails']/div/table/tbody)[1]/tr/td[4]")
			private static List<WebElement> reportTableColumn3RowsList;
			
			
			
			public boolean checkOpeningStockRegisterReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				 excelReader=new ExcelReader(POJOUtility.getExcelPath());
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(openingStocksRegister));
				openingStocksRegister.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);int reportItemGridListCount=reportTableColumn2RowsList.size();
			    for (int i = 0; i < reportItemGridListCount; i++) 
			    {
		             String data=reportTableColumn2RowsList.get(i).getText();
		             String data1=reportTableColumn3RowsList.get(i).getText();
		             
		             System.err.println(data+"  "+data1);
		             
		             if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 180, 5))/* && data1.equalsIgnoreCase("12.00")*/)
		             {
		            	 reportTableColumn2RowsList.get(i).click();
	            		 break;
	                 }
			    }
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
				sl_BackTrackBtn.click();
				
				ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
				//int actOpenWindowsCount = getDriver().getWindowHandles().size();
				
				
				String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
				String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 180, 6);
				excelReader.setCellData(xlfile, xlSheetName, 180, 7, actOpenWindowsCount);

				
				/*String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 180, 6);
				excelReader.setCellData(xlfile, xlSheetName, 180, 7, actOpenWindowsCount);
			 	*/
				getDriver().switchTo().window(openTabs.get(1));
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
				String actDocNo=documentNumberTxt.getAttribute("value");
				String expDocNo=excelReader.getCellData(xlSheetName, 181, 6);
				excelReader.setCellData(xlfile, xlSheetName, 181, 7, actDocNo);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				String actWarehouse=warehouseTxt.getAttribute("value");
				String expWarehouse=excelReader.getCellData(xlSheetName, 182, 6);
				excelReader.setCellData(xlfile, xlSheetName, 182, 7, actWarehouse);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				String actItem=select1stRow_1stColumn.getText();
				String expItem=excelReader.getCellData(xlSheetName, 183, 6);
				excelReader.setCellData(xlfile, xlSheetName, 183, 7, actItem);
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
				String actUnits=select1stRow_2ndColumn.getText();
				String expUnits=excelReader.getCellData(xlSheetName, 184, 6);
				excelReader.setCellData(xlfile, xlSheetName, 184, 7, actUnits);
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
				String actQuantity=select1stRow_3rdColumn.getText();
				String expQuantity=excelReader.getCellData(xlSheetName, 185, 6);
				excelReader.setCellData(xlfile, xlSheetName, 185, 7, actQuantity);
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
				String actRate=select1stRow_4thColumn.getText();
				String expRate=excelReader.getCellData(xlSheetName, 186, 6);
				excelReader.setCellData(xlfile, xlSheetName, 186, 7, actRate);
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
				String actGross=select1stRow_5thColumn.getText();
				String expGross=excelReader.getCellData(xlSheetName, 187, 6);
				excelReader.setCellData(xlfile, xlSheetName, 187, 7, actGross);
				
				

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
				String actBatch=select1stRow_6thColumn.getText();
				String expBatch=excelReader.getCellData(xlSheetName, 188, 6);
				excelReader.setCellData(xlfile, xlSheetName, 188, 7, actBatch);
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
				String actExpiryDate=select1stRow_8thColumn.getText();
				actExpiryDate="datefield";
				
				DateFormat df = new SimpleDateFormat("dd/MM/yyyy");
				Calendar cal=Calendar.getInstance();
				cal.add(Calendar.DAY_OF_WEEK, 5); 
				//String expExpiryDate=df.format(cal.getTime());
				String expExpiryDate=excelReader.getCellData(xlSheetName, 189, 6);
				excelReader.setCellData(xlfile, xlSheetName, 189, 7, actExpiryDate);
				

				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
				String actItem2=select2ndRow_1stColumn.getText();
				String expItem2=excelReader.getCellData(xlSheetName, 190, 6);
				excelReader.setCellData(xlfile, xlSheetName, 190, 7, actItem2);
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
				String actUnits2=select2ndRow_2ndColumn.getText();
				String expUnits2=excelReader.getCellData(xlSheetName, 191, 6);
				excelReader.setCellData(xlfile, xlSheetName, 191, 7, actUnits2);
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
				String actQuantity2=select2ndRow_3rdColumn.getText();
				String expQuantity2=excelReader.getCellData(xlSheetName, 192, 6);
				excelReader.setCellData(xlfile, xlSheetName, 192, 7, actQuantity2);
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
				String actRate2=select2ndRow_4thColumn.getText();
				String expRate2=excelReader.getCellData(xlSheetName, 193, 6);
				excelReader.setCellData(xlfile, xlSheetName, 193, 7,  actRate2);
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
				String actGross2=select2ndRow_5thColumn.getText();
				String expGross2=excelReader.getCellData(xlSheetName, 194, 6);
				excelReader.setCellData(xlfile, xlSheetName, 194, 7, actGross2);
				
				

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_6thColumn));
				String actBatch2=select2ndRow_6thColumn.getText();
				String expBatch2=excelReader.getCellData(xlSheetName, 195, 6);
				excelReader.setCellData(xlfile, xlSheetName, 195, 7, actBatch2);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_8thColumn));
				String actExpiryDate2=select2ndRow_8thColumn.getText();
				actExpiryDate2="datefield";
				
				Calendar cal2=Calendar.getInstance();
				cal2.add(Calendar.DAY_OF_WEEK, 7); 
				
				//String expExpiryDate2=df.format(cal2.getTime());
				String expExpiryDate2=excelReader.getCellData(xlSheetName, 196, 6);
				excelReader.setCellData(xlfile, xlSheetName, 196, 7, actExpiryDate2);
				

				
				System.out.println("**********************checkOpeningStockRegisterReportBackTrackOption*********************");
				
				
				System.out.println("*********** OpenWindowsCount : "+actOpenWindowsCount+"  value expected  "+expOpenWindowsCount);
				System.out.println("*********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
				System.out.println("*********** Warehouse  : "+actWarehouse  +"  value expected  "+expWarehouse);
				System.out.println("*********** Item       : "+actItem       +"  value expected  "+expItem);
				System.out.println("*********** Units      : "+actUnits      +"  value expected  "+expUnits);
				System.out.println("*********** Quantity   : "+actQuantity   +"  value expected  "+expQuantity);
				System.out.println("*********** Rate       : "+actRate       +"  value expected  "+expRate);
				System.out.println("*********** Gross      : "+actGross      +"  value expected  "+expGross);
				System.out.println("*********** Batch      : "+actBatch      +"  value expected  "+expBatch);
				System.out.println("*********** ExpiryDate : "+actExpiryDate +"  value expected  "+expExpiryDate);
				System.out.println("*********** Item2       : "+actItem2       +"  value expected  "+expItem2);
				System.out.println("*********** Units2      : "+actUnits2      +"  value expected  "+expUnits2);
				System.out.println("*********** Quantity2   : "+actQuantity2   +"  value expected  "+expQuantity2);
				System.out.println("*********** Rate2       : "+actRate2       +"  value expected  "+expRate2);
				System.out.println("*********** Gross2      : "+actGross2      +"  value expected  "+expGross2);
				System.out.println("*********** Batch2      : "+actBatch2      +"  value expected  "+expBatch2);
				System.out.println("*********** ExpiryDate2 : "+actExpiryDate2 +"  value expected  "+expExpiryDate2);
			 	
				getDriver().switchTo().window(openTabs.get(1)).close();
				
				Thread.sleep(1000);
				
				getDriver().switchTo().window(openTabs.get(0));
				
				if(/*actOpenWindowsCount==expOpenWindowsCount
					 && */actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
					 && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
					 && actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch) && actExpiryDate.equalsIgnoreCase(expExpiryDate)
					 && actItem2.equalsIgnoreCase(expItem2) && actUnits2.equalsIgnoreCase(expUnits2) && actQuantity2.equalsIgnoreCase(expQuantity2) 
					 && actRate2.equalsIgnoreCase(expRate2) && actGross2.equalsIgnoreCase(expGross2) && actBatch2.equalsIgnoreCase(expBatch2) 
					 && actExpiryDate2.equalsIgnoreCase(expExpiryDate2))
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 180, 8, resPass);
					return true;
				}
				else
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 180, 8, resFail);

					return false;
				}
			}
			
			
			@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[5]")
			private static WebElement stockStatement;
			
			public boolean checkStockStatementReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				excelReader=new ExcelReader(POJOUtility.getExcelPath());
				LogoutandLoginwithSU();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockStatement));
				stockStatement.click();
				
				Thread.sleep(2000);
				
				int rowcount=stockLedgerHometableRowCount.size();
				
				System.out.println(rowcount);
				
				for (int i = 1; i <= rowcount; i++) 
				{
					String actName = stockLedgerHometableItemNamesList.get(i).getText();
					
					System.out.println(actName);
					
					if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 198, 5)))
					{
						stockLedgerHometableItemChkboxList.get(i).click();
						break;
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));

				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List = excelReader.getCellData(xlSheetName, 198, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 198, 7, actRow1List);

				

				
				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List = excelReader.getCellData(xlSheetName, 199, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 199, 7, actRow2List);
			

				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List = excelReader.getCellData(xlSheetName, 200, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 200, 7, actRow3List);
				
				
				System.out.println("***************************checkStockStatementReport*********************************");
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				Thread.sleep(2000);
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
				{
					excelReader.setCellData(xlfile, xlSheetName, 198, 8, resPass);

					return true;
				}
				else
				{
					excelReader.setCellData(xlfile, xlSheetName, 198, 8, resFail);

					return false;
				}
			}

			@FindBy (xpath="//*[@id='id_reportmenudisplay']/ul/li[4]")
			private static WebElement sl_ReportPrintBtn;
			
			
			
			
			public boolean checkStockStatementReportPrintOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ReportPrintBtn));
				sl_ReportPrintBtn.click();
				
				//Thread.sleep(3000);
				getWaitForAlert();;
				Thread.sleep(2000);
				
				getAlert().accept();
				Thread.sleep(2000);
				
				ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
				System.err.println("openTabs"+openTabs);
				
				int actOpenWindowsCount = getDriver().getWindowHandles().size();
				int expOpenWindowsCount = 2;

			 	getDriver().switchTo().window(openTabs.get(0));
			 	
			 	Thread.sleep(1000);
			 	
			 	getDriver().switchTo().window(openTabs.get(1)).close();
			 	
			 	getDriver().switchTo().window(openTabs.get(0));
				
				System.out.println("***************************checkStockStatementReportPrintOption*********************************");
				
				System.err.println("openTabs"+openTabs);
				System.err.println("Open Tabs Count : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
				
				Thread.sleep(1000);
				
				if(actOpenWindowsCount==expOpenWindowsCount)
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					return true;
				}
				else
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					return false;
				}
			}

			
			public boolean checkStockStatementReportWithIncludeServiceTypeItem() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				excelReader=new ExcelReader(POJOUtility.getExcelPath());				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockStatement));
				stockStatement.click();
				
				Thread.sleep(2000);
				
				int rowcount=stockLedgerHometableRowCount.size();
				
				System.out.println(rowcount);
				
				for (int i = 1; i <= rowcount; i++) 
				{
					String actName = stockLedgerHometableItemNamesList.get(i).getText();
					
					System.out.println(actName);
					
					if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 206, 5)))
					{
						stockLedgerHometableItemChkboxList.get(i).click();
						break;
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				
				Thread.sleep(2000);
				
				getAction().moveToElement(sl_IncludeServiceTypeItemChkBox).build().perform();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_IncludeServiceTypeItemChkBox));
				sl_IncludeServiceTypeItemChkBox.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));

				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List = excelReader.getCellData(xlSheetName, 206, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 206, 7, actRow1List);

				
				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=2;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List = excelReader.getCellData(xlSheetName, 207, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 207, 7, actRow2List);
			

				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List = excelReader.getCellData(xlSheetName, 208, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 208, 7, actRow3List);
				
				
				System.out.println("***************************checkStockStatementReport*********************************");
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				Thread.sleep(2000);
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 206, 8, resPass);
					return true;
				}
				else
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 206, 8, resFail);
					return false;
				}
			}

			
			@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[6]")
			private static WebElement stockMovement;

			public boolean checkStockMovementReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				excelReader=new ExcelReader(POJOUtility.getExcelPath());
				getDriver().navigate().refresh();
				Thread.sleep(2000);
				
				LogoutandLoginwithSU();
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockMovement));
				stockMovement.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_HeaderSelectChkBox));
				sl_HeaderSelectChkBox.click();
				
				Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
				
				int reportsByWarehouseRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsByWarehouseRow1ListArray.add(data);
				}
				String actRow1List = reportsByWarehouseRow1ListArray.toString();
				String expRow1List = excelReader.getCellData(xlSheetName, 211, 6);
				excelReader.setCellData(xlfile, xlSheetName, 211, 7, actRow1List);
				

				int reportsByWarehouseRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
				for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsByWarehouseRow2ListArray.add(data);
				}
				String actRow2List = reportsByWarehouseRow2ListArray.toString();
				String expRow2List =  excelReader.getCellData(xlSheetName, 212, 6);
				excelReader.setCellData(xlfile, xlSheetName, 212, 7, actRow2List);
				
				
				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List =  excelReader.getCellData(xlSheetName, 213, 6);
				excelReader.setCellData(xlfile, xlSheetName, 213, 7, actRow3List);
				
				
			
				int reportsRow4ListCount = reportsRow4List.size();
				ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow4ListCount;i++)
				{
					String data = reportsRow4List.get(i).getText();
					reportsRow4ListArray.add(data);
				}
				String actRow4List = reportsRow4ListArray.toString();
				String expRow4List =  excelReader.getCellData(xlSheetName, 214, 6);
				excelReader.setCellData(xlfile, xlSheetName, 214, 7, actRow4List);
				
				
				
				int reportsRow5ListCount = reportsRow5List.size();
				ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow5ListCount;i++)
				{
					String data = reportsRow5List.get(i).getText();
					reportsRow5ListArray.add(data);
				}
				String actRow5List = reportsRow5ListArray.toString();
				String expRow5List =  excelReader.getCellData(xlSheetName, 215, 6);
				excelReader.setCellData(xlfile, xlSheetName, 215, 7, actRow5List);
				
				System.out.println("*********************************checkStockMovementReport*****************************************");
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				System.out.println(actRow4List);
				System.out.println(expRow4List);
				
				System.out.println(actRow5List);
				System.out.println(expRow5List);
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					 excelReader.setCellData(xlfile, xlSheetName, 211, 8, resPass);
					return true;
				}
				else
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					 excelReader.setCellData(xlfile, xlSheetName, 211, 8, resFail);
					return false;
				}
			}
			
			
			
			
			
			@FindBy (xpath="//select[@id='RITCombobox__1']")
			private static WebElement stockMovementDropdown;
			
			
			
			public boolean checkStockMovementReportsWithStockMovementByWarehouseAndIncludeServiceTypeItemOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
		       //Thread.sleep(2000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockMovementDropdown));
				
				Select sm=new Select(stockMovementDropdown);
				sm.selectByIndex(1);
				Thread.sleep(2000);
				
				getAction().moveToElement(sl_IncludeServiceTypeItemChkBox).build().perform();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_IncludeServiceTypeItemChkBox));
				sl_IncludeServiceTypeItemChkBox.click();
				
				Thread.sleep(1000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
				
				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List = excelReader.getCellData(xlSheetName, 218, 6);
				excelReader.setCellData(xlfile, xlSheetName, 218, 7, actRow1List);
				
				

				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List = excelReader.getCellData(xlSheetName, 219, 6);
				excelReader.setCellData(xlfile, xlSheetName, 219, 7, actRow2List);
				
				
				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List = excelReader.getCellData(xlSheetName, 220, 6);
				excelReader.setCellData(xlfile, xlSheetName, 220, 7, actRow3List);
				
				
				
				int reportsRow4ListCount = reportsRow4List.size();
				ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow4ListCount;i++)
				{
					String data = reportsRow4List.get(i).getText();
					reportsRow4ListArray.add(data);
				}
				String actRow4List = reportsRow4ListArray.toString();
				String expRow4List = excelReader.getCellData(xlSheetName, 221, 6);
				excelReader.setCellData(xlfile, xlSheetName, 221, 7, actRow4List);
				
				
				
				int reportsRow5ListCount = reportsRow5List.size();
				ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow5ListCount;i++)
				{
					String data = reportsRow5List.get(i).getText();
					reportsRow5ListArray.add(data);
				}
				String actRow5List = reportsRow5ListArray.toString();
				String expRow5List = excelReader.getCellData(xlSheetName, 222, 6);
				excelReader.setCellData(xlfile, xlSheetName, 222, 7, actRow5List);
				
				
				
				int reportsRow6ListCount = reportsRow6List.size();
				ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow6ListCount;i++)
				{
					String data = reportsRow6List.get(i).getText();
					reportsRow6ListArray.add(data);
				}
				String actRow6List = reportsRow6ListArray.toString();
				String expRow6List = excelReader.getCellData(xlSheetName, 223, 6);
				excelReader.setCellData(xlfile, xlSheetName, 223, 7, actRow6List);
				
				
				
				int reportsRow7ListCount = reportsRow7List.size();
				ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow7ListCount;i++)
				{
					String data = reportsRow7List.get(i).getText();
					reportsRow7ListArray.add(data);
				}
				String actRow7List = reportsRow7ListArray.toString();
				String expRow7List = excelReader.getCellData(xlSheetName, 224, 6);
				excelReader.setCellData(xlfile, xlSheetName, 224, 7, actRow7List);
				
				
				
				int reportsRow8ListCount = reportsRow8List.size();
				ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow8ListCount;i++)
				{
					String data = reportsRow8List.get(i).getText();
					reportsRow8ListArray.add(data);
				}
				String actRow8List = reportsRow8ListArray.toString();
				String expRow8List = excelReader.getCellData(xlSheetName, 225, 6);
				excelReader.setCellData(xlfile, xlSheetName, 225, 7, actRow8List);
				
				
				
				int reportsRow9ListCount = reportsRow9List.size();
				ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow9ListCount;i++)
				{
					String data = reportsRow9List.get(i).getText();
					reportsRow9ListArray.add(data);
				}
				String actRow9List = reportsRow9ListArray.toString();
				String expRow9List = excelReader.getCellData(xlSheetName, 226, 6);
				excelReader.setCellData(xlfile, xlSheetName, 226, 7, actRow9List);
				
				
				
				int reportsRow10ListCount = reportsRow10List.size();
				ArrayList<String> reportsRow10ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow10ListCount;i++)
				{
					String data = reportsRow10List.get(i).getText();
					reportsRow10ListArray.add(data);
				}
				String actRow10List = reportsRow10ListArray.toString();
				String expRow10List = excelReader.getCellData(xlSheetName, 227, 6);
				excelReader.setCellData(xlfile, xlSheetName, 227, 7, actRow10List);
				
				
				
				int reportsRow11ListCount = reportsRow11List.size();
				ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow11ListCount;i++)
				{
					String data = reportsRow11List.get(i).getText();
					reportsRow11ListArray.add(data);
				}
				String actRow11List = reportsRow11ListArray.toString();
				String expRow11List = excelReader.getCellData(xlSheetName, 228, 6);
				excelReader.setCellData(xlfile, xlSheetName, 228, 7, actRow11List);
				
				
				
				int reportsRow12ListCount = reportsRow12List.size();
				ArrayList<String> reportsRow12ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow12ListCount;i++)
				{
					String data = reportsRow12List.get(i).getText();
					reportsRow12ListArray.add(data);
				}
				String actRow12List = reportsRow12ListArray.toString();
				String expRow12List = excelReader.getCellData(xlSheetName, 229, 6);
				excelReader.setCellData(xlfile, xlSheetName, 229, 7, actRow12List);
				
				
				/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
				report_NextBtn.click();*/
				
				int reportsRow13ListCount = reportsRow13List.size();
				ArrayList<String> reportsRow13ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow13ListCount;i++)
				{
					String data = reportsRow13List.get(i).getText();
					System.err.println(data);
					reportsRow13ListArray.add(data);
				}
				String actRow13List = reportsRow13ListArray.toString();
				String expRow13List = excelReader.getCellData(xlSheetName, 230, 6);
				excelReader.setCellData(xlfile, xlSheetName, 230, 7, actRow13List);
				
				
				System.out.println("********************checkStockMovementReportsWithStockMovementByWarehouseAndIncludeServiceTypeItemOptions**********************");

				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				System.out.println(actRow4List);
				System.out.println(expRow4List);
				
				System.out.println(actRow5List);
				System.out.println(expRow5List);
				
				System.out.println(actRow6List);
				System.out.println(expRow6List);
				
				System.out.println(actRow7List);
				System.out.println(expRow7List);
				
				System.out.println(actRow8List);
				System.out.println(expRow8List);
				
				System.out.println(actRow9List);
				System.out.println(expRow9List);
				
				System.out.println(actRow10List);
				System.out.println(expRow10List);
				
				System.out.println(actRow11List);
				System.out.println(expRow11List);
				
				System.out.println(actRow12List);
				System.out.println(expRow12List);
				
				System.out.println(actRow13List);
				System.out.println(expRow13List);
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
					&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
					&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List) && actRow12List.equalsIgnoreCase(expRow12List) 
					&& actRow13List.equalsIgnoreCase(expRow13List))
				{
					 excelReader.setCellData(xlfile, xlSheetName, 218, 8, resPass);
					return true;
				}
				else
				{
					 excelReader.setCellData(xlfile, xlSheetName, 218, 8, resFail);

					return false;
				}
			}
			
			
			
			
			@FindBy (xpath="//a[contains(text(),' Warehouse')]//i")
			private static WebElement report_FilterWarehouseExpandBtn;
			
			@FindBy(xpath="//input[@id='5058']")
			private static WebElement report_FilterWarehouseNameChkbox;

			
			
			@FindBy (xpath="//*[@id='FilterFields_582_0']/ul[5]/li[1]/div/label")
			private static WebElement sa_report_FilterWarehouseNameChkbox;
			
			@FindBy (xpath="//input[@id='FOption_575_0_DefaultFilter_0']")
			private static WebElement sMovement_DefaultFilterTxt;
			
			@FindBy (xpath="(//*[@id='FilterFields_575_0']/ul[5]/li[1]/div/label)[1]")
			private static WebElement smreport_FilterWarehouseNameChkbox;

			
			
			public boolean checkStockMovementReportWithFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				  excelReader=new ExcelReader(POJOUtility.getExcelPath());
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
				report_FilterExpandBtn.click();
				Thread.sleep(1000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
				report_FilterBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
				report_FilterCustomizeBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterWarehouseExpandBtn));
				report_FilterWarehouseExpandBtn.click();
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(smreport_FilterWarehouseNameChkbox));
				
				if (report_FilterWarehouseNameChkbox.isSelected()==false) 
				{
					smreport_FilterWarehouseNameChkbox.click();
				}
				
				//smreport_FilterWarehouseNameChkbox.click();
				Thread.sleep(1000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
				reports_Filter_Customize_OkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sMovement_DefaultFilterTxt));
				sMovement_DefaultFilterTxt.click();
				sMovement_DefaultFilterTxt.sendKeys( excelReader.getCellData(xlSheetName, 233, 5));
				
				Thread.sleep(2000);
				
				sMovement_DefaultFilterTxt.sendKeys(Keys.TAB);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
				report_FilterOkBtn.click();
				
				Thread.sleep(2000);
			
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));

				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List =  excelReader.getCellData(xlSheetName, 233, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 233, 7, actRow1List);
				

				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List =  excelReader.getCellData(xlSheetName, 234, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 234, 7, actRow2List);
				
				
				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List =  excelReader.getCellData(xlSheetName, 235, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 235, 7, actRow3List);
				
				
				
				int reportsRow4ListCount = reportsRow4List.size();
				ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow4ListCount;i++)
				{
					String data = reportsRow4List.get(i).getText();
					reportsRow4ListArray.add(data);
				}
				String actRow4List = reportsRow4ListArray.toString();
				String expRow4List =  excelReader.getCellData(xlSheetName, 236, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 236, 7, actRow4List);
				
				
				
				int reportsRow5ListCount = reportsRow5List.size();
				ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow5ListCount;i++)
				{
					String data = reportsRow5List.get(i).getText();
					reportsRow5ListArray.add(data);
				}
				String actRow5List = reportsRow5ListArray.toString();
				String expRow5List =  excelReader.getCellData(xlSheetName, 237, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 237, 7, actRow5List);
				
				
				
				int reportsRow6ListCount = reportsRow6List.size();
				ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow6ListCount;i++)
				{
					String data = reportsRow6List.get(i).getText();
					reportsRow6ListArray.add(data);
				}
				String actRow6List = reportsRow6ListArray.toString();
				String expRow6List =  excelReader.getCellData(xlSheetName, 238, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 238, 7, actRow6List);
				
				System.out.println("****************************checkStockMovementReportWithFilterOption******************************");
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				System.out.println(actRow4List);
				System.out.println(expRow4List);
				
				System.out.println(actRow5List);
				System.out.println(expRow5List);
				
				System.out.println(actRow6List);
				System.out.println(expRow6List);
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
				report_FilterExpandBtn.click();
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
				report_FilterBtn.click();
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
				report_FilterCustomizeBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(smreport_FilterWarehouseNameChkbox));
				
				if (smreport_FilterWarehouseNameChkbox.isSelected()==true) 
				{
					smreport_FilterWarehouseNameChkbox.click();
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
				reports_Filter_Customize_OkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
				report_Filter_RefreshConditionBtn.click();

				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
				report_FilterOkBtn.click();
				
				Thread.sleep(2000);
				
				/*int actRowCount = reportsRowsCount.size();
				int expRowCount = 12; */
				
				String actRowCount=String.valueOf(reportsRowsCount.size());
				String expRowCount=excelReader.getCellData(xlSheetName, 239, 6);
				excelReader.setCellData(xlfile, xlSheetName, 239, 7, actRowCount);
				
				System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
				
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
					&& actRowCount.equals(expRowCount))
				{
					 excelReader.setCellData(xlfile, xlSheetName, 233, 8, resPass);
					return true;
				} 
				else 
				{
					 excelReader.setCellData(xlfile, xlSheetName, 233, 8, resFail);

					return false;
				}
			}
			
			
			
		/*	
			@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[1]/div[1]/div[2]/table/tbody/tr/td")
			private static List<WebElement> backTrackItemDetailsMonthsTableList;
			
			@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[1]/div[1]/div[2]/table")
			private static WebElement backTrackItemDetailsMonthsTable;
			
			@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[1]/div[1]/div[2]/table/tbody/tr[1]/td[1]")
			private static WebElement backTrackItemDetailsMonths1stRow1stCol;
			
			@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[1]/div[1]/div[2]/table/tbody/tr[2]/td[1]")
			private static WebElement backTrackItemDetailsMonths2ndRow1stCol;
			
			
			@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[2]/div[1]/div[2]/table/tbody/tr/td")
			private static List<WebElement> backTrackItemDetailsDaysTableList;
			
			@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[2]/div[1]/div[2]/table")
			private static WebElement backTrackItemDetailsDaysTable;
		
			@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[1]")
			private static WebElement backTrackItemDetailsDays1stRow1stCol;
			
			
			
			@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[3]/div[1]/div[2]/table/tbody/tr/td")
			private static List<WebElement> backTrackItemDetailsVouchersTableList;
			
			@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[3]/div[1]/div[2]/table")
			private static WebElement backTrackItemDetailsVouchersTable;
			
			@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[3]/div[1]/div[2]/table/tbody/tr[1]/td[1]")
			private static WebElement backTrackItemDetailsVouchers1stRow1stCol;
			*/
			
			
			
			
			public boolean checkStockMovementReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				 excelReader=new ExcelReader(POJOUtility.getExcelPath());
				 /* getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
					inventoryMenu.click();
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
					inventoryReportsMenu.click();
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockMovement));
					stockMovement.click();
					
					Thread.sleep(2000);
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
					Select s=new Select(sl_DateOptionDropdown);
					s.selectByIndex(1);
					Thread.sleep(2000);
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_HeaderSelectChkBox));
					sl_HeaderSelectChkBox.click();
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockMovementDropdown));
					Select sm=new Select(stockMovementDropdown);
					sm.selectByIndex(1);
					Thread.sleep(2000);
					
					//getAction().moveToElement(sl_IncludeServiceTypeItemChkBox).build().perform();
					//Thread.sleep(1000);
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_IncludeServiceTypeItemChkBox));
					sl_IncludeServiceTypeItemChkBox.click();
					
					Thread.sleep(1000);
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
					sl_OkBtn.click();
					*/
					Thread.sleep(3000);
					
				 int reportItemGridListCount=reportTableColumn1RowsList.size();
			    for (int i = 0; i < reportItemGridListCount; i++) 
			    {
		             String data=reportTableColumn1RowsList.get(i).getText();
		             
		             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 242, 5)))
		             {
	                     int count =i+5;
	                     for (int j = 0; j < count; j++) 
	         		     {
	                    	 if(j!=3)
	                    	 {
	                    	 String data1=reportTableColumn1RowsList.get(j).getText();
				             
	                    	 if (data1.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 243, 5)))
	                         {
	                    		 System.err.println("Item Name : "+data1);
	                    		 reportTableColumn1RowsList.get(j).click();
	                    		 Thread.sleep(2000);
	                    		 break;
	                         }
	                    	 }
	         		     }
		             }
			      }

				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
				sl_BackTrackBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
				
				int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
				ArrayList<String> actMonthlyList = new ArrayList<String>();
				for(int i=0; i<backTrackMonthlyColumnsCount; i++)
				{
					String data = backTrackItemDetailsMonthsTableList.get(i).getText();
					
					actMonthlyList.add(data);
				}
				
				String actMonth=actMonthlyList.toString();
				excelReader.setCellData(xlfile, xlSheetName, 243, 7,actMonth);
				
				Calendar cal=Calendar.getInstance();
				DateFormat df = new SimpleDateFormat("MMMM yyyy");
				String Row2Month = df.format(cal.getTime());
				
				//ArrayList<String> expMonthlyList = new ArrayList<String>();
			/*	expMonthlyList.add( excelReader.getCellData(xlSheetName, 243, 6));
				expMonthlyList.add("");
				expMonthlyList.add("36.00");
				expMonthlyList.add("");
				expMonthlyList.add("36.00");
				expMonthlyList.add(Row2Month);
				expMonthlyList.add("36.00");
				expMonthlyList.add("57.00");
				expMonthlyList.add("24.00");
				expMonthlyList.add("69.00");
				expMonthlyList.add("");
				expMonthlyList.add("");
				expMonthlyList.add("93.00");
				expMonthlyList.add("24.00");
				expMonthlyList.add("");*/
				
				//String expMonth=expMonthlyList.toString();

				String expMonthlyList=excelReader.getCellData(xlSheetName, 243, 6);
				System.out.println(actMonth);
				System.out.println(expMonthlyList);
				
				
				for(int i=0; i<backTrackMonthlyColumnsCount; i++)
				{
					String data = backTrackItemDetailsMonthsTableList.get(i).getText();
					
					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 244, 5))) 
					{
						getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
						break;
					}
				}
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
				
				int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
				ArrayList<String> actDayList = new ArrayList<String>();
				for(int i=0; i<backTrackDayColumnsCount; i++)
				{
					String data = backTrackItemDetailsDaysTableList.get(i).getText();
					actDayList.add(data);
				}
				
				String actDay=actDayList.toString();
				 excelReader.setCellData(xlfile, xlSheetName, 244, 7, actDay);
				//ArrayList<String> expDayList = new ArrayList<String>();
				/*expDayList.add("1");
				expDayList.add("");
				expDayList.add("36.00");
				expDayList.add("");
				expDayList.add("36.00");
				expDayList.add("");
				expDayList.add("");
				expDayList.add("36.00");
				expDayList.add("");
				expDayList.add("");
				*/
				//String expDay=expDayList.toString();
				String expDayList=excelReader.getCellData(xlSheetName, 244, 6);

				System.out.println(actDay);
				System.out.println(expDayList);
				
				for(int i=0; i<backTrackDayColumnsCount; i++)
				{
					String data = backTrackItemDetailsDaysTableList.get(i).getText();
					
					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 245, 5))) 
					{
						getAction().doubleClick().build().perform();
						break;
					}
				}
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
				
				int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
				ArrayList<String> actVouchersList = new ArrayList<String>();
				for(int i=0; i<backTrackVouchersColumnsCount; i++)
				{
					String data = backTrackItemDetailsVouchersTableList.get(i).getText();
					actVouchersList.add(data);
				}
				
				String actVoucher=actVouchersList.toString();

                excelReader.setCellData(xlfile, xlSheetName, 245, 7, actVoucher);
                
				//ArrayList<String> expVouchersList = new ArrayList<String>();
				/*expVouchersList.add("OpeStk:3");
				expVouchersList.add("");
				expVouchersList.add("36.00");
				expVouchersList.add("");
				expVouchersList.add("36.00");
				expVouchersList.add("");
				expVouchersList.add("");
				expVouchersList.add("36.00");
				expVouchersList.add("");
				expVouchersList.add("");
				*/
				//String expVoucher=expVouchersList.toString();
				String expVouchersList=excelReader.getCellData(xlSheetName, 245, 6);
				
				System.out.println(actVoucher);
				System.out.println(expVouchersList);
				
				for(int i=0; i<backTrackVouchersColumnsCount; i++)
				{
					String data = backTrackItemDetailsVouchersTableList.get(i).getText();

					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 246, 5))) 
					{
						getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
						break;
					}
				}
				
				Thread.sleep(2000);
				
				ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
				int actOpenWindowsCount = getDriver().getWindowHandles().size();
				int expOpenWindowsCount = 2;

			 	getDriver().switchTo().window(openTabs.get(1));
				
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
				String actDocNo=documentNumberTxt.getAttribute("value");
				String expDocNo=excelReader.getCellData(xlSheetName, 246, 6);
				excelReader.setCellData(xlfile, xlSheetName, 246, 7, actDocNo);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				String actWarehouse=warehouseTxt.getAttribute("value");
				String expWarehouse=excelReader.getCellData(xlSheetName, 247, 6);
				excelReader.setCellData(xlfile, xlSheetName, 247, 7, actWarehouse);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				String actR1Item=select1stRow_1stColumn.getText();
				String expR1Item=excelReader.getCellData(xlSheetName, 248, 6);
				excelReader.setCellData(xlfile, xlSheetName, 248, 7, actR1Item);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
				String actR1Units=select1stRow_2ndColumn.getText();
				String expR1Units=excelReader.getCellData(xlSheetName, 249, 6);
				excelReader.setCellData(xlfile, xlSheetName, 249, 7, actR1Units);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
				String actR1Quantity=select1stRow_3rdColumn.getText();
				String expR1Quantity=excelReader.getCellData(xlSheetName, 250, 6);
				excelReader.setCellData(xlfile, xlSheetName, 250, 7, actR1Quantity);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
				String actR1Rate=select1stRow_4thColumn.getText();
				String expR1Rate=excelReader.getCellData(xlSheetName, 251, 6);
				excelReader.setCellData(xlfile, xlSheetName, 251, 7, actR1Rate);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
				String actR1Gross=select1stRow_5thColumn.getText();
				String expR1Gross=excelReader.getCellData(xlSheetName, 252, 6);
				excelReader.setCellData(xlfile, xlSheetName, 252, 7, actR1Gross);

				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
				String actR1Bin=select1stRow_7thColumn.getText();
				String expR1Bin=excelReader.getCellData(xlSheetName, 253, 6);
				excelReader.setCellData(xlfile, xlSheetName, 253, 7, actR1Bin);

				
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
				String actR2Item=select2ndRow_1stColumn.getText();
				String expR2Item=excelReader.getCellData(xlSheetName, 254, 6);
				excelReader.setCellData(xlfile, xlSheetName, 254, 7, actR2Item);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
				String actR2Units=select2ndRow_2ndColumn.getText();
				String expR2Units=excelReader.getCellData(xlSheetName, 255, 6);
				excelReader.setCellData(xlfile, xlSheetName, 255, 7, actR2Units);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
				String actR2Quantity=select2ndRow_3rdColumn.getText();
				String expR2Quantity=excelReader.getCellData(xlSheetName, 256, 6);
				excelReader.setCellData(xlfile, xlSheetName, 256, 7, actR2Quantity);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
				String actR2Rate=select2ndRow_4thColumn.getText();
				String expR2Rate=excelReader.getCellData(xlSheetName, 257, 6);
				excelReader.setCellData(xlfile, xlSheetName, 257, 7, actR2Rate);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
				String actR2Gross=select2ndRow_5thColumn.getText();
				String expR2Gross=excelReader.getCellData(xlSheetName, 258, 6);
				excelReader.setCellData(xlfile, xlSheetName, 258, 7, actR2Gross);

				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_7thColumn));
				String actR2Bin=select2ndRow_7thColumn.getText();
				String expR2Bin=excelReader.getCellData(xlSheetName, 259, 6);
				excelReader.setCellData(xlfile, xlSheetName, 259, 7, actR2Bin);

				
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
				String actR3Item=select3rdRow_1stColumn.getText();
				String expR3Item=excelReader.getCellData(xlSheetName, 260, 6);
				excelReader.setCellData(xlfile, xlSheetName, 260, 7, actR3Item);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_2ndColumn));
				String actR3Units=select3rdRow_2ndColumn.getText();
				String expR3Units=excelReader.getCellData(xlSheetName, 261, 6);
				excelReader.setCellData(xlfile, xlSheetName, 261, 7, actR3Units);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_3rdColumn));
				String actR3Quantity=select3rdRow_3rdColumn.getText();
				String expR3Quantity=excelReader.getCellData(xlSheetName, 262, 6);
				excelReader.setCellData(xlfile, xlSheetName, 262, 7, actR3Quantity);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_4thColumn));
				String actR3Rate=select3rdRow_4thColumn.getText();
				String expR3Rate=excelReader.getCellData(xlSheetName, 263, 6);
				excelReader.setCellData(xlfile, xlSheetName, 263, 7, actR3Rate);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_5thColumn));
				String actR3Gross=select3rdRow_5thColumn.getText();
				String expR3Gross=excelReader.getCellData(xlSheetName, 264, 6);
				excelReader.setCellData(xlfile, xlSheetName, 264, 7, actR3Gross);

				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_7thColumn));
				String actR3Bin=select3rdRow_7thColumn.getText();
				String expR3Bin=excelReader.getCellData(xlSheetName, 265, 6);
				excelReader.setCellData(xlfile, xlSheetName, 265, 7, actR3Bin);

				

				System.out.println("**********************checkStockMovementReportBackTrackOption*********************");
				
				System.out.println("*********** Voucher No : "+actDocNo        +"  value expected  "+expDocNo);
				System.out.println("*********** Warehouse  : "+actWarehouse    +"  value expected  "+expWarehouse);
				System.out.println("*********************************Row1******************************************");
				System.out.println("*********** Item       : "+actR1Item       +"  value expected  "+expR1Item);
				System.out.println("*********** Units      : "+actR1Units      +"  value expected  "+expR1Units);
				System.out.println("*********** Quantity   : "+actR1Quantity   +"  value expected  "+expR1Quantity);
				System.out.println("*********** Rate       : "+actR1Rate       +"  value expected  "+expR1Rate);
				System.out.println("*********** Gross      : "+actR1Gross      +"  value expected  "+expR1Gross);
				System.out.println("*********** Bin        : "+actR1Bin        +"  value expected  "+expR1Bin);
				System.out.println("*********************************Row2******************************************");
				System.out.println("*********** Item       : "+actR2Item       +"  value expected  "+expR2Item);
				System.out.println("*********** Units      : "+actR2Units      +"  value expected  "+expR2Units);
				System.out.println("*********** Quantity   : "+actR2Quantity   +"  value expected  "+expR2Quantity);
				System.out.println("*********** Rate       : "+actR2Rate       +"  value expected  "+expR2Rate);
				System.out.println("*********** Gross      : "+actR2Gross      +"  value expected  "+expR2Gross);
				System.out.println("*********** Bin        : "+actR2Bin        +"  value expected  "+expR2Bin);
				System.out.println("*********************************Row3******************************************");
				System.out.println("*********** Item       : "+actR3Item       +"  value expected  "+expR3Item);
				System.out.println("*********** Units      : "+actR3Units      +"  value expected  "+expR3Units);
				System.out.println("*********** Quantity   : "+actR3Quantity   +"  value expected  "+expR3Quantity);
				System.out.println("*********** Rate       : "+actR3Rate       +"  value expected  "+expR3Rate);
				System.out.println("*********** Gross      : "+actR3Gross      +"  value expected  "+expR3Gross);
				System.out.println("*********** Bin        : "+actR3Bin        +"  value expected  "+expR3Bin);
				
				getDriver().switchTo().window(openTabs.get(1)).close();
				
				Thread.sleep(1000);
				
				getDriver().switchTo().window(openTabs.get(0));
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
				BackTrackItemDetailsBackBtn.click();
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
				BackTrackMonthDetailsBackBtn.click();
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
				BackTrackDateDetailsBackBtn.click();
				
				Thread.sleep(1000);
				
				if (actMonth.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucher.equals(expVouchersList) && actOpenWindowsCount==expOpenWindowsCount
						&& actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) 
						&& actR1Item.equalsIgnoreCase(expR1Item) && actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) 
						&& actR1Rate.equalsIgnoreCase(expR1Rate) && actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Bin.equalsIgnoreCase(expR1Bin) 
						
						&& actR2Item.equalsIgnoreCase(expR2Item) && actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) 
						&& actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Bin.equalsIgnoreCase(expR2Bin) 
						
						&& actR3Item.equalsIgnoreCase(expR3Item) && actR3Units.equalsIgnoreCase(expR3Units) && actR3Quantity.equalsIgnoreCase(expR3Quantity) 
						&& actR3Rate.equalsIgnoreCase(expR3Rate) && actR3Gross.equalsIgnoreCase(expR3Gross) && actR3Bin.equalsIgnoreCase(expR3Bin))
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					 excelReader.setCellData(xlfile, xlSheetName, 242, 8, resPass);
					return true;
				} 
				else 
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					 excelReader.setCellData(xlfile, xlSheetName, 242, 8, resFail);

					return false;
				}
			}
			
			
			
			@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[7]")
			private static WebElement multiLevelStockMovement;

			public boolean checkMultiLevelStockMovementReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				excelReader=new ExcelReader(POJOUtility.getExcelPath());
				
				LogoutandLoginwithSU();
				
				getDriver().navigate().refresh();
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(multiLevelStockMovement));
				multiLevelStockMovement.click();
				
				Thread.sleep(2000);
				
				int rowcount=stockLedgerHometableRowCount.size();
				
				System.out.println(rowcount);
				
				for (int i = 0; i < rowcount; i++) 
				{
					String actName = stockLedgerHometableItemNamesList.get(i).getText();
					
					System.out.println(actName);
					
					if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 268, 5)) || actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 269, 5)) 
							|| actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 270, 5)) 
							|| actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 271, 5)))
					{
						stockLedgerHometableItemChkboxList.get(i).click();
					}
				}
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
				
				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List =  excelReader.getCellData(xlSheetName, 271, 6);
				excelReader.setCellData(xlfile, xlSheetName, 271, 7, actRow1List);
				
				

				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List = excelReader.getCellData(xlSheetName, 272, 6);
				excelReader.setCellData(xlfile, xlSheetName, 272, 7, actRow2List);
				
				
				
				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List = excelReader.getCellData(xlSheetName, 273, 6);
				excelReader.setCellData(xlfile, xlSheetName, 273, 7, actRow3List);
				
				
				
				
				int reportsRow4ListCount = reportsRow4List.size();
				ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow4ListCount;i++)
				{
					String data = reportsRow4List.get(i).getText();
					reportsRow4ListArray.add(data);
				}
				String actRow4List = reportsRow4ListArray.toString();
				String expRow4List = excelReader.getCellData(xlSheetName, 274, 6);
				excelReader.setCellData(xlfile, xlSheetName, 274, 7, actRow4List);
				
				
				
				
				int reportsRow5ListCount = reportsRow5List.size();
				ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow5ListCount;i++)
				{
					String data = reportsRow5List.get(i).getText();
					reportsRow5ListArray.add(data);
				}
				String actRow5List = reportsRow5ListArray.toString();
				String expRow5List = excelReader.getCellData(xlSheetName, 275, 6);
				excelReader.setCellData(xlfile, xlSheetName, 275, 7, actRow5List);
				
				
				
				System.out.println("****************************checkMultiLevelStockMovementReports******************************");
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				System.out.println(actRow4List);
				System.out.println(expRow4List);
				
				System.out.println(actRow5List);
				System.out.println(expRow5List);
				
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
				{
					excelReader.setCellData(xlfile, xlSheetName, 268, 8, resPass);
					return true;
				}
				else
				{
					excelReader.setCellData(xlfile, xlSheetName, 268, 8, resFail);

					return false;
				}
			}
			
			
			
			
			
			
			@FindBy (xpath="//input[@id='FOption_576_0_DefaultFilter_0']")
			private static WebElement mlsm_DefaultFilterTxt;
			
			@FindBy (xpath="//*[@id='FilterFields_576_0']/ul[4]/li[1]/div/label/span")
			private static WebElement mlsm_report_FilterItemNameChkbox;
			
			
			public boolean checkMultiLevelStockMovementReportFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				 excelReader=new ExcelReader(POJOUtility.getExcelPath());
				 
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
					report_FilterExpandBtn.click();
					Thread.sleep(1000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
				report_FilterBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
				report_FilterCustomizeBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterItemExpansion));
				report_FilterItemExpansion.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mlsm_report_FilterItemNameChkbox));
				
				if (report_FilterItemNameChkbox.isSelected()==false) 
				{
					mlsm_report_FilterItemNameChkbox.click();
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
				reports_Filter_Customize_OkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mlsm_DefaultFilterTxt));
				mlsm_DefaultFilterTxt.click();
				mlsm_DefaultFilterTxt.sendKeys( excelReader.getCellData(xlSheetName,279 , 5));

				
				Thread.sleep(2000);
				
				mlsm_DefaultFilterTxt.sendKeys(Keys.TAB);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
				report_FilterOkBtn.click();
				
				Thread.sleep(3000);
				
				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List =  excelReader.getCellData(xlSheetName, 279, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 279, 7, actRow1List);
				
				
	
				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List = excelReader.getCellData(xlSheetName, 280, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 280, 7, actRow2List);
			

				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List = excelReader.getCellData(xlSheetName, 281, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 281, 7, actRow3List);
				
				
				System.out.println("***************************checkMultiLevelStockMovementReportFilterOption*********************************");
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
					
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
				report_FilterExpandBtn.click();
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
				report_FilterBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
				report_FilterCustomizeBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(mlsm_report_FilterItemNameChkbox));
				
				if (mlsm_report_FilterItemNameChkbox.isSelected()==true) 
				{
					mlsm_report_FilterItemNameChkbox.click();
				}

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
				reports_Filter_Customize_OkBtn.click();
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
				report_Filter_RefreshConditionBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
				report_FilterOkBtn.click();
				
				int actRowCount = reportsRowsCount.size();
				int expRowCount = 1; 
				
				System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
				
				Thread.sleep(2000);
				
				//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reportWithLevel_CloseBtn));
				//reportWithLevel_CloseBtn.click();
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
					&& actRowCount==expRowCount)
				{
					excelReader.setCellData(xlfile, xlSheetName, 279, 8, resPass);
					return true;
				}
				else
				{
					excelReader.setCellData(xlfile, xlSheetName, 279, 8, resFail);

					return false;
				}
			}

			
			
			
			
			
			
			
			public boolean checkMultiLevelStockMovementReportWithIncludeServiceTypeItemOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				//Thread.sleep(1000);
				 excelReader=new ExcelReader(POJOUtility.getExcelPath());
				 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
					inventoryMenu.click();
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
					inventoryReportsMenu.click();
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(multiLevelStockMovement));
					multiLevelStockMovement.click();
					
					Thread.sleep(2000);
					
					int rowcount=stockLedgerHometableRowCount.size();
					
					System.out.println(rowcount);
					
					for (int i = 0; i < rowcount; i++) 
					{
						String actName = stockLedgerHometableItemNamesList.get(i).getText();
						
						System.out.println(actName);
						
						if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 285, 5)) || actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 286, 5)) 
								|| actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 287, 5)) 
								|| actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 288, 5)))
						{
							stockLedgerHometableItemChkboxList.get(i).click();
						}
					}
					
					
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
					Select s=new Select(sl_DateOptionDropdown);
					
					s.selectByIndex(1);
					Thread.sleep(2000);
				getAction().moveToElement(sl_IncludeServiceTypeItemChkBox).build().perform();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_IncludeServiceTypeItemChkBox));
				sl_IncludeServiceTypeItemChkBox.click();
				
				Thread.sleep(1000);
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
				
				int reportsRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsRow1ListArray.add(data);
				}
				String actRow1List = reportsRow1ListArray.toString();
				String expRow1List =  excelReader.getCellData(xlSheetName, 288, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 288, 7, actRow1List);
				
				

				int reportsRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsRow2ListArray.add(data);
				}
				String actRow2List = reportsRow2ListArray.toString();
				String expRow2List =  excelReader.getCellData(xlSheetName, 289, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 289, 7, actRow2List);

				
				
				
				int reportsRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsRow3ListArray.add(data);
				}
				String actRow3List = reportsRow3ListArray.toString();
				String expRow3List =  excelReader.getCellData(xlSheetName, 290, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 290, 7, actRow3List);

				
				
				
				int reportsRow4ListCount = reportsRow4List.size();
				ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow4ListCount;i++)
				{
					String data = reportsRow4List.get(i).getText();
					reportsRow4ListArray.add(data);
				}
				String actRow4List = reportsRow4ListArray.toString();
				String expRow4List =  excelReader.getCellData(xlSheetName, 291, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 291, 7, actRow4List);

				
				
				
				int reportsRow5ListCount = reportsRow5List.size();
				ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
				for(int i=1;i<reportsRow5ListCount;i++)
				{
					String data = reportsRow5List.get(i).getText();
					reportsRow5ListArray.add(data);
				}
				String actRow5List = reportsRow5ListArray.toString();
				String expRow5List =  excelReader.getCellData(xlSheetName, 292, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 292, 7, actRow5List);

				
				
				System.out.println("********************checkMultiLevelStockMovementReportWithIncludeServiceTypeItemOption**********************");

				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				System.out.println(actRow4List);
				System.out.println(expRow4List);
				
				System.out.println(actRow5List);
				System.out.println(expRow5List);
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
					&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
				{
					excelReader.setCellData(xlfile, xlSheetName, 285, 8, resPass);
					return true;
				}
				else
				{
					excelReader.setCellData(xlfile, xlSheetName, 285, 8, resFail);
					return false;
				}
			}
			
			
			
			
			
			
			
			
			
			public boolean checkMultiLevelStockMovementReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				Thread.sleep(2000);
				int reportItemGridListCount=reportTableColumn1RowsList.size();
			    for (int i = 0; i < reportItemGridListCount; i++) 
			    {
		             String data=reportTableColumn1RowsList.get(i).getText();
		             
		             if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 299, 5)))
		             {
	            		 reportTableColumn1RowsList.get(i).click();
	            		 Thread.sleep(1000);
	            		 break;
	                 }
	 		     }
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
				sl_BackTrackBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
				
				int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
				ArrayList<String> actMonthlyList = new ArrayList<String>();
				for(int i=0; i<backTrackMonthlyColumnsCount; i++)
				{
					String data = backTrackItemDetailsMonthsTableList.get(i).getText();
					actMonthlyList.add(data);
				}
				
				Calendar cal=Calendar.getInstance();
				DateFormat df = new SimpleDateFormat("MMMM yyyy");
				String Row2Month = df.format(cal.getTime());
				
				String actMonthly=actMonthlyList.toString();
				
			/*	ArrayList<String> expMonthlyList = new ArrayList<String>();
				expMonthlyList.add("January 2020");
				expMonthlyList.add("");
				expMonthlyList.add("24.00");
				expMonthlyList.add("");
				expMonthlyList.add("24.00");
				expMonthlyList.add("November 2021");
				expMonthlyList.add("24.00");
				expMonthlyList.add("42.00");
				expMonthlyList.add("23.00");
				expMonthlyList.add("43.00");
				expMonthlyList.add("");
				expMonthlyList.add("");
				expMonthlyList.add("66.00");
				expMonthlyList.add("23.00");
				expMonthlyList.add("");*/
				
				String expMonthlyList=excelReader.getCellData(xlSheetName, 299, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 299, 7, actMonthly);

				
				System.out.println(actMonthly);
				System.out.println(expMonthlyList);
				
				for(int i=0; i<backTrackMonthlyColumnsCount; i++)
				{
					String data = backTrackItemDetailsMonthsTableList.get(i).getText();
					
					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 300, 5))) 
					{
						getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
						break;
					}
				}
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
				
				int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
				ArrayList<String> actDayList = new ArrayList<String>();
				for(int i=0; i<backTrackDayColumnsCount; i++)
				{
					String data = backTrackItemDetailsDaysTableList.get(i).getText();
					actDayList.add(data);
				}
				
				String actDay=actDayList.toString();
				
				/*ArrayList<String> expDayList = new ArrayList<String>();
				expDayList.add("1");
				expDayList.add("");
				expDayList.add("24.00");
				expDayList.add("");
				expDayList.add("24.00");
				expDayList.add("");
				expDayList.add("");
				expDayList.add("24.00");
				expDayList.add("");
				expDayList.add("");*/
				String expDayList= excelReader.getCellData(xlSheetName, 300, 6);
				excelReader.setCellData(xlfile, xlSheetName, 300, 7, actDay);

				
				System.out.println(actDay);
				System.out.println(expDayList);
				
				for(int i=0; i<backTrackDayColumnsCount; i++)
				{
					String data = backTrackItemDetailsDaysTableList.get(i).getText();

					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 301, 5))) 
					{
						getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
					}
				}
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
				
				int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
				ArrayList<String> actVouchersList = new ArrayList<String>();
				for(int i=0; i<backTrackVouchersColumnsCount; i++)
				{
					String data = backTrackItemDetailsVouchersTableList.get(i).getText();
					actVouchersList.add(data);
				}
				
				String actVoucher=actVouchersList.toString();
				
				/*ArrayList<String> expVouchersList = new ArrayList<String>();
				expVouchersList.add("OpeStk:4");
				expVouchersList.add("");
				expVouchersList.add("24.00");
				expVouchersList.add("");
				expVouchersList.add("24.00");
				expVouchersList.add("");
				expVouchersList.add("");
				expVouchersList.add("24.00");
				expVouchersList.add("");
				expVouchersList.add("");*/
				
				String expVouchersList= excelReader.getCellData(xlSheetName, 301, 6);
				excelReader.setCellData(xlfile, xlSheetName, 301, 7, actVoucher);
				
				System.out.println(actVouchersList);
				System.out.println(expVouchersList);
				
				for(int i=0; i<backTrackVouchersColumnsCount; i++)
				{
					String data = backTrackItemDetailsVouchersTableList.get(i).getText();

					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 302, 5))) 
					{
						getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
						break;
					}
				}
				
				Thread.sleep(2000);
				
				ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
				int actCount = getDriver().getWindowHandles().size();
				String actOpenWindowsCount=String.valueOf(actCount);
				
				String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 302, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 302, 7, actOpenWindowsCount);
				

			 	getDriver().switchTo().window(openTabs.get(1));
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
				String actDocNo=documentNumberTxt.getAttribute("value");
				String expDocNo=excelReader.getCellData(xlSheetName, 303, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 303, 7, actDocNo);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
				String actWarehouse=warehouseTxt.getAttribute("value");
				String expWarehouse=excelReader.getCellData(xlSheetName, 304, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 304, 7, actWarehouse);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
				String actItem=select1stRow_1stColumn.getText();
				String expItem=excelReader.getCellData(xlSheetName, 305, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 305, 7, actItem);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
				String actUnits=select1stRow_2ndColumn.getText();
				String expUnits=excelReader.getCellData(xlSheetName, 306, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 306, 7, actUnits);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
				String actQuantity=select1stRow_3rdColumn.getText();
				String expQuantity=excelReader.getCellData(xlSheetName, 307, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 307, 7, actQuantity);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
				String actRate=select1stRow_4thColumn.getText();
				String expRate=excelReader.getCellData(xlSheetName, 308, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 308, 7, actRate);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
				String actGross=select1stRow_5thColumn.getText();
				String expGross=excelReader.getCellData(xlSheetName, 309, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 309, 7, actGross);

				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
				String actRMA=select1stRow_9thColumn.getText();
				String expRMA=excelReader.getCellData(xlSheetName, 310, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 310, 7, actRMA);

				

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
				String actItem2=select2ndRow_1stColumn.getText();
				String expItem2=excelReader.getCellData(xlSheetName, 311, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 311, 7, actItem2);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
				String actUnits2=select2ndRow_2ndColumn.getText();
				String expUnits2=excelReader.getCellData(xlSheetName, 312, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 312, 7, actUnits2);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
				String actQuantity2=select2ndRow_3rdColumn.getText();
				String expQuantity2=excelReader.getCellData(xlSheetName,313, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 313, 7, actQuantity2);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
				String actRate2=select2ndRow_4thColumn.getText();
				String expRate2=excelReader.getCellData(xlSheetName, 314, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 314, 7, actRate2);

				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
				String actGross2=select2ndRow_5thColumn.getText();
				String expGross2=excelReader.getCellData(xlSheetName, 315, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 315, 7, actGross2);

				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
				String actRMA2=select2ndRow_9thColumn.getText();
				String expRMA2=excelReader.getCellData(xlSheetName, 316, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 316, 7, actRMA2);

				
				System.out.println("**********************checkMultiLevelStockMovementReportBackTrackOption*********************");
				
				System.out.println("*********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
				System.out.println("*********** Warehouse  : "+actWarehouse  +"  value expected  "+expWarehouse);
				System.out.println("*********** Item       : "+actItem       +"  value expected  "+expItem);
				System.out.println("*********** Units      : "+actUnits      +"  value expected  "+expUnits);
				System.out.println("*********** Quantity   : "+actQuantity   +"  value expected  "+expQuantity);
				System.out.println("*********** Rate       : "+actRate       +"  value expected  "+expRate);
				System.out.println("*********** Gross      : "+actGross      +"  value expected  "+expGross);
				System.out.println(actRMA);
				System.out.println(expRMA);
				System.out.println("*********** Item2       : "+actItem2       +"  value expected  "+expItem2);
				System.out.println("*********** Units2      : "+actUnits2      +"  value expected  "+expUnits2);
				System.out.println("*********** Quantity2   : "+actQuantity2   +"  value expected  "+expQuantity2);
				System.out.println("*********** Rate2       : "+actRate2       +"  value expected  "+expRate2);
				System.out.println("*********** Gross2      : "+actGross2      +"  value expected  "+expGross2);
				System.out.println(actRMA2);
				System.out.println(expRMA2);
				
				getDriver().switchTo().window(openTabs.get(1)).close();
				
				Thread.sleep(1000);
				
				getDriver().switchTo().window(openTabs.get(0));
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
				BackTrackItemDetailsBackBtn.click();
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
				BackTrackMonthDetailsBackBtn.click();
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
				BackTrackDateDetailsBackBtn.click();
				
				Thread.sleep(1000);
				
				if (actMonthly.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucher.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
					&& actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
					 && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
					 && actGross.equalsIgnoreCase(expGross) && actRMA.equalsIgnoreCase(expRMA)
					 && actItem2.equalsIgnoreCase(expItem2) && actUnits2.equalsIgnoreCase(expUnits2) && actQuantity2.equalsIgnoreCase(expQuantity2) 
					 && actRate2.equalsIgnoreCase(expRate2) && actGross2.equalsIgnoreCase(expGross2) && actRMA2.equalsIgnoreCase(expRMA2)) 
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reportWithLevel_CloseBtn));
					reportWithLevel_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 299, 8, resPass);

					return true;
				}
				else
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reportWithLevel_CloseBtn));
					reportWithLevel_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 299, 8, resFail);

					return false;
				}
			}
			
			
			@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[8]")
			private static WebElement virtualStockAnalysis;
			
			public boolean checkVirtualStockAnalysisReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				excelReader=new ExcelReader(POJOUtility.getExcelPath());
				LogoutandLoginwithSU();
				
				getDriver().navigate().refresh();
				Thread.sleep(2000);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(virtualStockAnalysis));
				virtualStockAnalysis.click();
				
				Thread.sleep(2000);
				
				int rowcount=stockLedgerHometableRowCount.size();
				System.out.println(rowcount);
				
				for (int i = 0; i < rowcount; i++) 
				{
					String actName = stockLedgerHometableItemNamesList.get(i).getText();
					
					System.out.println(actName);
					
					if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 321, 5)) || actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 322, 5)))
					{
						stockLedgerHometableItemChkboxList.get(i).click();
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				

				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
				
				int reportsByWarehouseRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsByWarehouseRow1ListArray.add(data);
				}
				String actRow1List = reportsByWarehouseRow1ListArray.toString();
				String expRow1List =  excelReader.getCellData(xlSheetName, 322, 6);
				  excelReader.setCellData(xlfile, xlSheetName, 322, 7, actRow1List);

				

				int reportsByWarehouseRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
				for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsByWarehouseRow2ListArray.add(data);
				}
				String actRow2List = reportsByWarehouseRow2ListArray.toString();
				String expRow2List =  excelReader.getCellData(xlSheetName, 323, 6);
				  excelReader.setCellData(xlfile, xlSheetName, 323, 7, actRow2List);

				
				
				
				int reportsByWarehouseRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
				for(int i=1;i<reportsByWarehouseRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsByWarehouseRow3ListArray.add(data);
				}
				String actRow3List = reportsByWarehouseRow3ListArray.toString();
				String expRow3List =  excelReader.getCellData(xlSheetName, 324, 6);
				  excelReader.setCellData(xlfile, xlSheetName, 324, 7, actRow3List);

				
				
				System.out.println("*********************************checkVirtualStockAnalysisReport*****************************************");
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					  excelReader.setCellData(xlfile, xlSheetName, 321, 8, resPass);

					return true;
				}
				else
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					  excelReader.setCellData(xlfile, xlSheetName, 321, 8, resFail);

					return false;
				}
			}
			
			
			
			
			
			
			
			
			public boolean checkVirtualStockAnalysisReportWithIncludeServiceTypeItemOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				
				excelReader=new ExcelReader(POJOUtility.getExcelPath());
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
				inventoryMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
				inventoryReportsMenu.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(virtualStockAnalysis));
				virtualStockAnalysis.click();
				
				Thread.sleep(2000);
				
				int rowcount=stockLedgerHometableRowCount.size();
				System.out.println(rowcount);
				
				for (int i = 0; i < rowcount; i++) 
				{
					String actName = stockLedgerHometableItemNamesList.get(i).getText();
					
					System.out.println(actName);
					
					if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 328, 5)) || actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 329, 5)))
					{
						stockLedgerHometableItemChkboxList.get(i).click();
					}
				}
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
				Select s=new Select(sl_DateOptionDropdown);
				s.selectByIndex(1);
				
				
				getAction().moveToElement(sl_IncludeServiceTypeItemChkBox).build().perform();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_IncludeServiceTypeItemChkBox));
				sl_IncludeServiceTypeItemChkBox.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
				sl_OkBtn.click();
				
				Thread.sleep(3000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
				
				int reportsByWarehouseRow1ListCount = reportsRow1List.size();
				ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
				for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
				{
					String data = reportsRow1List.get(i).getText();
					reportsByWarehouseRow1ListArray.add(data);
				}
				String actRow1List = reportsByWarehouseRow1ListArray.toString();
				String expRow1List =  excelReader.getCellData(xlSheetName, 329, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 329, 7, actRow1List);
				

				int reportsByWarehouseRow2ListCount = reportsRow2List.size();
				ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
				for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
				{
					String data = reportsRow2List.get(i).getText();
					reportsByWarehouseRow2ListArray.add(data);
				}
				String actRow2List = reportsByWarehouseRow2ListArray.toString();
				String expRow2List =  excelReader.getCellData(xlSheetName, 330, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 330, 7, actRow2List);
				
				
				
				int reportsByWarehouseRow3ListCount = reportsRow3List.size();
				ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
				for(int i=1;i<reportsByWarehouseRow3ListCount;i++)
				{
					String data = reportsRow3List.get(i).getText();
					reportsByWarehouseRow3ListArray.add(data);
				}
				String actRow3List = reportsByWarehouseRow3ListArray.toString();
				String expRow3List = excelReader.getCellData(xlSheetName, 331, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 331, 7, actRow3List);
				
				
				System.out.println("*********************************checkVirtualStockAnalysisReportWithIncludeServiceTypeItemOption*****************************************");
				
				System.out.println(actRow1List);
				System.out.println(expRow1List);
				
				System.out.println(actRow2List);
				System.out.println(expRow2List);
				
				System.out.println(actRow3List);
				System.out.println(expRow3List);
				
				if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
				{
					  excelReader.setCellData(xlfile, xlSheetName, 328, 8, resPass);
					return true;
				}
				else
				{
					 excelReader.setCellData(xlfile, xlSheetName, 328, 8, resFail);
					return false;
				}
			}
			
			
			
			
			
			
			
			
			public boolean checkVirtualStockAnalysisReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
			{
				
				int reportItemGridListCount=reportTableColumn1RowsList.size();
			    for (int i = 0; i < reportItemGridListCount; i++) 
			    {
		             String data=reportTableColumn1RowsList.get(i).getText();
		             
		             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 336, 5)))
		             {
	            		 reportTableColumn1RowsList.get(i).click();
	            		 Thread.sleep(1000);
	            		 break;
	                 }
	 		     }
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
				sl_BackTrackBtn.click();
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
				
				int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
				ArrayList<String> actMonthlyList = new ArrayList<String>();
				for(int i=0; i<backTrackMonthlyColumnsCount; i++)
				{
					String data = backTrackItemDetailsMonthsTableList.get(i).getText();
					actMonthlyList.add(data);
				}
				
				Calendar cal=Calendar.getInstance();
				DateFormat df = new SimpleDateFormat("MMMM yyyy");
				String Row2Month = df.format(cal.getTime());
				
				String actMonthly=actMonthlyList.toString();
				
				String expMonthlyList= excelReader.getCellData(xlSheetName, 336, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 336, 7, actMonthly);
				/*ArrayList<String> expMonthlyList = new ArrayList<String>();
				expMonthlyList.add("January 2020");
				expMonthlyList.add("");
				expMonthlyList.add("36.00");
				expMonthlyList.add("");
				expMonthlyList.add("36.00");
				//expMonthlyList.add(Row2Month);
				expMonthlyList.add("November 2021");
				expMonthlyList.add("36.00");
				expMonthlyList.add("57.00");
				expMonthlyList.add("24.00");
				expMonthlyList.add("69.00");
				expMonthlyList.add("");
				expMonthlyList.add("");
				expMonthlyList.add("93.00");
				expMonthlyList.add("24.00");
				expMonthlyList.add("");
				*/
				System.out.println(actMonthly);
				System.out.println(expMonthlyList);
				
				for(int i=0; i<backTrackMonthlyColumnsCount; i++)
				{
					String data = backTrackItemDetailsMonthsTableList.get(i).getText();
					
					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 337, 5))) 
					{
						getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
						break;
					}
				}
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
				
				int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
				ArrayList<String> actDayList = new ArrayList<String>();
				for(int i=0; i<backTrackDayColumnsCount; i++)
				{
					String data = backTrackItemDetailsDaysTableList.get(i).getText();
					actDayList.add(data);
				}
				
				DateFormat dt = new SimpleDateFormat("d");
				String Row1Date = dt.format(cal.getTime());
				
				String actDay=actDayList.toString();
				
				
				
				/*ArrayList<String> expDayList = new ArrayList<String>();
				//expDayList.add(Row1Date);
				expDayList.add("23");
				expDayList.add("36.00");
				expDayList.add("57.00");
				expDayList.add("24.00");
				expDayList.add("69.00");
				expDayList.add("");
				expDayList.add("");
				expDayList.add("57.00");
				expDayList.add("24.00");
				expDayList.add("");*/
				String expDayList=excelReader.getCellData(xlSheetName, 337, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 337, 7, actDay);

				
				System.out.println(actDay);
				System.out.println(expDayList);
				
				for(int i=0; i<backTrackDayColumnsCount; i++)
				{
					String data = backTrackItemDetailsDaysTableList.get(i).getText();
					
					if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 338, 5))) 
					{
						getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
						break;
					}
				}
				
				Thread.sleep(2000);
				
				getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
				
				int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
				ArrayList<String> actVouchersList = new ArrayList<String>();
				for(int i=0; i<backTrackVouchersColumnsCount; i++)
				{
					String data = backTrackItemDetailsVouchersTableList.get(i).getText();
					actVouchersList.add(data);
				}
				
				String actVoucher=actVouchersList.toString();
			String expVouchersList=excelReader.getCellData(xlSheetName, 338, 6);
			 excelReader.setCellData(xlfile, xlSheetName, 338, 7, actVoucher);

				
				System.out.println(actVoucher);
				System.out.println(expVouchersList);
				
				
				for(int i=0; i<backTrackVouchersColumnsCount; i++)
				{
					String data = backTrackItemDetailsVouchersTableList.get(i).getText();
					
					if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 339, 5))) 
					{
						getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
					}
				}
				
				Thread.sleep(2000);
				
				ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
				
				int actCount = getDriver().getWindowHandles().size();
				String actOpenWindowsCount=String.valueOf(actCount);
				
				String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 339, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 339, 7, actOpenWindowsCount);
				
				

			 	getDriver().switchTo().window(openTabs.get(1));
			 	
			 	checkValidationMessage("");
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
				String actDocNo=documentNumberTxt.getAttribute("value");
				String expDocNo=excelReader.getCellData(xlSheetName, 340, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 340, 7, actDocNo);
					
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
				String actCustomer=customerAccountTxt.getAttribute("value");
				String expCustomer=excelReader.getCellData(xlSheetName, 341, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 341, 7, actCustomer);
				
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
			 	String actWarehouse = select1stRow_1stColumn.getText();
			 	String expWarehouse = excelReader.getCellData(xlSheetName, 342, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 342, 7, actWarehouse);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
			 	String actItem = select1stRow_2ndColumn.getText();
			 	String expItem = excelReader.getCellData(xlSheetName, 343, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 343, 7, actItem);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
			 	String actTaxCode = select1stRow_3rdColumn.getText();
			 	String expTaxCode = excelReader.getCellData(xlSheetName, 344, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 344, 7, actTaxCode);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
			 	String actSalesAccount = select1stRow_4thColumn.getText();
			 	String expSalesAccount = excelReader.getCellData(xlSheetName, 345, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 345, 7, actSalesAccount);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
			 	String actUnits = select1stRow_5thColumn.getText();
			 	String expUnits = excelReader.getCellData(xlSheetName, 346, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 346, 7, actUnits);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
			 	String actQuantity = select1stRow_11thColumn.getText();
			 	String expQuantity = excelReader.getCellData(xlSheetName, 347, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 347, 7, actQuantity);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
			 	String actLSalesOrder = select1stRow_12thColumn.getText();
			 	String expLSalesOrder =excelReader.getCellData(xlSheetName, 348, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 348, 7, actLSalesOrder);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
			 	String actRate = select1stRow_14thColumn.getText();
			 	String expRate = excelReader.getCellData(xlSheetName, 349, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 349, 7, actRate);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
			 	String actGross = select1stRow_15thColumn.getText();
			 	String expGross = excelReader.getCellData(xlSheetName, 350, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 350, 7, actGross);
			 	
			 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_20thColumn));
			 	String actBin = select1stRow_20thColumn.getText();
			 	String expBin =excelReader.getCellData(xlSheetName, 351, 6);
				 excelReader.setCellData(xlfile, xlSheetName, 351, 7, actBin);
				 
			 	
				System.out.println("**********************checkVirtualStockAnalysisReportBackTrackOption*********************");
				 
				System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
				
				System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
				System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
				System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
				System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
				System.out.println("Tax Code         : "+actTaxCode       +"  Value Expected  "+expTaxCode);
				System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
				System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
				System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
				System.out.println("L Sales Order    : "+actLSalesOrder   +"  Value Expected  "+expLSalesOrder);
				System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
				System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
				System.out.println("Bin              : "+actBin           +"  Value Expected  "+expBin);
				
				getDriver().switchTo().window(openTabs.get(1)).close();
				
				Thread.sleep(1000);
				
				getDriver().switchTo().window(openTabs.get(0));
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
				BackTrackItemDetailsBackBtn.click();
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
				BackTrackMonthDetailsBackBtn.click();
				
				Thread.sleep(1000);
				
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
				BackTrackDateDetailsBackBtn.click();
				
				Thread.sleep(1000);
				
				if(actMonthly.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucher.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
						   && actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
						   && actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
						   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesOrder.equalsIgnoreCase(expLSalesOrder)
						   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) && actBin.equalsIgnoreCase(expBin)) 
					{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					 excelReader.setCellData(xlfile, xlSheetName, 336, 8, resPass);
					return true;
				}
				else
				{
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
					report_CloseBtn.click();
					getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
					sl_CloseBtn.click();
					excelReader.setCellData(xlfile, xlSheetName, 336, 8, resFail);
					return false;
				}
			}
			
			
			
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[9]")
	private static WebElement stockValuation;
	
	public boolean checkStockValuationReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockValuation));
		stockValuation.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(  excelReader.getCellData(xlSheetName, 358, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List =   excelReader.getCellData(xlSheetName, 358, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 358, 7, actRow1List);

		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 359, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 359, 7, actRow2List);
		
		System.out.println("*********************************checkStockValuationReport*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 358, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 358, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	@FindBy (xpath="//span[@id='id_Item_spanplus']")
	private static WebElement sv_CustomizationFieldsItemExpansion;
	
	@FindBy (xpath="//*[@id='id_Item_main_40']/span/span/i")
	private static WebElement sv_CustomizationFieldsPurchaseVarianceExpansion;
	
	@FindBy (xpath="//li[@id='id_Item_main_40_0']")
	private static WebElement sv_CustomizationFieldsPurchaseVarianceName;
	
	
	public boolean checkStockValuationCustomizationReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockValuation));
		stockValuation.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 363, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_ReportCustomizeBtn));
		report_ReportCustomizeBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sv_CustomizationFieldsItemExpansion));
		sv_CustomizationFieldsItemExpansion.click();
		
		Thread.sleep(1000);
		
		/*JavascriptExecutor js=(JavascriptExecutor)getDriver();
		js.executeScript("arguments[0].scrollIntoView();", sv_CustomizationFieldsPurchaseVarianceExpansion);
		*/
		
		
		getAction().moveToElement(sv_CustomizationFieldsPurchaseVarianceExpansion).build().perform();
		
		Thread.sleep(1000);
		
		sv_CustomizationFieldsPurchaseVarianceExpansion.click();
		
		
		getAction().moveToElement(sv_CustomizationFieldsPurchaseVarianceName).build().perform();
		
		Thread.sleep(1000);
		getAction().doubleClick(sv_CustomizationFieldsPurchaseVarianceName).build().perform();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeSaveBtn));
		osr_customizeSaveBtn.click();
		
		Thread.sleep(1000);
		String expMessage =  excelReader.getCellData(xlSheetName, 363, 6);
	    
	    String actMessage = checkValidationMessage(expMessage);
	    excelReader.setCellData(xlfile, xlSheetName, 363, 7, actMessage);
		
		Thread.sleep(2000);
		
		/*if (sl_OkBtn.isDisplayed()==true) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(4000);
		}
	    else
	    {
	    	if (osr_customizeCloseBtn.isDisplayed()==true) 
	    	{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeCloseBtn));
				osr_customizeCloseBtn.click();
				
				Thread.sleep(4000);
			}
	    }*/
		
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =   excelReader.getCellData(xlSheetName, 364, 6);
      
        excelReader.setCellData(xlfile, xlSheetName, 364, 7, actRow1List);
		
		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 365, 6);
      
        excelReader.setCellData(xlfile, xlSheetName, 365, 7, actRow2List);
		
		
		System.out.println("************************************checkStockValuationCustomizationReport********************************************");
		
		System.out.println(actMessage);
		System.out.println(expMessage);
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_ReportCustomizeBtn));
		report_ReportCustomizeBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CustomizationDeleteLayoutBtn));
		report_CustomizationDeleteLayoutBtn.click();
	
		Thread.sleep(1000);
		
		getWaitForAlert();
		getAlert().accept();
		
		String expDeleteMessage = excelReader.getCellData(xlSheetName, 366, 6);
                   
	    String actDeleteMessage = checkValidationMessage(expDeleteMessage);
	    excelReader.setCellData(xlfile, xlSheetName, 366, 7, actDeleteMessage);
	    
	    Thread.sleep(2000);
	    
	    
	    int reportsHeaderListCount = reportsHeaderList.size();
		ArrayList<String> reportsHeaderListArray = new ArrayList<String>();
		for(int i=0;i<reportsHeaderListCount;i++)
		{
			String data = reportsHeaderList.get(i).getText();
			reportsHeaderListArray.add(data);
		}
		String actHeaderList = reportsHeaderListArray.toString();
		String expHeaderList = excelReader.getCellData(xlSheetName, 367, 6);
         excelReader.setCellData(xlfile, xlSheetName, 367, 7, actHeaderList);

	    
	    System.out.println(actHeaderList);
	    System.out.println(expHeaderList);
	    
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
	    report_CloseBtn.click();
		
		Thread.sleep(2000);
	    
		if(actMessage.equalsIgnoreCase(expMessage)
			&& actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List)
			&& actDeleteMessage.equalsIgnoreCase(expDeleteMessage)
			&& actHeaderList.equalsIgnoreCase(expHeaderList))
		{
			excelReader.setCellData(xlfile, xlSheetName, 363, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 363, 8, resFail);
			return false;
		}
	}
	
	
	
	@FindBy (xpath="//*[@id='id_reportmenudisplay']/ul/li[14]")
	private static WebElement reportWithLevel_CloseBtn;

	
	
	@FindBy (xpath="//*[@id='dvReportInputs']/div[8]/label/span")
	private static WebElement sl_DisplayLinearReportChkBox;
	
	
	
	
	public boolean checkStockValuationReportWithStockValuationDisplayLinearReportIncludeServiceTypeItemOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		   excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockValuation));
		stockValuation.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 373, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		Select s1 = new Select(sl_StockValuationDropdown);
		s1.selectByIndex(2);
		
		getAction().moveToElement(sl_DisplayLinearReportChkBox).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DisplayLinearReportChkBox));
		sl_DisplayLinearReportChkBox.click();
		
		Thread.sleep(1000);
		
		getAction().moveToElement(sl_IncludeServiceTypeItemChkBox).build().perform();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_IncludeServiceTypeItemChkBox));
		sl_IncludeServiceTypeItemChkBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 373, 6);
		 excelReader.setCellData(xlfile, xlSheetName,373, 7, actRow1List);
		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 374, 6);
		 excelReader.setCellData(xlfile, xlSheetName,374, 7, actRow2List);
		
		
		
		int reportsByWarehouseRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsByWarehouseRow3ListArray.add(data);
		}
		String actRow3List = reportsByWarehouseRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 375, 6);
		 excelReader.setCellData(xlfile, xlSheetName,375, 7, actRow3List);
		
		
		System.out.println("************checkStockValuationReportWithStockValuationDisplayLinearReportIncludeServiceTypeItemOption***********");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 373, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 373, 8, resFail);
			return false;
		}
	}
	
	
	
	@FindBy (xpath="//*[@id='dvmainBacktracking']/div[2]/div/table/tbody/tr/td")
	private static List<WebElement> backTrackItemDetailsMonthsTableList;
	
	@FindBy (xpath="(//*[@id='dvmainBacktracking']/div[2]/div/table)[1]")
	private static WebElement backTrackItemDetailsMonthsTable;
	
				
	@FindBy (xpath="(//*[@id='dvmainBacktracking']/div[2]/div/table)[2]/tbody/tr/td")
	private static List<WebElement> backTrackItemDetailsDaysTableList;
	
	@FindBy (xpath="(//*[@id='dvmainBacktracking']/div[2]/div/table)[2]")
	private static WebElement backTrackItemDetailsDaysTable;
	
	/*@FindBy (xpath="//div[@id='dvReportRenderBacktracking']/div[2]/div[1]/div[2]/table/tbody/tr[1]/td[1]")
	private static WebElement backTrackItemDetailsDays1stRow1stCol;*/
	
	
	
	@FindBy (xpath="(//*[@id='dvmainBacktracking']/div[2]/div/table)[3]/tbody/tr/td")
	private static List<WebElement> backTrackItemDetailsVouchersTableList;
	
	@FindBy (xpath="(//*[@id='dvmainBacktracking']/div[2]/div/table)[3]")
	private static WebElement backTrackItemDetailsVouchersTable;
	
	
	
	
	@FindBy (xpath="//*[@id='dvReportDetails']/div/table/tbody/tr/td[2]")
	private static List<WebElement> reportTableColumn1RowsList;
	
	


	
	public boolean checkStockValuationReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockValuation));
		stockValuation.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 378, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		Select s1 = new Select(sl_StockValuationDropdown);
		s1.selectByIndex(2);
		
		getAction().moveToElement(sl_DisplayLinearReportChkBox).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DisplayLinearReportChkBox));
		sl_DisplayLinearReportChkBox.click();
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		*/
		int reportItemGridListCount=reportTableColumn1RowsList.size();
		System.err.println("List"+reportItemGridListCount);
	   
		for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             System.err.println(data);
             
             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 379, 5)))
             {
                 int count =i+2;
                // System.err.println("C:"+count);
                 for (int j = i; j <count; j++) 
     		     {
                	 String data1=reportTableColumn1RowsList.get(j).getText();
                	 
                	 System.err.println(data1);
                	 if (data1.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 380, 5)))
                     {
                		 System.err.println("Item Name : "+data1);
                		 reportTableColumn1RowsList.get(j).click();
                		 Thread.sleep(1000);
                		 break;
                     }
     		     }
             }
	      }
	    
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		System.err.println(Row2Month);
		
		String actMonth=actMonthlyList.toString();
		
	/*	ArrayList<String> expMonthlyList = new ArrayList<String>();
		expMonthlyList.add("January 2020");
		expMonthlyList.add("");
		expMonthlyList.add("36.00");
		expMonthlyList.add("");
		expMonthlyList.add("36.00");
		//expMonthlyList.add(Row2Month);
		expMonthlyList.add("November 2021");
		expMonthlyList.add("36.00");
		expMonthlyList.add("96.00");
		expMonthlyList.add("36.00");
		expMonthlyList.add("96.00");
		expMonthlyList.add("");
		expMonthlyList.add("");
		expMonthlyList.add("132.00");
		expMonthlyList.add("36.00");
		expMonthlyList.add("");
		*/
		String expMonthlyList= excelReader.getCellData(xlSheetName, 380, 6);
		excelReader.setCellData(xlfile, xlSheetName, 380, 7, actMonth);
		
		System.out.println(actMonthlyList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 381, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		String actDay=actDayList.toString();
		
	/*	ArrayList<String> expDayList = new ArrayList<String>();
		//expDayList.add(Row1Date);
		expDayList.add("23");
		expDayList.add("36.00");
		expDayList.add("96.00");
		expDayList.add("36.00");
		expDayList.add("96.00");
		expDayList.add("");
		expDayList.add("");
		expDayList.add("96.00");
		expDayList.add("36.00");
		expDayList.add("");
		*/
		String expDayList=excelReader.getCellData(xlSheetName, 381, 6);
		excelReader.setCellData(xlfile, xlSheetName, 381, 7, actDay);
		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 382, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucher=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 382, 6);
		  excelReader.setCellData(xlfile, xlSheetName, 382, 7, actVoucher);
		  
		System.out.println(actVouchersList);
		System.out.println(expVouchersList);
		
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 383, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		int actCount = getDriver().getWindowHandles().size();
		String actOpenWindowsCount=String.valueOf(actCount);
		
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 383, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 383, 7, actOpenWindowsCount);
		

	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		
		String expDocNo= excelReader.getCellData(xlSheetName, 384, 6);
        excelReader.setCellData(xlfile, xlSheetName, 384, 7, actDocNo);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomer=customerAccountTxt.getAttribute("value");
		String expCustomer= excelReader.getCellData(xlSheetName, 385, 6);
        excelReader.setCellData(xlfile, xlSheetName, 385, 7, actCustomer);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	String actWarehouse = select1stRow_1stColumn.getText();
	 	String expWarehouse =  excelReader.getCellData(xlSheetName, 386, 6);
        excelReader.setCellData(xlfile, xlSheetName, 386, 7, actWarehouse);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	 	String actItem = select1stRow_2ndColumn.getText();
	 	String expItem =  excelReader.getCellData(xlSheetName, 387, 6);
        excelReader.setCellData(xlfile, xlSheetName, 387, 7, actItem);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
	 	String actTaxCode = select1stRow_3rdColumn.getText();
	 	String expTaxCode =  excelReader.getCellData(xlSheetName, 388, 6);
        excelReader.setCellData(xlfile, xlSheetName, 388, 7, actTaxCode);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
	 	String actSalesAccount = select1stRow_4thColumn.getText();
	 	String expSalesAccount =  excelReader.getCellData(xlSheetName, 389, 6);
        excelReader.setCellData(xlfile, xlSheetName, 389, 7, actSalesAccount);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	 	String actUnits = select1stRow_5thColumn.getText();
	 	String expUnits =  excelReader.getCellData(xlSheetName, 390, 6);
        excelReader.setCellData(xlfile, xlSheetName, 390, 7, actUnits);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	 	String actQuantity = select1stRow_11thColumn.getText();
	 	String expQuantity =  excelReader.getCellData(xlSheetName, 391, 6);
        excelReader.setCellData(xlfile, xlSheetName, 391, 7, actQuantity);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
	 	String actLSalesOrder = select1stRow_12thColumn.getText();
	 	String expLSalesOrder =  excelReader.getCellData(xlSheetName, 392, 6);
        excelReader.setCellData(xlfile, xlSheetName, 392, 7, actLSalesOrder);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	 	String actRate = select1stRow_14thColumn.getText();
	 	String expRate =  excelReader.getCellData(xlSheetName, 393, 6);
        excelReader.setCellData(xlfile, xlSheetName, 393, 7, actRate);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
	 	String actGross = select1stRow_15thColumn.getText();
	 	String expGross =  excelReader.getCellData(xlSheetName, 394, 6);
        excelReader.setCellData(xlfile, xlSheetName, 394, 7, actGross);
	 	
		System.out.println("**********************checkStockValuationReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
		System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
		System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
		System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
		System.out.println("Tax Code         : "+actTaxCode       +"  Value Expected  "+expTaxCode);
		System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
		System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
		System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
		System.out.println("L Sales Order    : "+actLSalesOrder   +"  Value Expected  "+expLSalesOrder);
		System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
		System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonth.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucher.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
			   && actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
			   && actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
			   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesOrder.equalsIgnoreCase(expLSalesOrder)
			   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 373, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 373, 8, resFail);
			return false;
		}
	}
	
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[10]")
	private static WebElement stockReportByTag;
	
	public boolean checkStockReportbyTagReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();
		 getDriver().navigate().refresh();
			Thread.sleep(2000);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockReportByTag));
		stockReportByTag.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 399, 5)) || actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 400, 5)) || actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 401, 5)) 
					|| actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 402, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
	
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 402, 6);
		excelReader.setCellData(xlfile, xlSheetName, 402, 7, actRow1List);

		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 403, 6);
		excelReader.setCellData(xlfile, xlSheetName, 403, 7, actRow2List);

		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 404, 6);
		excelReader.setCellData(xlfile, xlSheetName, 404, 7, actRow3List);

		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 405, 6);
		excelReader.setCellData(xlfile, xlSheetName, 405, 7, actRow4List);

		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = excelReader.getCellData(xlSheetName, 406, 6);
		excelReader.setCellData(xlfile, xlSheetName, 406, 7, actRow5List);


		
		System.out.println("************************************checkStockReportbyTagReports********************************************");
	
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) /*&& actRow3List.equalsIgnoreCase(expRow3List) 
			&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List)*/)
		{
			 excelReader.setCellData(xlfile, xlSheetName, 399, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 399, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	public boolean checkStockReportbyTagReportPrintOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ReportPrintBtn));
		sl_ReportPrintBtn.click();
		
		getWaitForAlert();
		Thread.sleep(1000);
		
		getAlert().accept();
		
		Thread.sleep(3000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		int actCount = getDriver().getWindowHandles().size();
		String actOpenWindowsCount=String.valueOf(actCount);
		
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 412, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 412, 7, actOpenWindowsCount);
		
		

	 	getDriver().switchTo().window(openTabs.get(0));
	 	
	 	Thread.sleep(1000);
	 	
	 	getDriver().switchTo().window(openTabs.get(1)).close();
	 	
	 	getDriver().switchTo().window(openTabs.get(0));
		
		System.out.println("***************************checkStockReportbyTagReportPrintOption*********************************");
		
		System.out.println("openTabs"+openTabs);
		System.out.println("Open Tabs Count : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
		
		Thread.sleep(1000);
		
		if(actOpenWindowsCount.equals(expOpenWindowsCount))
		{
			excelReader.setCellData(xlfile, xlSheetName, 412, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 412, 8, resFail);
			return false;
		}
	}

	
	
	
	
	
	
	public boolean checkStockReportbyTagReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockReportByTag));
		stockReportByTag.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 418, 5)) || actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 419, 5)) || actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 420, 5)) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 421, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		int reportItemGridListCount=reportTableColumn1RowsList.size();
	    for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 422, 5)))
             {
        		 reportTableColumn1RowsList.get(i).click();
        		 Thread.sleep(1000);
        		 break;
             }
             
		     }
	  
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		
		String actMonth=actMonthlyList.toString();
		
		String expMonthlyList=excelReader.getCellData(xlSheetName, 422, 6);
		excelReader.setCellData(xlfile, xlSheetName, 422, 7, actMonth);
		
		System.out.println(actMonth);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 423, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		String actDay=actDayList.toString();
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String expDayList=excelReader.getCellData(xlSheetName, 423, 6);
		excelReader.setCellData(xlfile, xlSheetName, 423, 7, actDay);
		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 424, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucher=actVouchersList.toString();
		String expVouchersList=excelReader.getCellData(xlSheetName, 424, 6);
		excelReader.setCellData(xlfile, xlSheetName, 424, 7, actVoucher);
		
		System.out.println(actVoucher);
		System.out.println(expVouchersList);
		
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 425, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		int actCount = getDriver().getWindowHandles().size();
		String actOpenWindowsCount=String.valueOf(actCount);
		
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 426, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 427, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(1));
	 /*	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo="1";
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomer=customerAccountTxt.getAttribute("value");
		String expCustomer="Customer A";
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	String actWarehouse = select1stRow_1stColumn.getText();
	 	String expWarehouse = "HYDERABAD";
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	 	String actItem = select1stRow_2ndColumn.getText();
	 	String expItem = "STD RATE COGS ITEM";
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
	 	String actSalesAccount = select1stRow_3rdColumn.getText();
	 	String expSalesAccount = "Sales - Computers";
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
	 	String actUnits = select1stRow_4thColumn.getText();
	 	String expUnits = "Pcs";
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
	 	String actQuantity = select1stRow_6thColumn.getText();
	 	String expQuantity = "12.00";
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
	 	String actLSalesInvoice = select1stRow_7thColumn.getText();
	 	String expLSalesInvoice = "NDT50:6";
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	 	String actRate = select1stRow_8thColumn.getText();
	 	String expRate = "10.00";
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
	 	String actGross = select1stRow_9thColumn.getText();
	 	String expGross = "120.00";
	 	
		System.out.println("**********************checkStockReportbyTagReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
		System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
		System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
		System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
		System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
		System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
		System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
		System.out.println("L Sales Order    : "+actLSalesInvoice +"  Value Expected  "+expLSalesInvoice);
		System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
		System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
	*/	
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonth.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucher.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
			  /* && actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
			   && actItem.equalsIgnoreCase(expItem) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
			   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesInvoice.equalsIgnoreCase(expLSalesInvoice)
			   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross)*/) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reportWithLevel_CloseBtn));
			reportWithLevel_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 418, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reportWithLevel_CloseBtn));
			reportWithLevel_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 418, 8, resFail);
			return false;
		}
	}
	
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[11]")
	private static WebElement abcAnalysis;
	
	public boolean checkABCAnalysisReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 getDriver().navigate().refresh();
			Thread.sleep(2000);
			LogoutandLoginwithSU();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(abcAnalysis));
		abcAnalysis.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 440, 5)) || actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 441, 5)) || actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 442, 5)) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 443, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 443, 6);
		excelReader.setCellData(xlfile, xlSheetName, 443, 7, actRow1List);

		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 444, 6);
		excelReader.setCellData(xlfile, xlSheetName, 444, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 445, 6);
		excelReader.setCellData(xlfile, xlSheetName, 445, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 446, 6);
		excelReader.setCellData(xlfile, xlSheetName, 446, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =excelReader.getCellData(xlSheetName, 447, 6);
		excelReader.setCellData(xlfile, xlSheetName, 447, 7, actRow5List);
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List =excelReader.getCellData(xlSheetName, 448, 6);
		excelReader.setCellData(xlfile, xlSheetName, 448, 7, actRow6List);
		
		
		System.out.println("************************************checkABCAnalysisReports********************************************");
	
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
			&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 440, 8, resPass);

			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 440, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	@FindBy (xpath="//select[@id='RITCombobox__1']")
	private static WebElement ABCClassificationDropdown;
	
	@FindBy (xpath="//select[@id='RITCombobox__2']")
	private static WebElement ABCCriteriaDropdown;
	
	
	public boolean checkABCAnalysisReportWithClassificationAndCriteriaOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		 Select Classification = new Select(ABCClassificationDropdown);
		Classification.selectByIndex(1);
		
		Select Criteria = new Select(ABCCriteriaDropdown);
		Criteria.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 457, 6);
         excelReader.setCellData(xlfile, xlSheetName, 457, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 458, 6);
        excelReader.setCellData(xlfile, xlSheetName, 458, 7, actRow2List);
		
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =excelReader.getCellData(xlSheetName, 459, 6);
        excelReader.setCellData(xlfile, xlSheetName, 459, 7, actRow3List);
		
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 460, 6);
        excelReader.setCellData(xlfile, xlSheetName, 460, 7, actRow4List);
		
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = excelReader.getCellData(xlSheetName, 461, 6);
        excelReader.setCellData(xlfile, xlSheetName, 461, 7, actRow5List);
		
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List = excelReader.getCellData(xlSheetName, 462, 6);
        excelReader.setCellData(xlfile, xlSheetName, 462, 7, actRow6List);
		
		
		
		
		
		int reportsRow7ListCount = reportsRow7List.size();
		ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow7ListCount;i++)
		{
			String data = reportsRow7List.get(i).getText();
			reportsRow7ListArray.add(data);
		}
		String actRow7List = reportsRow7ListArray.toString();
		String expRow7List = excelReader.getCellData(xlSheetName, 463, 6);
        excelReader.setCellData(xlfile, xlSheetName, 463, 7, actRow7List);
		
		
		
		
		
		int reportsRow8ListCount = reportsRow8List.size();
		ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow8ListCount;i++)
		{
			String data = reportsRow8List.get(i).getText();
			reportsRow8ListArray.add(data);
		}
		String actRow8List = reportsRow8ListArray.toString();
		String expRow8List = excelReader.getCellData(xlSheetName, 464, 6);
        excelReader.setCellData(xlfile, xlSheetName, 464, 7, actRow8List);
		
		
		
		System.out.println("**************************checkABCAnalysisReportWithClassificationAndCriteriaOptions**********************************");
	
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		System.out.println(actRow7List);
		System.out.println(expRow7List);
		
		System.out.println(actRow8List);
		System.out.println(expRow8List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
			&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List)
			&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 457, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 457, 8, resFail);
			return false;
		}
	}
	
	
	
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[13]")
	private static WebElement stockAnalysisByBatch;
	
	public boolean checkStockAnalysisByBatchReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LogoutandLoginwithSU();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAnalysisByBatch));
		stockAnalysisByBatch.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 470, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=1;i<5;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 470, 6);
		excelReader.setCellData(xlfile, xlSheetName, 470, 7, actRow1List);

		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=1;i<5;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 471, 6);
		excelReader.setCellData(xlfile, xlSheetName, 471, 7, actRow2List);
		
		
		
		int reportsByWarehouseRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
		for(int i=1;i<5;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsByWarehouseRow3ListArray.add(data);
		}
		String actRow3List = reportsByWarehouseRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 472, 6);
		excelReader.setCellData(xlfile, xlSheetName, 472, 7, actRow3List);
		
		
		
		int reportsByWarehouseRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsByWarehouseRow4ListArray = new ArrayList<String>();
		for(int i=1;i<5;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsByWarehouseRow4ListArray.add(data);
		}
		String actRow4List = reportsByWarehouseRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 473, 6);
		excelReader.setCellData(xlfile, xlSheetName, 473, 7, actRow4List);
		
		
		
		int reportsByWarehouseRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsByWarehouseRow5ListArray = new ArrayList<String>();
		for(int i=1;i<5;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsByWarehouseRow5ListArray.add(data);
		}
		String actRow5List = reportsByWarehouseRow5ListArray.toString();
		String expRow5List = excelReader.getCellData(xlSheetName, 474, 6);
		excelReader.setCellData(xlfile, xlSheetName, 474, 7, actRow5List);
		
		
		
		/*int reportsByWarehouseRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsByWarehouseRow6ListArray = new ArrayList<String>();
		for(int i=1;i<5;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsByWarehouseRow6ListArray.add(data);
		}
		String actRow6List = reportsByWarehouseRow6ListArray.toString();
		String expRow6List = "[Grand Total, 86.00, 3.77, 67.25]";*/
		
		
		System.out.println("*********************************checkStockAnalysisByBatchReport*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		/*System.out.println(actRow6List);
		System.out.println(expRow6List);*/
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
			&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) /*&& actRow6List.equalsIgnoreCase(expRow6List)*/)
		{
			excelReader.setCellData(xlfile, xlSheetName, 470, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 470, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_580_0_DefaultFilter_0']")
	private static WebElement sab_DefaultFilterTxt;
	
	@FindBy (xpath="//*[@id='FilterFields_580_0']/ul[5]/li[1]/div/label/span")
	private static WebElement sab_report_FilterWarehouseNameChkbox;

	
	public boolean checkStockAnalysisByBatchReportFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
			report_FilterExpandBtn.click();
			Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterWarehouseExpandBtn));
		report_FilterWarehouseExpandBtn.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sab_report_FilterWarehouseNameChkbox));
		
		if (report_FilterWarehouseNameChkbox.isSelected()==false) 
		{
			sab_report_FilterWarehouseNameChkbox.click();
		}

		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sab_DefaultFilterTxt));
		sab_DefaultFilterTxt.click();
		sab_DefaultFilterTxt.sendKeys(excelReader.getCellData(xlSheetName, 479, 5));
		
		Thread.sleep(2000);
		
		sab_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
		
		//boolean actReportTable = reportsTable.getText().isEmpty();
		//boolean expReportTable = true;
		
		String actReportTable=String.valueOf(reportsTable.getText().isEmpty());
		String expReportTable=excelReader.getCellData(xlSheetName, 479, 6);
		excelReader.setCellData(xlfile, xlSheetName, 479, 7, actReportTable);

				
		System.out.println("*********************************checkStockAnalysisByBatchReportFilterOption*****************************************");
		
		System.out.println(actReportTable);
		
		System.out.println(expReportTable);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sab_report_FilterWarehouseNameChkbox));
		
		if (sab_report_FilterWarehouseNameChkbox.isSelected()==true) 
		{
			sab_report_FilterWarehouseNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
		
		//int actRowCount = reportsRowsCount.size();
		//int expRowCount = 5; 
		
		String actRowCount=String.valueOf(reportsRowsCount.size());
		
		String expRowCount = excelReader.getCellData(xlSheetName, 480, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 480, 7, actRowCount);
		
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		//getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
		//report_CloseBtn.click();
		
		if(actReportTable.equals(expReportTable) && actRowCount.equals(expRowCount))
		{
			excelReader.setCellData(xlfile, xlSheetName, 479, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 479, 8, resFail);

			return false;
		}
	}
	
	
	
	
	
	
	
	@FindBy (xpath="//select[@id='RITCombobox__1']")
	private static WebElement stockTypeDropdown;
	
	@FindBy(xpath="//div[@id='dvReportDetails']/div/table/tbody")
	private static WebElement reportsTable;
	
	
	public boolean checkStockAnalysisByBatchReportWithStockTypeExpiredBatches() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAnalysisByBatch));
		stockAnalysisByBatch.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 524, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTypeDropdown));
		Select s1=new Select(stockTypeDropdown);
		s1.selectByIndex(2);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		//boolean actReportTable = reportsTable.getText().isEmpty();
		//boolean expReportTable = true;
		
		//String actReportTable=String.valueOf(reportsTable.getText().isEmpty());
		//String expReportTable=excelReader.getCellData(xlSheetName, 13, 5);
		
		
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 524, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 524, 7, actRow1List);
		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List =  excelReader.getCellData(xlSheetName, 525, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 525, 7, actRow2List);
		
		
		
		int reportsByWarehouseRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsByWarehouseRow3ListArray.add(data);
		}
		String actRow3List = reportsByWarehouseRow3ListArray.toString();
		String expRow3List =  excelReader.getCellData(xlSheetName, 526, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 526, 7, actRow3List);
		
		
		
		int reportsByWarehouseRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsByWarehouseRow4ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsByWarehouseRow4ListArray.add(data);
		}
		String actRow4List = reportsByWarehouseRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 527, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 527, 7, actRow4List);
			
		
		
		
		int reportsByWarehouseRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsByWarehouseRow5ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsByWarehouseRow5ListArray.add(data);
		}
		String actRow5List = reportsByWarehouseRow5ListArray.toString();
		String expRow5List = excelReader.getCellData(xlSheetName, 528, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 528, 7, actRow5List);
			
		
		
		
		
		
	
		System.out.println("*********************************checkStockAnalysisByBatchReportWithStockTypeExpiredBatches*****************************************");
		
		//System.out.println(actReportTable);
		//System.out.println(expReportTable);
		
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
			&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) )
		{
			excelReader.setCellData(xlfile, xlSheetName, 524, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 524, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	public boolean checkStockAnalysisByBatchReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		int reportItemGridListCount=reportTableColumn1RowsList.size();
	    for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             String data1=reportTableColumn2RowsList.get(i).getText();
             
             System.out.println(data+"  "+data1);
             
             if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 488, 5)) /*&& data1.equalsIgnoreCase("28.00")*/)
             {
        		 reportTableColumn1RowsList.get(i).click();
        		 Thread.sleep(1000);
        		 break;
             }
		     }
	    
	    //Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		String actMonthList=actMonthlyList.toString();
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String expMonthlyList=excelReader.getCellData(xlSheetName, 488, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 488, 7,actMonthList);

		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 489, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}

		//Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList=excelReader.getCellData(xlSheetName, 489, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 489, 7, actDay);
		
		System.out.println(actDayList);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(
	                excelReader.getCellData(xlSheetName, 490, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList=excelReader.getCellData(xlSheetName, 490, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 490, 7, actVoucherList);
		
		System.out.println(actVouchersList);
		System.out.println(expVouchersList);
		
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 491, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 491, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 491, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(1));
	 
	 	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 492, 6);
		excelReader.setCellData(xlfile, xlSheetName, 492, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 493, 6);
		excelReader.setCellData(xlfile, xlSheetName, 493, 7, actVendor);

		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 494, 6);
		excelReader.setCellData(xlfile, xlSheetName, 494, 7, actCurrency);

		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 495, 6);
		excelReader.setCellData(xlfile, xlSheetName, 495, 7, actDepartment);

		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 496, 6);
		excelReader.setCellData(xlfile, xlSheetName, 496, 7, actPlaceOfSupply);


		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 497, 6);
		excelReader.setCellData(xlfile, xlSheetName, 497, 7, actJurisdiction);


		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 498, 6);
		excelReader.setCellData(xlfile, xlSheetName, 498, 7, actR1Warehouse);

		
		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 499, 6);
		excelReader.setCellData(xlfile, xlSheetName, 499, 7, actR1Item);

		
		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 500, 6);
		excelReader.setCellData(xlfile, xlSheetName, 500, 7, actR1TaxCode);

		
		String actR1PurchaseAccount = select1stRow_4thColumn.getText();
		String expR1PurchaseAccount = excelReader.getCellData(xlSheetName, 501, 6);
		excelReader.setCellData(xlfile, xlSheetName, 501, 7, actR1PurchaseAccount);

		
		String actR1Units = select1stRow_5thColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 502, 6);
		excelReader.setCellData(xlfile, xlSheetName, 502, 7, actR1Units);

		
		String actR1Quantity = select1stRow_9thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 503, 6);
		excelReader.setCellData(xlfile, xlSheetName, 503, 7, actR1Quantity);

		
		String actR1Rate = select1stRow_11thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 504, 6);
		excelReader.setCellData(xlfile, xlSheetName, 504, 7, actR1Rate);

		
		String actR1Gross = select1stRow_12thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 505, 6);
		excelReader.setCellData(xlfile, xlSheetName, 505, 7, actR1Gross);

		
		String actR1Batch = select1stRow_16thColumn.getText();
		String expR1Batch = excelReader.getCellData(xlSheetName, 506, 6);
		excelReader.setCellData(xlfile, xlSheetName, 506, 7, actR1Batch);


		String actR1Expiry = select1stRow_18thColumn.getText();
		DateFormat dtf = new SimpleDateFormat("dd/MM/yyyy");
		Calendar R1cal=Calendar.getInstance();
		R1cal.add(Calendar.DATE, 2); 
		//String expR1Expiry=dtf.format(R1cal.getTime());
		
		String expR1Expiry=excelReader.getCellData(xlSheetName, 507, 6);
		excelReader.setCellData(xlfile, xlSheetName, 507, 7, actR1Expiry);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 508, 6);
		excelReader.setCellData(xlfile, xlSheetName, 508, 7, actR2Warehouse);
		
		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 509, 6);
		excelReader.setCellData(xlfile, xlSheetName, 509, 7, actR2Item);
		
		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode =excelReader.getCellData(xlSheetName, 510, 6);
		excelReader.setCellData(xlfile, xlSheetName, 510, 7, actR2TaxCode);
		
		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 511, 6);
		excelReader.setCellData(xlfile, xlSheetName, 511, 7, actR2PurchaseAccount);
		
		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 512, 6);
		excelReader.setCellData(xlfile, xlSheetName, 512, 7, actR2Units);
		
		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 513, 6);
		excelReader.setCellData(xlfile, xlSheetName, 513, 7, actR2Quantity);
		
		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 514, 6);
		excelReader.setCellData(xlfile, xlSheetName, 514, 7, actR2Rate);
		
		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 515, 6);
		excelReader.setCellData(xlfile, xlSheetName, 515, 7, actR2Gross);
		
		String actR2Batch = select2ndRow_16thColumn.getText();
		String expR2Batch = excelReader.getCellData(xlSheetName, 516, 6);
		excelReader.setCellData(xlfile, xlSheetName, 516, 7, actR2Batch);

		String actR2Expiry = select2ndRow_18thColumn.getText();
		Calendar R2cal=Calendar.getInstance();
		R2cal.add(Calendar.DATE, 5); 
		//String expR2Expiry=dtf.format(R2cal.getTime());
		
		String expR2Expiry=excelReader.getCellData(xlSheetName, 517, 6);
		excelReader.setCellData(xlfile, xlSheetName, 517, 7, actR2Expiry);
	 	
		System.out.println("**********************checkStockAnalysisByBatchReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs       : "+actOpenWindowsCount +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);
		
		System.out.println("*******************************ROW1**********************************");
		
		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("TaxCode         : "+actR1TaxCode			+"  Value Expected  "+expR1TaxCode);
		System.out.println("PurchaseAccount : "+actR1PurchaseAccount	+"  Value Expected  "+expR1PurchaseAccount);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		System.out.println("Batch           : "+actR1Batch				+"  Value Expected  "+expR1Batch);
		System.out.println("Expiry          : "+actR1Expiry				+"  Value Expected  "+expR1Expiry);
		
		System.out.println("*******************************ROW2**********************************");
		
		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("TaxCode         : "+actR2TaxCode			+"  Value Expected  "+expR2TaxCode);
		System.out.println("PurchaseAccount : "+actR2PurchaseAccount	+"  Value Expected  "+expR2PurchaseAccount);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Batch           : "+actR2Batch				+"  Value Expected  "+expR2Batch);
		System.out.println("Expiry          : "+actR2Expiry				+"  Value Expected  "+expR2Expiry);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
			  && actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
			   && actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
			   && actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
			   && actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actR1PurchaseAccount.equalsIgnoreCase(expR1PurchaseAccount) 
			   && actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) 
			   && actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Batch.equalsIgnoreCase(expR1Batch) && actR1Expiry.equalsIgnoreCase(expR1Expiry)
			   && actR2Warehouse.equalsIgnoreCase(expR2Warehouse) 
			   && actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
			   && actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
			   && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Batch.equalsIgnoreCase(expR2Batch) && actR2Expiry.equalsIgnoreCase(expR2Expiry)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 488, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 488, 8, resFail);
			return false;
		}
	}
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[14]//a")
	private static WebElement stockAnalysisByRMA;
	
	@FindBy (xpath="(//div[@id='dvReportDetails']/div/table/tbody)[1]/tr/td")
	private static List<WebElement> reportTableCells;
	
	public boolean checkStockAnalysisByRmaReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAnalysisByRMA));
		stockAnalysisByRMA.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 538, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportPage1Count = reportTableCells.size();
		ArrayList<String> reportPage1ListArray = new ArrayList<String>();
		for(int i=0;i<reportPage1Count;i++)
		{
			String data = reportTableCells.get(i).getText();
			reportPage1ListArray.add(data);
		}
		String actPage1List = reportPage1ListArray.toString();
		String expPage1List =  excelReader.getCellData(xlSheetName, 538, 6);
		excelReader.setCellData(xlfile, xlSheetName, 538, 7, actPage1List);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
		report_NextBtn.click();
		
		Thread.sleep(2000);
		
		int reportPage2Count = reportTableCells.size();
		ArrayList<String> reportPage2ListArray = new ArrayList<String>();
		for(int i=0;i<reportPage2Count;i++)
		{
			String data = reportTableCells.get(i).getText();
			reportPage2ListArray.add(data);
		}
		String actPage2List = reportPage2ListArray.toString();
		String expPage2List = excelReader.getCellData(xlSheetName, 539, 6);
		excelReader.setCellData(xlfile, xlSheetName, 539, 7, actPage2List);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
		report_NextBtn.click();
		
		Thread.sleep(2000);
		
		int reportPage3Count = reportTableCells.size();
		ArrayList<String> reportPage3ListArray = new ArrayList<String>();
		for(int i=0;i<reportPage3Count;i++)
		{
			String data = reportTableCells.get(i).getText();
			reportPage3ListArray.add(data);
		}
		String actPage3List = reportPage3ListArray.toString();
		String expPage3List = excelReader.getCellData(xlSheetName, 540, 6);
		excelReader.setCellData(xlfile, xlSheetName, 540, 7, actPage3List);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
		report_NextBtn.click();
		
		Thread.sleep(2000);
		
		int reportPage4Count = reportTableCells.size();
		ArrayList<String> reportPage4ListArray = new ArrayList<String>();
		for(int i=0;i<reportPage4Count;i++)
		{
			String data = reportTableCells.get(i).getText();
			reportPage4ListArray.add(data);
		}
		String actPage4List = reportPage4ListArray.toString();
		String expPage4List = excelReader.getCellData(xlSheetName, 541, 6);
		excelReader.setCellData(xlfile, xlSheetName, 541, 7, actPage4List);*/
		
		System.out.println("*****************************************checkStockAnalysisByRmaReports***************************************");
		
		System.out.println(actPage1List);
		System.out.println(expPage1List);
		
		System.out.println(actPage2List);
		System.out.println(expPage2List);
		
		System.out.println(actPage3List);
		System.out.println(expPage3List);
		
	/*	System.out.println(actPage4List);
		System.out.println(expPage4List);*/
		
		if(actPage1List.equalsIgnoreCase(expPage1List) && actPage2List.equalsIgnoreCase(expPage2List) && actPage3List.equalsIgnoreCase(expPage3List) 
			/*&& actPage4List.equalsIgnoreCase(expPage4List)*/)
		{
			excelReader.setCellData(xlfile, xlSheetName, 538, 8, resPass);

			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 538, 8, resFail);

			return false;
		}
	}
	
	
	
	
	
	
	
	
	public boolean checkStockAnalysisByRmaReportsPrintOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ReportPrintBtn));
		sl_ReportPrintBtn.click();
		
		getWaitForAlert();
		Thread.sleep(1000);
		getAlert().accept();
		
		Thread.sleep(3000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 549, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 549, 7, actOpenWindowsCount);


	 	getDriver().switchTo().window(openTabs.get(0));
	 	
	 	Thread.sleep(1000);
	 	
	 	getDriver().switchTo().window(openTabs.get(1)).close();
	 	
	 	getDriver().switchTo().window(openTabs.get(0));
		
		System.out.println("***************************checkStockAnalysisByRmaReportsPrintOption*********************************");
		
		System.out.println("openTabs"+openTabs);
		System.out.println("Open Tabs Count : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
		
		Thread.sleep(1000);
		
		if(actOpenWindowsCount.equals(expOpenWindowsCount))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 549, 8, resPass);
			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 549, 8, resFail);
			return false;
		}
	}

	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[15]//i")
	
//	@FindBy(xpath="//a[@id='540']//span[contains(text(),'Bins Report')]")
	private static WebElement binsReportMenu;
	
	//@FindBy(xpath="//span[contains(text(),'Stock Details by Bins')]")
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[15]/ul/li[2]")
	private static WebElement stockDetailsByBins;
	
	
	public boolean checkStockDetailsByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		//getDriver().navigate().refresh();
		
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);
		
		getAction().sendKeys(Keys.END).build().perform();
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", binsReportMenu);
		
		//getAction().moveToElement(binsReportMenu).build().perform();
		Thread.sleep(2000);
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binsReportMenu));
		binsReportMenu.click();
		
		Thread.sleep(3000);
		
		//((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);", stockDetailsByBins);
		
		getAction().moveToElement(stockDetailsByBins).build().perform();
		Thread.sleep(2000);
		((JavascriptExecutor)getDriver()).executeScript("arguments[0].click();", stockDetailsByBins);
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockDetailsByBins));
		stockDetailsByBins.click();
		*/
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 556, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 556, 6);
		excelReader.setCellData(xlfile, xlSheetName, 556, 7, actRow1List);


		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 557, 6);
		excelReader.setCellData(xlfile, xlSheetName, 557, 7, actRow2List);

		
		
		
		int reportsByWarehouseRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsByWarehouseRow3ListArray.add(data);
		}
		String actRow3List = reportsByWarehouseRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 558, 6);
		excelReader.setCellData(xlfile, xlSheetName, 558, 7, actRow3List);

		
		
		
		int reportsByWarehouseRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsByWarehouseRow4ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsByWarehouseRow4ListArray.add(data);
		}
		String actRow4List = reportsByWarehouseRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 559, 6);
		excelReader.setCellData(xlfile, xlSheetName, 559, 7, actRow4List);

		
		
		System.out.println("*********************************checkStockDetailsByBinsReports*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
			&& actRow4List.equalsIgnoreCase(expRow4List))
		
		{
			 excelReader.setCellData(xlfile, xlSheetName, 558, 8, resPass);
			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 558, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	@FindBy (xpath="//a[contains(text(),'Bins')]//i")
	private static WebElement report_FilterBinExpansion;
	
	
	@FindBy (xpath="//input[@id='5188']")
	private static WebElement binsReport_report_FilterBinNameChkbox;
	
	@FindBy(xpath="//*[@id='FilterFields_587_0']/ul[23]/li[1]/div/label/span")
	private static WebElement report_FilterBinNameChkbox;
	
	@FindBy (xpath="//input[@id='FOption_587_0_DefaultFilter_0']")
	private static WebElement sdb_DefaultFilterTxt;
	
	
	
	public boolean checkStockDetailsByBinsReportFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBinExpansion));
		report_FilterBinExpansion.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBinNameChkbox));
		
		Thread.sleep(1000);
		
		if (binsReport_report_FilterBinNameChkbox.isSelected()==false) 
		{
			report_FilterBinNameChkbox.click();
		}

		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sdb_DefaultFilterTxt));
		sdb_DefaultFilterTxt.click();
		sdb_DefaultFilterTxt.sendKeys( excelReader.getCellData(xlSheetName, 565, 5));
		
		Thread.sleep(2000);
		
		sdb_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 565, 6);
        excelReader.setCellData(xlfile, xlSheetName, 565, 7, actRow1List);
		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 566, 6);
        excelReader.setCellData(xlfile, xlSheetName, 566, 7, actRow2List);
		
		
		
		
		int reportsByWarehouseRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsByWarehouseRow3ListArray.add(data);
		}
		String actRow3List = reportsByWarehouseRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 567, 6);
        excelReader.setCellData(xlfile, xlSheetName, 567, 7, actRow3List);
		
		
		
		System.out.println("*********************************checkStockDetailsByBinsReportFilterOption*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBinNameChkbox));
		
		if (binsReport_report_FilterBinNameChkbox.isSelected()==true) 
		{
			report_FilterBinNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
		
		//int actRowCount = reportsRowsCount.size();
		//int expRowCount = 4; 
		
		String actRowCount=String.valueOf(reportsRowsCount.size());
		
		String expRowCount = excelReader.getCellData(xlSheetName, 568, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 568, 7, actRowCount);
		 
		 
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
				&& actRowCount.equals(expRowCount))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 565, 8, resPass);

			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 565, 8, resFail);
			return false;
		}
	}
	
	
	@FindBy(xpath="//input[@id='id_header_86']")
	private static WebElement  stockTransferheaderWarehouse1Txt;
	
	@FindBy(xpath="//tbody[@id='id_header_86_table_body']/tr/td[2]")
    private static List<WebElement> wareHouseHeaderListCount;
	
	@FindBy(xpath="//input[@id='id_body_87']")
	private static WebElement  enter_BodyWarehouse2Txt;
	
	@FindBy(xpath="//tbody[@id='id_body_87_table_body']/tr/td[2]")
    private static List<WebElement> wareHouseBodyListCount;
	
	@FindBy(xpath="//i[@id='id_transactionentry_workflow_popup_close']")
	private static WebElement  workFlowCloseBtn;
	
	@FindBy(xpath="//i[@class='icon-pick icon-font5']")
	private static WebElement batchPickOnFIFOIcon;
	
	@FindBy(xpath="//label[@id='div_Ok']")
	private static WebElement batchOkIcon;
	
	
	
	
	
	public boolean checkStockDetailsByBinsReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());

		int reportItemGridListCount=reportTableColumn1RowsList.size();
	    for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 574, 5)))
             {
        		 reportTableColumn1RowsList.get(i).click();
        		 Thread.sleep(1000);
        		 break;
             }
		     }
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		
		String expMonthlyList=excelReader.getCellData(xlSheetName, 574, 6);
		excelReader.setCellData(xlfile, xlSheetName, 574, 7, actMonthList);

		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			//excelReader.getCellData(xlSheetName, 575, 5)
			
			if (data.equalsIgnoreCase(Row2Month)) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList=excelReader.getCellData(xlSheetName, 575, 6);
		excelReader.setCellData(xlfile, xlSheetName, 575, 7, actDay);
		
		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(Row1Date)) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList=excelReader.getCellData(xlSheetName, 576, 6);
		excelReader.setCellData(xlfile, xlSheetName, 576, 7, actVoucherList);
		
		System.out.println(actVouchersList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 577, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 577, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 577, 7, actOpenWindowsCount);

		
		

	 	getDriver().switchTo().window(openTabs.get(1));
	 
	 	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 578, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 578, 7, actDocNo);

		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferheaderWarehouse1Txt));
		String actFromWarehouse = stockTransferheaderWarehouse1Txt.getAttribute("value");
		String expFromWarehouse = excelReader.getCellData(xlSheetName, 579, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 579, 7, actFromWarehouse);
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 580, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 580, 7, actR1Warehouse);
		
		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 581, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 581, 7, actR1Item);
		
		String actR1Units = select1stRow_3rdColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 582, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 582, 7, actR1Units);
		
		String actR1Quantity = select1stRow_4thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 583, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 583, 7, actR1Quantity);
		
		String actR1Rate = select1stRow_6thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 584, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 584, 7, actR1Rate);
		
		String actR1Gross = select1stRow_7thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 585, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 585, 7, actR1Gross);
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 586, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 586, 7, actR2Warehouse);
			
		
		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 587, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 587, 7, actR2Item);
			
		
		String actR2Units = select2ndRow_3rdColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 588, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 588, 7, actR2Units);
			
		
		String actR2Quantity = select2ndRow_4thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 589, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 589, 7, actR2Quantity);
			
		
		String actR2Rate = select2ndRow_6thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 590, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 590, 7, actR2Rate);
			
		
		String actR2Gross = select2ndRow_7thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 591, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 591, 7, actR2Gross);
			
		
		String actR2Batch = select2ndRow_8thColumn.getText();
		String expR2Batch = excelReader.getCellData(xlSheetName, 592, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 592, 7, actR2Batch);
			

		

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		String actR3Warehouse = select3rdRow_1stColumn.getText();
		String expR3Warehouse = excelReader.getCellData(xlSheetName, 593, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 593, 7, actR3Warehouse);
		
		String actR3Item = select3rdRow_2ndColumn.getText();
		String expR3Item = excelReader.getCellData(xlSheetName, 594, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 594, 7, actR3Item);
		
		String actR3Units = select3rdRow_3rdColumn.getText();
		String expR3Units = excelReader.getCellData(xlSheetName, 595, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 595, 7, actR3Units);
		
		String actR3Quantity = select3rdRow_4thColumn.getText();
		String expR3Quantity = excelReader.getCellData(xlSheetName, 596, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 596, 7, actR3Quantity);
		
		String actR3Rate = select3rdRow_6thColumn.getText();
		String expR3Rate = excelReader.getCellData(xlSheetName, 597, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 597, 7, actR3Rate);
		
		String actR3Gross = select3rdRow_7thColumn.getText();
		String expR3Gross = excelReader.getCellData(xlSheetName, 598, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 598, 7, actR3Gross);
		
		String actR3Bins = select3rdRow_9thColumn.getText();
		String expR3Bins = excelReader.getCellData(xlSheetName, 599, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 599, 7, actR3Bins);
		
		String actR3Bins2 = select3rdRow_10thColumn.getText();
		String expR3Bins2 = excelReader.getCellData(xlSheetName, 600, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 600, 7, actR3Bins2);
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
		String actR4Warehouse = select4thRow_1stColumn.getText();
		String expR4Warehouse = excelReader.getCellData(xlSheetName, 601, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 601, 7, actR4Warehouse);
		
		String actR4Item = select4thRow_2ndColumn.getText();
		String expR4Item = excelReader.getCellData(xlSheetName, 602, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 602, 7, actR4Item);
		
		String actR4Units = select4thRow_3rdColumn.getText();
		String expR4Units = excelReader.getCellData(xlSheetName, 603, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 603, 7, actR4Units);
		
		String actR4Quantity = select4thRow_4thColumn.getText();
		String expR4Quantity = excelReader.getCellData(xlSheetName, 604, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 604, 7, actR4Quantity);
		
		String actR4Rate = select4thRow_6thColumn.getText();
		String expR4Rate =excelReader.getCellData(xlSheetName, 605, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 605, 7, actR4Rate);
		
		String actR4Gross = select4thRow_7thColumn.getText();
		String expR4Gross = excelReader.getCellData(xlSheetName, 606, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 606, 7, actR4Gross);
		
		String actR4RMA = select4thRow_11thColumn.getText();
		String expR4RMA =excelReader.getCellData(xlSheetName, 607, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 607, 7, actR4RMA);
		
		
		System.out.println("**********************checkStockDetailsByBinsReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs       : "+actOpenWindowsCount +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("From Warehouse  : "+actFromWarehouse	+"  Value Expected  "+expFromWarehouse);
		
		System.out.println("*******************************ROW1**********************************");
		
		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		
		
		System.out.println("*******************************ROW2**********************************");
		
		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Bins            : "+actR2Batch				+"  Value Expected  "+expR2Batch);
		
		System.out.println("*******************************ROW3**********************************");
		
		System.out.println("Warehouse       : "+actR3Warehouse			+"  Value Expected  "+expR3Warehouse);
		System.out.println("Item            : "+actR3Item				+"  Value Expected  "+expR3Item);
		System.out.println("Units           : "+actR3Units				+"  Value Expected  "+expR3Units);
		System.out.println("Quantity        : "+actR3Quantity			+"  Value Expected  "+expR3Quantity);
		System.out.println("Rate            : "+actR3Rate				+"  Value Expected  "+expR3Rate);
		System.out.println("Gross           : "+actR3Gross				+"  Value Expected  "+expR3Gross);
		System.out.println("Bins            : "+actR3Bins				+"  Value Expected  "+expR3Bins);
		System.out.println("Bins2           : "+actR3Bins2				+"  Value Expected  "+expR3Bins2);
		
		System.out.println("*******************************ROW4**********************************");
		
		System.out.println("Warehouse       : "+actR4Warehouse			+"  Value Expected  "+expR4Warehouse);
		System.out.println("Item            : "+actR4Item				+"  Value Expected  "+expR4Item);
		System.out.println("Units           : "+actR4Units				+"  Value Expected  "+expR4Units);
		System.out.println("Quantity        : "+actR4Quantity			+"  Value Expected  "+expR4Quantity);
		System.out.println("Rate            : "+actR4Rate				+"  Value Expected  "+expR4Rate);
		System.out.println("Gross           : "+actR4Gross				+"  Value Expected  "+expR4Gross);
		System.out.println("RMA             : "+actR4RMA				+"  Value Expected  "+expR4RMA);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
			   && actDocNo.equalsIgnoreCase(expDocNo) && actFromWarehouse.equalsIgnoreCase(expFromWarehouse)
			   
			   
			   && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) && actR1Item.equalsIgnoreCase(expR1Item) && actR1Units.equalsIgnoreCase(expR1Units) 
			   && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) && actR1Gross.equalsIgnoreCase(expR1Gross) 
			   
			   && actR2Warehouse.equalsIgnoreCase(expR2Warehouse) && actR2Item.equalsIgnoreCase(expR2Item) && actR2Units.equalsIgnoreCase(expR2Units) 
			   && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Batch.equalsIgnoreCase(expR2Batch) 
			   
			   && actR3Warehouse.equalsIgnoreCase(expR3Warehouse) && actR3Item.equalsIgnoreCase(expR3Item) && actR3Units.equalsIgnoreCase(expR3Units) 
			   && actR3Quantity.equalsIgnoreCase(expR3Quantity) && actR3Rate.equalsIgnoreCase(expR3Rate) && actR3Gross.equalsIgnoreCase(expR3Gross) 
			   && actR3Bins.equalsIgnoreCase(expR3Bins) && actR3Bins2.equalsIgnoreCase(expR3Bins2)
			   
			   && actR4Warehouse.equalsIgnoreCase(expR4Warehouse) && actR4Item.equalsIgnoreCase(expR4Item) && actR4Units.equalsIgnoreCase(expR4Units) 
			   && actR4Quantity.equalsIgnoreCase(expR4Quantity) && actR4Rate.equalsIgnoreCase(expR4Rate) && actR4Gross.equalsIgnoreCase(expR4Gross) 
			   && actR4RMA.equalsIgnoreCase(expR4RMA)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 574, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 574, 8, resFail);

			return false;
		}
	}
	
	
	@FindBy(xpath="//span[contains(text(),'Stock Balances by Bins')]")
	private static WebElement stockBalancesByBins;
	
	public boolean checkStockBalanaceByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LogoutandLoginwithSU();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);
		
		getAction().sendKeys(Keys.END).build().perform();
		//Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binsReportMenu));
		binsReportMenu.click();
		
		//Thread.sleep(2000);
		//getAction().moveToElement(stockBalancesByBins).build().perform();
		//Thread.sleep(2000);
				
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockBalancesByBins));
		stockBalancesByBins.click();
		
		Thread.sleep(2000);	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
		sl_SelectAllItemsChkBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 619, 6);
         excelReader.setCellData(xlfile, xlSheetName, 619, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =  excelReader.getCellData(xlSheetName, 620, 6);
        excelReader.setCellData(xlfile, xlSheetName, 620, 7, actRow2List);
	

		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =  excelReader.getCellData(xlSheetName, 621, 6);
        excelReader.setCellData(xlfile, xlSheetName, 621, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List =  excelReader.getCellData(xlSheetName, 622, 6);
        excelReader.setCellData(xlfile, xlSheetName, 622, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =  excelReader.getCellData(xlSheetName, 623, 6);
        excelReader.setCellData(xlfile, xlSheetName, 623, 7, actRow5List);
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List =  excelReader.getCellData(xlSheetName, 624, 6);
        excelReader.setCellData(xlfile, xlSheetName, 624, 7, actRow6List);
		
		
		
		int reportsRow7ListCount = reportsRow7List.size();
		ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow7ListCount;i++)
		{
			String data = reportsRow7List.get(i).getText();
			reportsRow7ListArray.add(data);
		}
		String actRow7List = reportsRow7ListArray.toString();
		String expRow7List =  excelReader.getCellData(xlSheetName, 625, 6);
        excelReader.setCellData(xlfile, xlSheetName, 625, 7, actRow7List);
		
		
		
		System.out.println("***************************checkStockBalanaceByBinsReports*********************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		System.out.println(actRow7List);
		System.out.println(expRow7List);
		
		Thread.sleep(1000);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List)
				&& actRow7List.equalsIgnoreCase(expRow7List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 619, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 619, 8, resFail);
			return false;
		}
	}
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_588_0_DefaultFilter_0']")
	private static WebElement sbb_DefaultFilterTxt;
	
	@FindBy (xpath="//*[@id='FilterFields_588_0']/ul[23]/li[1]/div/label/span")
	private static WebElement sbb_report_FilterBinNameChkbox;
	
	public boolean checkStockBalanaceByBinsReportFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
			report_FilterExpandBtn.click();
			Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		Thread.sleep(2000);
		
		((JavascriptExecutor) getDriver()).executeScript("arguments[0].scrollIntoView(true);", report_FilterBinExpansion);
		
		Thread.sleep(2000); 
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBinExpansion));
		report_FilterBinExpansion.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sbb_report_FilterBinNameChkbox));
		
		Thread.sleep(1000);
		
		if (binsReport_report_FilterBinNameChkbox.isSelected()==false) 
		{
			sbb_report_FilterBinNameChkbox.click();
		}

		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sbb_DefaultFilterTxt));
		sbb_DefaultFilterTxt.click();
		sbb_DefaultFilterTxt.sendKeys( excelReader.getCellData(xlSheetName, 634, 5));
		
		Thread.sleep(2000);
		
		sbb_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsByWarehouseRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
		}
		String actRow1List = reportsByWarehouseRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 634, 6);
        
		excelReader.setCellData(xlfile, xlSheetName, 634, 7, actRow1List);
		

		int reportsByWarehouseRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsByWarehouseRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsByWarehouseRow2ListArray.add(data);
		}
		String actRow2List = reportsByWarehouseRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 635, 6);
        
		excelReader.setCellData(xlfile, xlSheetName, 635, 7, actRow2List);
		
		
		
		int reportsByWarehouseRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsByWarehouseRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsByWarehouseRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsByWarehouseRow3ListArray.add(data);
		}
		String actRow3List = reportsByWarehouseRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 636, 6);
        
		excelReader.setCellData(xlfile, xlSheetName, 636, 7, actRow3List);
		
		
		System.out.println("*********************************checkStockBalanaceByBinsReportFilterOption*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sbb_report_FilterBinNameChkbox));
		
		if (binsReport_report_FilterBinNameChkbox.isSelected()==true) 
		{
			sbb_report_FilterBinNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
		
		//int actRowCount = reportsRowsCount.size();
		//int expRowCount = 7; 
		
		String actRowCount=String.valueOf(getDriver().getWindowHandles().size());
		
		String expRowCount = excelReader.getCellData(xlSheetName, 637, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 637, 7, actRowCount);

		
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List)
				&& actRowCount.equals(expRowCount))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 634, 8, resPass);

			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 634, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	public boolean checkStockBalanaceByBinsReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		int reportItemGridListCount=reportTableColumn1RowsList.size();
	    for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 644, 5)))
             {
        		 reportTableColumn1RowsList.get(i).click();
        		 Thread.sleep(1000);
        		 break;
             }
		     }
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 644, 6);
         excelReader.setCellData(xlfile, xlSheetName, 644, 7, actMonthList);

		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 645, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 645, 6);
         excelReader.setCellData(xlfile, xlSheetName, 645, 7, actDay);

		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 646, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 646, 6);
         excelReader.setCellData(xlfile, xlSheetName, 646, 7, actVoucherList);

		
		System.out.println(actVoucherList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 647, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 647, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 647, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 648, 6);
        excelReader.setCellData(xlfile, xlSheetName, 648, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferheaderWarehouse1Txt));
		String actFromWarehouse = stockTransferheaderWarehouse1Txt.getAttribute("value");
		String expFromWarehouse = excelReader.getCellData(xlSheetName, 649, 6);
        excelReader.setCellData(xlfile, xlSheetName, 649, 7, actFromWarehouse);
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 650, 6);
        excelReader.setCellData(xlfile, xlSheetName, 650, 7, actR1Warehouse);
		
		
		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 651, 6);
        excelReader.setCellData(xlfile, xlSheetName, 651, 7, actR1Item);
		
		
		String actR1Units = select1stRow_3rdColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 652, 6);
        excelReader.setCellData(xlfile, xlSheetName, 652, 7, actR1Units);
		
		
		String actR1Quantity = select1stRow_4thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 653, 6);
        excelReader.setCellData(xlfile, xlSheetName, 653, 7, actR1Quantity);
		
		
		String actR1Rate = select1stRow_6thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 654, 6);
        excelReader.setCellData(xlfile, xlSheetName, 654, 7, actR1Rate);
		
		
		String actR1Gross = select1stRow_7thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 655, 6);
        excelReader.setCellData(xlfile, xlSheetName, 655, 7, actR1Gross);
		
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 656, 6);
        excelReader.setCellData(xlfile, xlSheetName, 656, 7, actR2Warehouse);
		
		
		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 657, 6);
        excelReader.setCellData(xlfile, xlSheetName, 657, 7, actR2Item);
		
		
		String actR2Units = select2ndRow_3rdColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 658, 6);
        excelReader.setCellData(xlfile, xlSheetName, 658, 7, actR2Units);
		
		
		String actR2Quantity = select2ndRow_4thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 659, 6);
        excelReader.setCellData(xlfile, xlSheetName, 659, 7, actR2Quantity);
		
		
		String actR2Rate = select2ndRow_6thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 660, 6);
        excelReader.setCellData(xlfile, xlSheetName, 660, 7, actR2Rate);
		
		
		String actR2Gross = select2ndRow_7thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 661, 6);
        excelReader.setCellData(xlfile, xlSheetName, 661, 7, actR2Gross);
		
		
		String actR2Batch = select2ndRow_8thColumn.getText();
		String expR2Batch =excelReader.getCellData(xlSheetName, 662, 6);
        excelReader.setCellData(xlfile, xlSheetName, 662, 7, actR2Batch);
		

		

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		String actR3Warehouse = select3rdRow_1stColumn.getText();
		String expR3Warehouse = excelReader.getCellData(xlSheetName, 663, 6);
        excelReader.setCellData(xlfile, xlSheetName, 663, 7, actR3Warehouse);
		
		String actR3Item = select3rdRow_2ndColumn.getText();
		String expR3Item = excelReader.getCellData(xlSheetName, 664, 6);
        excelReader.setCellData(xlfile, xlSheetName, 664, 7, actR3Item);
		
		String actR3Units = select3rdRow_3rdColumn.getText();
		String expR3Units = excelReader.getCellData(xlSheetName, 665, 6);
        excelReader.setCellData(xlfile, xlSheetName, 665, 7, actR3Units);
		
		String actR3Quantity = select3rdRow_4thColumn.getText();
		String expR3Quantity = excelReader.getCellData(xlSheetName, 666, 6);
        excelReader.setCellData(xlfile, xlSheetName, 666, 7, actR3Quantity);
		
		String actR3Rate = select3rdRow_6thColumn.getText();
		String expR3Rate = excelReader.getCellData(xlSheetName, 667, 6);
        excelReader.setCellData(xlfile, xlSheetName, 667, 7, actR3Rate);
		
		String actR3Gross = select3rdRow_7thColumn.getText();
		String expR3Gross =excelReader.getCellData(xlSheetName, 668, 6);
        excelReader.setCellData(xlfile, xlSheetName, 668, 7, actR3Gross);
		
		String actR3Bins = select3rdRow_9thColumn.getText();
		String expR3Bins = excelReader.getCellData(xlSheetName, 669, 6);
        excelReader.setCellData(xlfile, xlSheetName, 669, 7, actR3Bins);
		
		String actR3Bins2 = select3rdRow_10thColumn.getText();
		String expR3Bins2 = excelReader.getCellData(xlSheetName, 670, 6);
        excelReader.setCellData(xlfile, xlSheetName, 670, 7, actR3Bins2);
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
		String actR4Warehouse = select4thRow_1stColumn.getText();
		String expR4Warehouse = excelReader.getCellData(xlSheetName, 671, 6);
        excelReader.setCellData(xlfile, xlSheetName, 671, 7, actR4Warehouse);
		
		
		String actR4Item = select4thRow_2ndColumn.getText();
		String expR4Item = excelReader.getCellData(xlSheetName, 672, 6);
        excelReader.setCellData(xlfile, xlSheetName, 672, 7, actR4Item);
		
		
		String actR4Units = select4thRow_3rdColumn.getText();
		String expR4Units = excelReader.getCellData(xlSheetName, 673, 6);
        excelReader.setCellData(xlfile, xlSheetName, 673, 7, actR4Units);
		
		
		String actR4Quantity = select4thRow_4thColumn.getText();
		String expR4Quantity = excelReader.getCellData(xlSheetName, 674, 6);
        excelReader.setCellData(xlfile, xlSheetName, 674, 7,actR4Quantity);
		
		
		String actR4Rate = select4thRow_6thColumn.getText();
		String expR4Rate = excelReader.getCellData(xlSheetName, 675, 6);
        excelReader.setCellData(xlfile, xlSheetName, 675, 7, actR4Rate);
		
		
		String actR4Gross = select4thRow_7thColumn.getText();
		String expR4Gross = excelReader.getCellData(xlSheetName, 676, 6);
        excelReader.setCellData(xlfile, xlSheetName, 676, 7, actR4Gross);
		
		
		String actR4RMA = select4thRow_11thColumn.getText();
		String expR4RMA = excelReader.getCellData(xlSheetName, 677, 6);
        excelReader.setCellData(xlfile, xlSheetName, 677, 7, actR4RMA);
		
		
		
		System.out.println("**********************checkStockDetailsByBinsReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs       : "+actOpenWindowsCount +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("From Warehouse  : "+actFromWarehouse	+"  Value Expected  "+expFromWarehouse);
		
		System.out.println("*******************************ROW1**********************************");
		
		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		
		
		System.out.println("*******************************ROW2**********************************");
		
		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Bins            : "+actR2Batch				+"  Value Expected  "+expR2Batch);
		
		System.out.println("*******************************ROW3**********************************");
		
		System.out.println("Warehouse       : "+actR3Warehouse			+"  Value Expected  "+expR3Warehouse);
		System.out.println("Item            : "+actR3Item				+"  Value Expected  "+expR3Item);
		System.out.println("Units           : "+actR3Units				+"  Value Expected  "+expR3Units);
		System.out.println("Quantity        : "+actR3Quantity			+"  Value Expected  "+expR3Quantity);
		System.out.println("Rate            : "+actR3Rate				+"  Value Expected  "+expR3Rate);
		System.out.println("Gross           : "+actR3Gross				+"  Value Expected  "+expR3Gross);
		System.out.println("Bins            : "+actR3Bins				+"  Value Expected  "+expR3Bins);
		System.out.println("Bins2           : "+actR3Bins2				+"  Value Expected  "+expR3Bins2);
		
		System.out.println("*******************************ROW4**********************************");
		
		System.out.println("Warehouse       : "+actR4Warehouse			+"  Value Expected  "+expR4Warehouse);
		System.out.println("Item            : "+actR4Item				+"  Value Expected  "+expR4Item);
		System.out.println("Units           : "+actR4Units				+"  Value Expected  "+expR4Units);
		System.out.println("Quantity        : "+actR4Quantity			+"  Value Expected  "+expR4Quantity);
		System.out.println("Rate            : "+actR4Rate				+"  Value Expected  "+expR4Rate);
		System.out.println("Gross           : "+actR4Gross				+"  Value Expected  "+expR4Gross);
		System.out.println("RMA             : "+actR4RMA				+"  Value Expected  "+expR4RMA);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) &&actOpenWindowsCount.equals(expOpenWindowsCount)
			   && actDocNo.equalsIgnoreCase(expDocNo) && actFromWarehouse.equalsIgnoreCase(expFromWarehouse)
			   
			   
			   && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) && actR1Item.equalsIgnoreCase(expR1Item) && actR1Units.equalsIgnoreCase(expR1Units) 
			   && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) && actR1Gross.equalsIgnoreCase(expR1Gross) 
			   
			   && actR2Warehouse.equalsIgnoreCase(expR2Warehouse) && actR2Item.equalsIgnoreCase(expR2Item) && actR2Units.equalsIgnoreCase(expR2Units) 
			   && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Batch.equalsIgnoreCase(expR2Batch) 
			   
			   && actR3Warehouse.equalsIgnoreCase(expR3Warehouse) && actR3Item.equalsIgnoreCase(expR3Item) && actR3Units.equalsIgnoreCase(expR3Units) 
			   && actR3Quantity.equalsIgnoreCase(expR3Quantity) && actR3Rate.equalsIgnoreCase(expR3Rate) && actR3Gross.equalsIgnoreCase(expR3Gross) 
			   && actR3Bins.equalsIgnoreCase(expR3Bins) && actR3Bins2.equalsIgnoreCase(expR3Bins2)
			   
			   && actR4Warehouse.equalsIgnoreCase(expR4Warehouse) && actR4Item.equalsIgnoreCase(expR4Item) && actR4Units.equalsIgnoreCase(expR4Units) 
			   && actR4Quantity.equalsIgnoreCase(expR4Quantity) && actR4Rate.equalsIgnoreCase(expR4Rate) && actR4Gross.equalsIgnoreCase(expR4Gross) 
			   && actR4RMA.equalsIgnoreCase(expR4RMA)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 644, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 644, 8, resFail);
			return false;
		}
	}
	
	
	@FindBy(xpath="//span[contains(text(),'Stock Balances by Item by Bins')]")
	
	private static WebElement stockBalancesByItemByBins;
	
	public boolean checkStockBalancesByItemByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);
		getAction().sendKeys(Keys.END).build().perform();
		//Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binsReportMenu));
		binsReportMenu.click();
		
		//Thread.sleep(2000);	
		//getAction().moveToElement(stockBalancesByItemByBins).build().perform();
		//Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockBalancesByItemByBins));
		stockBalancesByItemByBins.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 689, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 689, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 689, 7, actRow1List);

		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 690, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 690, 7, actRow2List);

	

		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 691, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 691, 7, actRow3List);

		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 692, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 692, 7, actRow4List);

		
		
		System.out.println("***************************checkStockBalanaceByBinsReports*********************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		Thread.sleep(1000);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List))			
		{
			excelReader.setCellData(xlfile, xlSheetName, 689, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 689, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	public boolean checkStockBalancesByItemByBinsReportPrintOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());


		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ReportPrintBtn));
		sl_ReportPrintBtn.click();
		getWaitForAlert();
		Thread.sleep(1000);
		getAlert().accept();
		
		Thread.sleep(3000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 699, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 699, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(0));
	 	
	 	Thread.sleep(1000);
	 	
	 	getDriver().switchTo().window(openTabs.get(1)).close();
	 	
	 	getDriver().switchTo().window(openTabs.get(0));
		
		System.out.println("***************************checkStockBalancesByItemByBinsReportPrintOption*********************************");
		
		System.out.println("openTabs"+openTabs);
		System.out.println("Open Tabs Count : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
		
		Thread.sleep(1000);
		
		if(actOpenWindowsCount.equals(expOpenWindowsCount))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 699, 8, resPass);

			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 699, 8, resFail);
			return false;
		}
	}

	
	
	
	
	
	@FindBy(xpath="//span[contains(text(),'Expired Stock by Bins')]")
	
	private static WebElement expiredStockByBins;
	
	
	public boolean checkExpiredStockByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);
		getAction().sendKeys(Keys.END).build().perform();

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(binsReportMenu));
		binsReportMenu.click();
		
		Thread.sleep(2000);
		//getAction().moveToElement(expiredStockByBins).build().perform();
		//Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(expiredStockByBins));
		expiredStockByBins.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 704, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);	
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		String  actReportsTable =String.valueOf(reportsTable.getText().isEmpty());
		String  expReportsTable = excelReader.getCellData(xlSheetName, 704, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 704, 7,  actReportsTable );



		System.out.println("report_Body : "+actReportsTable+" Value Expected : "+expReportsTable);
		
		if(actReportsTable.equals(expReportsTable))
		{
			return true;
		}
		else
		{
			return false;
		}
	}
	
	
	@FindBy (xpath="//span[@id='id_prog_spanPlus']")
	private static WebElement customizationFieldsProgrammableFieldExpansion;
	
	@FindBy (xpath="//li[@id='TColProg-2']/span")
	private static WebElement customizationFieldsProgrammableField;
	
	
	public boolean checkExpiredStockByBinsReportCustomizationOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_ReportCustomizeBtn));
		report_ReportCustomizeBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationFieldsProgrammableFieldExpansion));
		customizationFieldsProgrammableFieldExpansion.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customizationFieldsProgrammableField));
		getAction().doubleClick(customizationFieldsProgrammableField).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeSaveBtn));
		osr_customizeSaveBtn.click();
		
		Thread.sleep(1000);
		
		String expMessage =  excelReader.getCellData(xlSheetName, 709, 6);
		String actMessage = checkValidationMessage(expMessage);
        
      excelReader.setCellData(xlfile, xlSheetName, 709, 7, actMessage);
	    
	    
	    
	    Thread.sleep(4000);
	    
	    /*if (sl_OkBtn.isDisplayed()==true) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(4000);
		}
	    else
	    {
	    	if (osr_customizeCloseBtn.isDisplayed()==true) 
	    	{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeCloseBtn));
				osr_customizeCloseBtn.click();
				
				Thread.sleep(4000);
			}
	    }*/
		
	    int reportsHeaderListCount1 = reportsHeaderList.size();
		ArrayList<String> reportsHeaderListArray1 = new ArrayList<String>();
		for(int i=0;i<reportsHeaderListCount1;i++)
		{
			String data = reportsHeaderList.get(i).getText();
			reportsHeaderListArray1.add(data);
		}
		String actHeaderList1 = reportsHeaderListArray1.toString();
		String expHeaderList1=  excelReader.getCellData(xlSheetName, 710, 6);
        
      excelReader.setCellData(xlfile, xlSheetName, 710, 7, actHeaderList1);
	    
	    System.out.println(actHeaderList1);
	    System.out.println(expHeaderList1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_ReportCustomizeBtn));
		report_ReportCustomizeBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CustomizationDeleteLayoutBtn));
		report_CustomizationDeleteLayoutBtn.click();
	
		getWaitForAlert();
		getAlert().accept();
		
		String expDeleteMessage =  excelReader.getCellData(xlSheetName, 711, 6);
		 String actDeleteMessage = checkValidationMessage(expDeleteMessage);
          excelReader.setCellData(xlfile, xlSheetName, 711, 7, actDeleteMessage);
	    
	   
	    
	    Thread.sleep(4000);
	    
	    /*if (sl_OkBtn.isDisplayed()==true) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
			sl_OkBtn.click();
			
			Thread.sleep(4000);
		}
	    else
	    {
	    	if (osr_customizeCloseBtn.isDisplayed()==true) 
	    	{
				getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(osr_customizeCloseBtn));
				osr_customizeCloseBtn.click();
				
				Thread.sleep(4000);
			}
	    }*/
	    
	    int reportsHeaderListCount = reportsHeaderList.size();
		ArrayList<String> reportsHeaderListArray = new ArrayList<String>();
		for(int i=0;i<reportsHeaderListCount;i++)
		{
			String data = reportsHeaderList.get(i).getText();
			reportsHeaderListArray.add(data);
		}
		String actHeaderList = reportsHeaderListArray.toString();
		String expHeaderList =  excelReader.getCellData(xlSheetName, 712, 6);
        
      excelReader.setCellData(xlfile, xlSheetName, 712, 7, actHeaderList);
	    
	    System.out.println(actHeaderList);
	    System.out.println(expHeaderList);
	    
		if(actMessage.equalsIgnoreCase(expMessage) && actHeaderList1.equalsIgnoreCase(expHeaderList1) && actDeleteMessage.equalsIgnoreCase(expDeleteMessage)
				&& actHeaderList.equalsIgnoreCase(expHeaderList))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 709, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 709, 8, resFail);
			return false;
		}
	}
	

	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[15]")
	private static WebElement stockBalanceByWarehouse;

	public boolean checkStockBalanceByWarehouseReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LogoutandLoginwithSU();

		getDriver().navigate().refresh();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getAction().sendKeys(Keys.END).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockBalanceByWarehouse));
		stockBalanceByWarehouse.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
		sl_SelectAllItemsChkBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(4000);
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =    excelReader.getCellData(xlSheetName, 719, 6);
        excelReader.setCellData(xlfile, xlSheetName, 719, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =     excelReader.getCellData(xlSheetName, 720, 6);
        excelReader.setCellData(xlfile, xlSheetName, 720, 7, actRow2List);
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =   excelReader.getCellData(xlSheetName, 721, 6);
        excelReader.setCellData(xlfile, xlSheetName, 721, 7, actRow3List);
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			if (i!=2) 
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List =    excelReader.getCellData(xlSheetName, 722, 6);
        excelReader.setCellData(xlfile, xlSheetName, 722, 7, actRow4List);
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =    excelReader.getCellData(xlSheetName, 723, 6);
        excelReader.setCellData(xlfile, xlSheetName, 723, 7, actRow5List);
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List =    excelReader.getCellData(xlSheetName, 724, 6);
        excelReader.setCellData(xlfile, xlSheetName, 724, 7, actRow6List);
		
		
		System.out.println("************************************checkStockBalanceByWarehouseReports********************************************");
	
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
			&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 719, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 719, 8, resFail);
			return false;
		}
	}


	
	
	
	
	
	
	
	public boolean checkStockBalanceByWarehouseReportIncludeServiceTypeItemOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_IncludeServiceTypeItemChkBox));
		sl_IncludeServiceTypeItemChkBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 730, 6);
        excelReader.setCellData(xlfile, xlSheetName, 730, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 731, 6);
        excelReader.setCellData(xlfile, xlSheetName, 731, 7, actRow2List);
		
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 732, 6);
        excelReader.setCellData(xlfile, xlSheetName, 732, 7, actRow3List);
		
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			if (i!=2) 
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List =  excelReader.getCellData(xlSheetName, 733, 6);
        excelReader.setCellData(xlfile, xlSheetName, 733, 7, actRow4List);
		
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =  excelReader.getCellData(xlSheetName, 734, 6);
        excelReader.setCellData(xlfile, xlSheetName, 734, 7, actRow5List);
		
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List =  excelReader.getCellData(xlSheetName, 735, 6);
        excelReader.setCellData(xlfile, xlSheetName, 735, 7, actRow6List);
		
		
		
		System.out.println("************************************checkStockBalanceByWarehouseReportIncludeServiceTypeItemOption********************************************");
	
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
			&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 730, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 730, 8, resFail);
			return false;
		}
	}

	
	
	
	
	
	
	
	public boolean checkStockBalanceByWarehouseReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		int reportItemGridListCount=reportTableColumn1RowsList.size();
	   
		for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 740, 5)))
             {
                 int count =i+5;
                 
                 for (int j = i; j < count; j++) 
     		     {
                	 String data1=reportTableColumn1RowsList.get(j).getText();
		             
                	 if (data1.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 741, 5)))
                     {
                		 System.err.println("Item Name : "+data1);
                		 reportTableColumn1RowsList.get(j).click();
                		 Thread.sleep(1000);
                		 break;
                     }
     		     }
             }
	      }
	    
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 741, 6);
		excelReader.setCellData(xlfile, xlSheetName, 741, 7, actMonthList);
		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 742, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
	String actDay=actDayList.toString();
	String expDayList=  excelReader.getCellData(xlSheetName, 742, 6);
	excelReader.setCellData(xlfile, xlSheetName, 742, 7,actDay);

		System.out.println(actDayList);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 743, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
	
		String actVoucherList=actVouchersList.toString();
		String expVouchersList=excelReader.getCellData(xlSheetName, 743, 6);
        excelReader.setCellData(xlfile, xlSheetName, 743, 7, actVoucherList);
        
		System.out.println(actVouchersList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 744, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		

		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 744, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 744, 7, actOpenWindowsCount);
	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 745, 6);
        excelReader.setCellData(xlfile, xlSheetName, 745, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(vendorAccountTxt));
		String actVendor = vendorAccountTxt.getAttribute("value");
		String expVendor = excelReader.getCellData(xlSheetName, 746, 6);
        excelReader.setCellData(xlfile, xlSheetName, 746, 7, actVendor);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(voucherHeaderCurrency));
		String actCurrency = voucherHeaderCurrency.getAttribute("value");
		String expCurrency = excelReader.getCellData(xlSheetName, 747, 6);
        excelReader.setCellData(xlfile, xlSheetName, 747, 7, actCurrency);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(departmentTxt));
		String actDepartment = departmentTxt.getAttribute("value");
		String expDepartment = excelReader.getCellData(xlSheetName, 748, 6);
        excelReader.setCellData(xlfile, xlSheetName, 748, 7, actDepartment);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(placeOFSupplyTxt));
		String actPlaceOfSupply = placeOFSupplyTxt.getAttribute("value");
		String expPlaceOfSupply = excelReader.getCellData(xlSheetName, 749, 6);
        excelReader.setCellData(xlfile, xlSheetName, 749, 7, actPlaceOfSupply);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(jurisdictionTxt));
		String actJurisdiction = jurisdictionTxt.getAttribute("value");
		String expJurisdiction = excelReader.getCellData(xlSheetName, 750, 6);
        excelReader.setCellData(xlfile, xlSheetName, 750, 7, actJurisdiction);

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 751, 6);
        excelReader.setCellData(xlfile, xlSheetName, 751, 7, actR1Warehouse);
		
		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 752, 6);
        excelReader.setCellData(xlfile, xlSheetName, 752, 7, actR1Item);
		
		String actR1TaxCode = select1stRow_3rdColumn.getText();
		String expR1TaxCode = excelReader.getCellData(xlSheetName, 753, 6);
        excelReader.setCellData(xlfile, xlSheetName, 753, 7, actR1TaxCode);
		
		String actR1PurchaseAccount = select1stRow_4thColumn.getText();
		String expR1PurchaseAccount = excelReader.getCellData(xlSheetName, 754, 6);
        excelReader.setCellData(xlfile, xlSheetName, 754, 7, actR1PurchaseAccount);
		
		String actR1Units = select1stRow_5thColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 755, 6);
        excelReader.setCellData(xlfile, xlSheetName, 755, 7, actR1Units);
		
		String actR1Quantity = select1stRow_9thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 756, 6);
        excelReader.setCellData(xlfile, xlSheetName, 756, 7,actR1Quantity);
		
		String actR1Rate = select1stRow_11thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 757, 6);
        excelReader.setCellData(xlfile, xlSheetName, 757, 7, actR1Rate);
		
		String actR1Gross = select1stRow_12thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 758, 6);
        excelReader.setCellData(xlfile, xlSheetName, 758, 7, actR1Gross);
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 759, 6);
        excelReader.setCellData(xlfile, xlSheetName, 759, 7, actR2Warehouse);
		
		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 760, 6);
        excelReader.setCellData(xlfile, xlSheetName, 760, 7, actR2Item);
		
		String actR2TaxCode = select2ndRow_3rdColumn.getText();
		String expR2TaxCode = excelReader.getCellData(xlSheetName, 761, 6);
        excelReader.setCellData(xlfile, xlSheetName, 761, 7, actR2TaxCode);
		
		String actR2PurchaseAccount = select2ndRow_4thColumn.getText();
		String expR2PurchaseAccount = excelReader.getCellData(xlSheetName, 762, 6);
        excelReader.setCellData(xlfile, xlSheetName, 762, 7, actR2PurchaseAccount);
		
		String actR2Units = select2ndRow_5thColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 763, 6);
        excelReader.setCellData(xlfile, xlSheetName, 763, 7, actR2Units);
		
		String actR2Quantity = select2ndRow_9thColumn.getText();
		String expR2Quantity =excelReader.getCellData(xlSheetName, 764, 6);
        excelReader.setCellData(xlfile, xlSheetName, 764, 7, actR2Quantity);
		
		String actR2Rate = select2ndRow_11thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 765, 6);
        excelReader.setCellData(xlfile, xlSheetName, 765, 7, actR2Rate);
		
		String actR2Gross = select2ndRow_12thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 766, 6);
        excelReader.setCellData(xlfile, xlSheetName, 766, 7, actR2Gross);
	 	
		System.out.println("**********************checkStockValuationReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("Vendor          : "+actVendor			+"  Value Expected  "+expVendor);
		System.out.println("Currency        : "+actCurrency			+"  Value Expected  "+expCurrency);
		System.out.println("Department      : "+actDepartment		+"  Value Expected  "+expDepartment);
		System.out.println("PlaceOfSupply   : "+actPlaceOfSupply	+"  Value Expected  "+expPlaceOfSupply);
		System.out.println("Jurisdiction    : "+actJurisdiction		+"  Value Expected  "+expJurisdiction);
		
		System.out.println("*******************************ROW1**********************************");
		
		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("TaxCode         : "+actR1TaxCode			+"  Value Expected  "+expR1TaxCode);
		System.out.println("PurchaseAccount : "+actR1PurchaseAccount	+"  Value Expected  "+expR1PurchaseAccount);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		
		System.out.println("*******************************ROW2**********************************");
		
		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("TaxCode         : "+actR2TaxCode			+"  Value Expected  "+expR2TaxCode);
		System.out.println("PurchaseAccount : "+actR2PurchaseAccount	+"  Value Expected  "+expR2PurchaseAccount);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
			   && actDocNo.equalsIgnoreCase(expDocNo) && actVendor.equalsIgnoreCase(expVendor) && actCurrency.equalsIgnoreCase(expCurrency) 
			   && actDepartment.equalsIgnoreCase(expDepartment) && actPlaceOfSupply.equalsIgnoreCase(expPlaceOfSupply) 
			   && actJurisdiction.equalsIgnoreCase(expJurisdiction) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
			   && actR1Item.equalsIgnoreCase(expR1Item) && actR1TaxCode.equalsIgnoreCase(expR1TaxCode) && actR1PurchaseAccount.equalsIgnoreCase(expR1PurchaseAccount) 
			   && actR1Units.equalsIgnoreCase(expR1Units) && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) 
			   && actR1Gross.equalsIgnoreCase(expR1Gross) && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) 
			   && actR2Item.equalsIgnoreCase(expR2Item) && actR2TaxCode.equalsIgnoreCase(expR2TaxCode) && actR2PurchaseAccount.equalsIgnoreCase(expR2PurchaseAccount) 
			   && actR2Units.equalsIgnoreCase(expR2Units) && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) 
			   && actR2Gross.equalsIgnoreCase(expR2Gross)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 740, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 740, 8, resFail);
			return false;
		}
	}
	
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[16]")
	private static WebElement stockAgeingAnalysisMenu;
	
	
	@FindBy(xpath="//span[contains(text(),'Ageing Analysis')]")
	private static WebElement ageingAnalysis;
	
	
	public boolean checkStockAgeingAnalysisReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		//getDriver().navigate().refresh();
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);
		
		getAction().sendKeys(Keys.END).build().perform();
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
		stockAgeingAnalysisMenu.click();
		
		
		Thread.sleep(2000);
		getAction().sendKeys(Keys.END).build().perform();
		
	//((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);",ageingAnalysis );
		
	
		
		getAction().moveToElement(ageingAnalysis).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingAnalysis));
		ageingAnalysis.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
		sl_SelectAllItemsChkBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 774, 6);
        excelReader.setCellData(xlfile, xlSheetName, 774, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 775, 6);
        excelReader.setCellData(xlfile, xlSheetName, 775, 7, actRow2List);
		
	

		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 776, 6);
        excelReader.setCellData(xlfile, xlSheetName, 776, 7, actRow3List);
		
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List =excelReader.getCellData(xlSheetName, 777, 6);
        excelReader.setCellData(xlfile, xlSheetName, 777, 7, actRow4List);
		
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = excelReader.getCellData(xlSheetName, 778, 6);
        excelReader.setCellData(xlfile, xlSheetName, 778, 7, actRow5List);
		
		
		
		System.out.println("***************************checkStockAgeingAnalysisReports*********************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		Thread.sleep(1000);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))	
		{
			 excelReader.setCellData(xlfile, xlSheetName, 774, 8, resPass);
			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 774, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_582_0_DefaultFilter_0']")
	private static WebElement saa_DefaultFilterTxt;
	
	
	public boolean checkStockAgeingAnalysisReportFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterWarehouseExpandBtn));
		report_FilterWarehouseExpandBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sa_report_FilterWarehouseNameChkbox));
		
		if (report_FilterWarehouseNameChkbox.isSelected()==false) 
		{
			sa_report_FilterWarehouseNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saa_DefaultFilterTxt));
		saa_DefaultFilterTxt.click();
		saa_DefaultFilterTxt.sendKeys(excelReader.getCellData(xlSheetName, 785, 5));
		
		Thread.sleep(2000);
		
		saa_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 785, 6);
        excelReader.setCellData(xlfile, xlSheetName, 785, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =excelReader.getCellData(xlSheetName, 786, 6);
        excelReader.setCellData(xlfile, xlSheetName, 786, 7, actRow2List);
	

		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =excelReader.getCellData(xlSheetName, 787, 6);
        excelReader.setCellData(xlfile, xlSheetName, 787, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 788, 6);
        excelReader.setCellData(xlfile, xlSheetName, 788, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =excelReader.getCellData(xlSheetName, 789, 6);
        excelReader.setCellData(xlfile, xlSheetName, 789, 7, actRow4List);
		
		
		System.out.println("***************************checkStockAgeingAnalysisReportFilterOption*********************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sa_report_FilterWarehouseNameChkbox));
		
		if (report_FilterWarehouseNameChkbox.isSelected()==true) 
		{
			sa_report_FilterWarehouseNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))	
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 785, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 785, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='MasterGroup__101']")
	private static WebElement stockAgeingAnalysisWarehouseTxt;
	
	@FindBy (xpath="//select[@id='RITCombobox__1']")
	private static WebElement stockAgeingDropdown;
	
	@FindBy (xpath="(//*[@id='dvReportInputs']//label/span)[2]")
	private static WebElement ignoreInternalTransferForAgeingChkbox;
	
	@FindBy(xpath="//*[@id='dvReportInputs']/div[12]/label/span")
	private static WebElement saa_IncludeServiceTypeItemChkBox;
	
	
	
	public boolean checkStockAgeingAnalysisReportWithOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		/*
		getDriver().navigate().refresh();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);
		
		getAction().sendKeys(Keys.END).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
		stockAgeingAnalysisMenu.click();
		
		
		//getAction().sendKeys(Keys.END).build().perform();
		Thread.sleep(1000);
		
	//((JavascriptExecutor)getDriver()).executeScript("arguments[0].scrollIntoView(true);",ageingAnalysis );
		
		getAction().moveToElement(stockAgeingAnalysisMenu).build().perform();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingAnalysis));
		ageingAnalysis.click();
		
		Thread.sleep(2000);
		*/

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisWarehouseTxt));
		stockAgeingAnalysisWarehouseTxt.sendKeys(  excelReader.getCellData(xlSheetName, 795, 5));
		Thread.sleep(2000);
		stockAgeingAnalysisWarehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingDropdown));
		Select stockAgeing = new Select(stockAgeingDropdown);
		stockAgeing.selectByIndex(1);
		
		Thread.sleep(2000);
		
		getAction().moveToElement(ignoreInternalTransferForAgeingChkbox).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ignoreInternalTransferForAgeingChkbox));
		ignoreInternalTransferForAgeingChkbox.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTypeDropdown));
		Select stockType = new Select(stockTypeDropdown);
		stockType.selectByIndex(1);
		
		getAction().moveToElement(saa_IncludeServiceTypeItemChkBox).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(saa_IncludeServiceTypeItemChkBox));
		saa_IncludeServiceTypeItemChkBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =   excelReader.getCellData(xlSheetName, 795, 6);
		excelReader.setCellData(xlfile, xlSheetName, 795, 7, actRow1List);

		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =  excelReader.getCellData(xlSheetName, 796, 6);
		excelReader.setCellData(xlfile, xlSheetName, 796, 7, actRow2List);
	

		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =  excelReader.getCellData(xlSheetName, 797, 6);
		excelReader.setCellData(xlfile, xlSheetName, 797, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List =  excelReader.getCellData(xlSheetName, 798, 6);
		excelReader.setCellData(xlfile, xlSheetName, 798, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =  excelReader.getCellData(xlSheetName, 799, 6);
		excelReader.setCellData(xlfile, xlSheetName, 799, 7, actRow5List);
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List =  excelReader.getCellData(xlSheetName, 800, 6);
		excelReader.setCellData(xlfile, xlSheetName, 800, 7, actRow6List);
		
		
		System.out.println("***************************checkStockAgeingAnalysisReportWithOptions*********************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List))	
		{
			 excelReader.setCellData(xlfile, xlSheetName, 795, 8, resPass);
			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 795, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	public boolean checkStockAgeingAnalysisReportWithBackTrackOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		int reportItemGridListCount=reportTableColumn1RowsList.size();
	   
		for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 805, 5)))
             {
                 int count =i+5;
                 
                 for (int j = i; j < count; j++) 
     		     {
                	 String data1=reportTableColumn1RowsList.get(j).getText();
		             
                	 if (data1.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 806, 5)))
                     {
                		 System.err.println("Item Name : "+data1);
                		 reportTableColumn1RowsList.get(j).click();
                		 Thread.sleep(1000);
                		 break;
                     }
     		     }
             }
	      }
	    
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 806, 6);
		excelReader.setCellData(xlfile, xlSheetName, 806, 7, actMonthList);
		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 807, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 807, 6);
		excelReader.setCellData(xlfile, xlSheetName, 807, 7, actDay);
		
		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 808, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 808, 6);
		excelReader.setCellData(xlfile, xlSheetName, 808, 7, actVoucherList);
		
		
		System.out.println(actVoucherList);
		System.out.println(expVouchersList);

		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 809, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		

		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 809, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 809, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 810, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 810, 7, actDocNo);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomer=customerAccountTxt.getAttribute("value");
		String expCustomer=excelReader.getCellData(xlSheetName, 811, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 811, 7, actCustomer);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	String actWarehouse = select1stRow_1stColumn.getText();
	 	String expWarehouse = excelReader.getCellData(xlSheetName, 812, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 812, 7, actWarehouse);
			
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	 	String actItem = select1stRow_2ndColumn.getText();
	 	String expItem = excelReader.getCellData(xlSheetName, 813, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 813, 7, actItem);
			
	 	
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
	 	String actTaxCode = select1stRow_3rdColumn.getText();
	 	String expTaxCode = excelReader.getCellData(xlSheetName, 814, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 814, 7, actTaxCode);
			
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
	 	String actSalesAccount = select1stRow_4thColumn.getText();
	 	String expSalesAccount =excelReader.getCellData(xlSheetName, 815, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 815, 7, actSalesAccount);
			
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	 	String actUnits = select1stRow_5thColumn.getText();
	 	String expUnits = excelReader.getCellData(xlSheetName, 816, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 816, 7, actUnits);
			
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	 	String actQuantity = select1stRow_11thColumn.getText();
	 	String expQuantity = excelReader.getCellData(xlSheetName, 817, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 817, 7, actQuantity);
			
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
	 	String actLSalesOrder = select1stRow_12thColumn.getText();
	 	String expLSalesOrder = excelReader.getCellData(xlSheetName, 818, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 818, 7, actLSalesOrder);
			
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	 	String actRate = select1stRow_14thColumn.getText();
	 	String expRate = excelReader.getCellData(xlSheetName, 819, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 819, 7, actRate);
			
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
	 	String actGross = select1stRow_15thColumn.getText();
	 	String expGross = excelReader.getCellData(xlSheetName, 820, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 820, 7, actGross);
			
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
	 	String actRMA = select1stRow_21stColumn.getText();
	 	String expRMA = excelReader.getCellData(xlSheetName, 821, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 821, 7, actRMA);
			
	 	
		System.out.println("**********************checkStockAgeingAnalysisReportWithBackTrackOptions*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
		System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
		System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
		System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
		System.out.println("Tax Code         : "+actTaxCode       +"  Value Expected  "+expTaxCode);
		System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
		System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
		System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
		System.out.println("L Sales Order    : "+actLSalesOrder   +"  Value Expected  "+expLSalesOrder);
		System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
		System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
		System.out.println("RMA              : "+actRMA           +"  Value Expected  "+expRMA);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
				&& actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
				   && actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
				   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesOrder.equalsIgnoreCase(expLSalesOrder)
				   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) && actRMA.equalsIgnoreCase(expRMA)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			  excelReader.setCellData(xlfile, xlSheetName, 805, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 805, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	@FindBy(xpath="//span[contains(text(),'Ageing analysis by batch')]")
	
	private static WebElement ageingAnalysisByBatch;
	
	
	public boolean checkAgeingAnalysisByBatchReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 Thread.sleep(1000);
		 LogoutandLoginwithSU();
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
			inventoryMenu.click();
			
			Thread.sleep(2000);
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
			inventoryReportsMenu.click();
			Thread.sleep(2000);
			
			getAction().sendKeys(Keys.END).build().perform();
			
			
			
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
			stockAgeingAnalysisMenu.click();
						
			Thread.sleep(2000);
			getAction().sendKeys(Keys.END).build().perform();
		
		getAction().moveToElement(ageingAnalysisByBatch).build().perform();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingAnalysisByBatch));
		ageingAnalysisByBatch.click();
		
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 828, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 828, 6);
        excelReader.setCellData(xlfile, xlSheetName, 828, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount-1;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 829, 6);
        excelReader.setCellData(xlfile, xlSheetName, 829, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount-1;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 830, 6);
        excelReader.setCellData(xlfile, xlSheetName, 830, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount-1;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 831, 6);
        excelReader.setCellData(xlfile, xlSheetName, 831, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount-1;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = excelReader.getCellData(xlSheetName, 832, 6);
        excelReader.setCellData(xlfile, xlSheetName, 832, 7, actRow5List);
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount-1;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List = excelReader.getCellData(xlSheetName, 833, 6);
        excelReader.setCellData(xlfile, xlSheetName, 833, 7, actRow6List);
		
		
		
		/*int reportsRow7ListCount = reportsRow7List.size();
		ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow7ListCount;i++)
		{
			String data = reportsRow7List.get(i).getText();
			reportsRow7ListArray.add(data);
		}
		String actRow7List = reportsRow7ListArray.toString();
		String expRow7List = "[Total, 62.00, 42.75, 38.00, 32.75, , , , , , , , , , , 24.00, 15.00, , , , ]";*/
		
		
		System.out.println("*****************************************checkAgeingAnalysisByBatchReports***************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		/*System.out.println(actRow7List);
		System.out.println(expRow7List);*/
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
				/*&& actRow7List.equalsIgnoreCase(expRow7List)*/)
		{
			excelReader.setCellData(xlfile, xlSheetName, 828, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 828, 8, resFail);

			return false;
		}
	}
	
	
	
	
	
	
	@FindBy (xpath="//ul[@id='FilterFields_622_0']/li[5]/a/i")
	private static WebElement stockAnalysisByBatch_Filter_Customize_WarehouseExpandBtn;
	
	@FindBy (xpath="//*[@id='FilterFields_622_0']/ul[5]/li[1]/div/label/input")
	private static WebElement stockAnalysisByBatch_Filter_Customize_WarehouseNameChkBox;
	
	@FindBy (xpath="//input[@id='FOption_622_0_DefaultFilter_0']")
	private static WebElement stockAnalysisByBatch_DefaultFilterTxt;
	
	
	public boolean checkStockAnalysisByBatchReportPrintOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ReportPrintBtn));
		sl_ReportPrintBtn.click();
		
		getWaitForAlert();
		Thread.sleep(1000);
		getAlert().accept();
		
		
		Thread.sleep(3000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 839, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 839, 7, actOpenWindowsCount);

		 
	 	getDriver().switchTo().window(openTabs.get(0));
	 	
	 	Thread.sleep(1000);
	 	
	 	getDriver().switchTo().window(openTabs.get(1)).close();
	 	
	 	getDriver().switchTo().window(openTabs.get(0));
		
		System.out.println("***************************checkStockAnalysisByBatchReportPrintOption*********************************");
		
		System.out.println("openTabs"+openTabs);
		System.out.println("Open Tabs Count : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
		
		Thread.sleep(1000);
		
		if(actOpenWindowsCount.equals(expOpenWindowsCount))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			return true;
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	public boolean checkAgeingAnalysisByBatchReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_WarehouseTxt));
		sl_WarehouseTxt.click();
		sl_WarehouseTxt.sendKeys( excelReader.getCellData(xlSheetName, 844, 5));
		Thread.sleep(2000);
		sl_WarehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ignoreInternalTransferForAgeingChkbox));
		ignoreInternalTransferForAgeingChkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 844, 6);
        excelReader.setCellData(xlfile, xlSheetName, 844, 7, actRow1List);

		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount-1;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 845, 6);
        excelReader.setCellData(xlfile, xlSheetName, 845, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount-1;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 846, 6);
        excelReader.setCellData(xlfile, xlSheetName, 847, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount-1;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 847, 6);
        excelReader.setCellData(xlfile, xlSheetName, 847, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount-1;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = excelReader.getCellData(xlSheetName, 848, 6);
        excelReader.setCellData(xlfile, xlSheetName, 848, 7, actRow5List);
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount-1;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List = excelReader.getCellData(xlSheetName, 849, 6);
        excelReader.setCellData(xlfile, xlSheetName, 849, 7, actRow6List);
		
		
		
		/*int reportsRow7ListCount = reportsRow7List.size();
		ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow7ListCount;i++)
		{
			String data = reportsRow7List.get(i).getText();
			reportsRow7ListArray.add(data);
		}
		String actRow7List = reportsRow7ListArray.toString();
		String expRow7List = "[Total, 62.00, 42.75, 38.00, 32.75, , , , , , , , , , , 24.00, 15.00, , , , ]";*/
		
		
		System.out.println("*****************************************checkAgeingAnalysisByBatchReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions***************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		/*System.out.println(actRow7List);
		System.out.println(expRow7List);*/
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
				/*&& actRow7List.equalsIgnoreCase(expRow7List)*/)
		{
			 excelReader.setCellData(xlfile, xlSheetName, 844, 8, resPass);
			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 844, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	public boolean checkAgeingAnalysisByBatchReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());

		int reportItemGridListCount=reportTableColumn1RowsList.size();
	   
		for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 855, 5)))
             {
                 int count =i+5;
                 
                 for (int j = i; j < count; j++) 
     		     {
                	 String data1=reportTableColumn1RowsList.get(j).getText();
		             
                	 if (data1.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 856, 5)))
                     {
                		 System.err.println("Item Name : "+data1);
                		 reportTableColumn1RowsList.get(j).click();
                		 Thread.sleep(1000);
                		 break;
                     }
     		     }
             }
	      }
	    
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 856, 6);
		excelReader.setCellData(xlfile, xlSheetName, 856, 7, actMonthList);
		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 857, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 857, 6);
		excelReader.setCellData(xlfile, xlSheetName, 857, 7, actDay);
		
		System.out.println(actDayList);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 858, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 858, 6);
		excelReader.setCellData(xlfile, xlSheetName, 858, 7, actVoucherList);
		
		System.out.println(actVouchersList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 859, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 859, 6);
		excelReader.setCellData(xlfile, xlSheetName, 859, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 860, 6);
		excelReader.setCellData(xlfile, xlSheetName, 860, 7, actDocNo);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomer=customerAccountTxt.getAttribute("value");
		String expCustomer=excelReader.getCellData(xlSheetName, 861, 6);
		excelReader.setCellData(xlfile, xlSheetName, 861, 7, actCustomer);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	String actWarehouse = select1stRow_1stColumn.getText();
	 	String expWarehouse =excelReader.getCellData(xlSheetName, 862, 6);
		excelReader.setCellData(xlfile, xlSheetName, 862, 7, actWarehouse);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	 	String actItem = select1stRow_2ndColumn.getText();
	 	String expItem = excelReader.getCellData(xlSheetName, 863, 6);
		excelReader.setCellData(xlfile, xlSheetName, 863, 7, actItem);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
	 	String actTaxCode = select1stRow_3rdColumn.getText();
	 	String expTaxCode = excelReader.getCellData(xlSheetName, 864, 6);
		excelReader.setCellData(xlfile, xlSheetName, 864, 7, actTaxCode);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
	 	String actSalesAccount = select1stRow_4thColumn.getText();
	 	String expSalesAccount =excelReader.getCellData(xlSheetName, 865, 6);
		excelReader.setCellData(xlfile, xlSheetName, 865, 7, actSalesAccount);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	 	String actUnits = select1stRow_5thColumn.getText();
	 	String expUnits = excelReader.getCellData(xlSheetName, 866, 6);
		excelReader.setCellData(xlfile, xlSheetName, 866, 7, actUnits);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	 	String actQuantity = select1stRow_11thColumn.getText();
	 	String expQuantity = excelReader.getCellData(xlSheetName, 867, 6);
		excelReader.setCellData(xlfile, xlSheetName, 867, 7, actQuantity);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
	 	String actLSalesOrder = select1stRow_12thColumn.getText();
	 	String expLSalesOrder = excelReader.getCellData(xlSheetName, 868, 6);
		excelReader.setCellData(xlfile, xlSheetName, 868, 7, actLSalesOrder);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	 	String actRate = select1stRow_14thColumn.getText();
	 	String expRate = excelReader.getCellData(xlSheetName, 869, 6);
		excelReader.setCellData(xlfile, xlSheetName, 869, 7, actRate);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
	 	String actGross = select1stRow_15thColumn.getText();
	 	String expGross = excelReader.getCellData(xlSheetName, 870, 6);
		excelReader.setCellData(xlfile, xlSheetName, 870, 7, actGross);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
	 	String actBatch = select1stRow_19thColumn.getText();
	 	String expBatch = excelReader.getCellData(xlSheetName, 871, 6);
		excelReader.setCellData(xlfile, xlSheetName, 871, 7, actBatch);
	 	
		System.out.println("**********************checkStockAgeingAnalysisReportWithBackTrackOptions*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
		System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
		System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
		System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
		System.out.println("Tax Code         : "+actTaxCode       +"  Value Expected  "+expTaxCode);
		System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
		System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
		System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
		System.out.println("L Sales Order    : "+actLSalesOrder   +"  Value Expected  "+expLSalesOrder);
		System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
		System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
		System.out.println("Batch            : "+actBatch         +"  Value Expected  "+expBatch);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) &&actOpenWindowsCount.equals(expOpenWindowsCount)
				&& actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
				   && actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
				   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesOrder.equalsIgnoreCase(expLSalesOrder)
				   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 855, 8, resPass);

			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 855, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
/*	@FindBy (xpath="(//div[@id='dvReportDetails']/div/table/tbody)[1]/tr/td")
	private static List<WebElement> reportTableCells;
	*/
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[16]/ul/li[4]")
	private static WebElement ageingAnalysisByRMA;
	
	
	
	public boolean checkAgeingAnalysisByRMAReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);

		getAction().sendKeys(Keys.END).build().perform();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
		stockAgeingAnalysisMenu.click();
		
		
		Thread.sleep(2000);
		getAction().moveToElement(ageingAnalysisByRMA).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingAnalysisByRMA));
		ageingAnalysisByRMA.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 880, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportPage1Count = reportTableCells.size();
		ArrayList<String> reportPage1ListArray = new ArrayList<String>();
		for(int i=0;i<reportPage1Count;i++)
		{
			String data = reportTableCells.get(i).getText();
			reportPage1ListArray.add(data);
		}
		String actPage1List = reportPage1ListArray.toString();
		String expPage1List = excelReader.getCellData(xlSheetName, 880, 6);
        excelReader.setCellData(xlfile, xlSheetName, 880, 7, actPage1List);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
		report_NextBtn.click();
		
		Thread.sleep(2000);
		
		if (report_PageNoBtn.getText().equalsIgnoreCase("3")) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_PreviousBtn));
			report_PreviousBtn.click();
			Thread.sleep(2000);
		}
		
		int reportPage2Count = reportTableCells.size();
		ArrayList<String> reportPage2ListArray = new ArrayList<String>();
		for(int i=0;i<reportPage2Count;i++)
		{
			String data = reportTableCells.get(i).getText();
			reportPage2ListArray.add(data);
		}
		String actPage2List = reportPage2ListArray.toString();
		String expPage2List = excelReader.getCellData(xlSheetName, 881, 6);
        excelReader.setCellData(xlfile, xlSheetName, 881, 7, actPage2List);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
		report_NextBtn.click();
		
		Thread.sleep(2000);
		
		int reportPage3Count = reportTableCells.size();
		ArrayList<String> reportPage3ListArray = new ArrayList<String>();
		for(int i=0;i<reportPage3Count;i++)
		{
			String data = reportTableCells.get(i).getText();
			reportPage3ListArray.add(data);
		}
		String actPage3List = reportPage3ListArray.toString();
		String expPage3List = excelReader.getCellData(xlSheetName, 882, 6);
        excelReader.setCellData(xlfile, xlSheetName, 882, 7, actPage3List);
		
		
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_NextBtn));
		report_NextBtn.click();
		
		Thread.sleep(2000);
		
		int reportPage4Count = reportTableCells.size();
		ArrayList<String> reportPage4ListArray = new ArrayList<String>();
		for(int i=0;i<reportPage4Count;i++)
		{
			String data = reportTableCells.get(i).getText();
			reportPage4ListArray.add(data);
		}
		String actPage4List = reportPage4ListArray.toString();
		String expPage4List = "[43, RMA#R2PVVAT8, 1.00, 7.41, , , , , , , , , , , , , , , 1.00, 7.41, 44, RMA#R2PVVAT9, 1.00, 7.41, 1.00, 7.41, , , , , , , , , , , , , , , 45, Total, 43.00, 318.66, 32.00, 237.14, , , , , , , , , , , 11.00, 81.52, , ]";*/
		
		System.out.println("*****************************************checkAgeingAnalysisByRMAReports***************************************");
		
		System.out.println(actPage1List);
		System.out.println(expPage1List);
		
		System.out.println(actPage2List);
		System.out.println(expPage2List);
		
		System.out.println(actPage3List);
		System.out.println(expPage3List);
		
		/*System.out.println(actPage4List);
		System.out.println(expPage4List);*/
		
		if(actPage1List.equalsIgnoreCase(expPage1List) && actPage2List.equalsIgnoreCase(expPage2List) && actPage3List.equalsIgnoreCase(expPage3List) 
			/*&& actPage4List.equalsIgnoreCase(expPage4List)*/)
		{
			 excelReader.setCellData(xlfile, xlSheetName, 880, 8, resPass);
			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 880, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_623_0_DefaultFilter_0']")
	private static WebElement stockAnalysisByRMA_DefaultFilterTxt;
	
	@FindBy (xpath="//*[@id='FilterFields_623_0']/ul[5]/li[1]/div/label/span")
	private static WebElement ageingAnalysis_report_FilterWarehouseNameChkbox;
	
	
	public boolean checkAgeingAnalysisByRMAReportWithFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		  excelReader=new ExcelReader(POJOUtility.getExcelPath());
		  
		  getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
			report_FilterExpandBtn.click();
			Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterWarehouseExpandBtn));
		report_FilterWarehouseExpandBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingAnalysis_report_FilterWarehouseNameChkbox));
		
		if (report_FilterWarehouseNameChkbox.isSelected()==false) 
		{
			ageingAnalysis_report_FilterWarehouseNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAnalysisByRMA_DefaultFilterTxt));
		stockAnalysisByRMA_DefaultFilterTxt.click();
		stockAnalysisByRMA_DefaultFilterTxt.sendKeys( excelReader.getCellData(xlSheetName, 888, 5));
		
		Thread.sleep(2000);
		
		stockAnalysisByRMA_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
	
		
		String actReportTable=String.valueOf(reportsTable.getText().isEmpty());
		String expReportTable = excelReader.getCellData(xlSheetName, 888, 6);
		excelReader.setCellData(xlfile, xlSheetName, 888, 7, actReportTable);

		
		System.out.println("****************************checkStockAnalysisByRMAReportWithFilterOption******************************");
		
		System.out.println(actReportTable);
		System.out.println(expReportTable);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingAnalysis_report_FilterWarehouseNameChkbox));
		
		if (report_FilterWarehouseNameChkbox.isSelected()==true) 
		{
			ageingAnalysis_report_FilterWarehouseNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
			
		Thread.sleep(2000);
		
		
		
		String actRowCount=String.valueOf(reportsRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 889, 6);
		excelReader.setCellData(xlfile, xlSheetName, 889, 7, actRowCount);

		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		if(actReportTable.equals(expReportTable) && actRowCount.equals(expRowCount))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 888, 8, resPass);
			
			return true;
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			
			Thread.sleep(2000);
			excelReader.setCellData(xlfile, xlSheetName, 888, 8, resFail);
			
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public boolean checkAgeingAnalysisByRMAReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_WarehouseTxt));
		sl_WarehouseTxt.click();
		sl_WarehouseTxt.sendKeys(
                excelReader.getCellData(xlSheetName, 895, 5));
		Thread.sleep(2000);
		sl_WarehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ignoreInternalTransferForAgeingChkbox));
		ignoreInternalTransferForAgeingChkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		
		
		String actReportTable=String.valueOf(reportsTable.getText().isEmpty());
		String expReportTable = excelReader.getCellData(xlSheetName, 895, 6);
		excelReader.setCellData(xlfile, xlSheetName, 895, 7, actReportTable);

		
		System.out.println("*****************************************checkAgeingAnalysisByRMAReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions***************************************");
		
		System.out.println(actReportTable);
		System.out.println(expReportTable);
		
		if(actReportTable.equals(expReportTable))
		{
			excelReader.setCellData(xlfile, xlSheetName, 895, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 895, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public boolean checkAgeingAnalysisByRMAReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
		report_CloseBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_WarehouseTxt));
		sl_WarehouseTxt.click();
		sl_WarehouseTxt.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		sl_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 900, 5));
		Thread.sleep(2000);
		sl_WarehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		int reportItemGridListCount=reportTableColumn1RowsList.size();
		   
		for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 901, 5)))
             {
                 int count =i+7;
                 
                 for (int j = i; j < count; j++) 
     		     {
                	 String data1=reportTableColumn1RowsList.get(j).getText();
		             
                	 if (data1.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 902, 5)))
                     {
                		 System.err.println("Item Name : "+data1);
                		 reportTableColumn1RowsList.get(j).click();
                		 Thread.sleep(1000);
                		 break;
                     }
     		     }
             }
	      }
	    
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 902, 6);
		excelReader.setCellData(xlfile, xlSheetName, 902, 7, actMonthList);
		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 903, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 903, 6);
		excelReader.setCellData(xlfile, xlSheetName, 903, 7, actDay);
		
		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 904, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 904, 6);
		excelReader.setCellData(xlfile, xlSheetName, 904, 7, actVoucherList);
		
		System.out.println(actVouchersList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 905, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 905, 6);
		excelReader.setCellData(xlfile, xlSheetName, 905, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 906, 6);
		excelReader.setCellData(xlfile, xlSheetName, 906, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferheaderWarehouse1Txt));
		String actFromWarehouse = stockTransferheaderWarehouse1Txt.getAttribute("value");
		String expFromWarehouse = excelReader.getCellData(xlSheetName, 907, 6);
		excelReader.setCellData(xlfile, xlSheetName, 907, 7, actFromWarehouse);
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 908, 6);
		excelReader.setCellData(xlfile, xlSheetName, 908, 7, actR1Warehouse);
		
		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 909, 6);
		excelReader.setCellData(xlfile, xlSheetName, 909, 7, actR1Item);
		
		String actR1Units = select1stRow_3rdColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 910, 6);
		excelReader.setCellData(xlfile, xlSheetName, 910, 7, actR1Units);
		
		String actR1Quantity = select1stRow_4thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 911, 6);
		excelReader.setCellData(xlfile, xlSheetName, 911, 7, actR1Quantity);
		
		String actR1Rate = select1stRow_6thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 912, 6);
		excelReader.setCellData(xlfile, xlSheetName, 912, 7, actR1Rate);
		
		String actR1Gross = select1stRow_7thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 913, 6);
		excelReader.setCellData(xlfile, xlSheetName, 913, 7, actR1Gross);
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 914, 6);
		excelReader.setCellData(xlfile, xlSheetName, 914, 7, actR2Warehouse);
		
		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 915, 6);
		excelReader.setCellData(xlfile, xlSheetName, 915, 7, actR2Item);
		
		String actR2Units = select2ndRow_3rdColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 916, 6);
		excelReader.setCellData(xlfile, xlSheetName, 916, 7, actR2Units);
		
		String actR2Quantity = select2ndRow_4thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 917, 6);
		excelReader.setCellData(xlfile, xlSheetName, 917, 7, actR2Quantity);
		
		String actR2Rate = select2ndRow_6thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 918, 6);
		excelReader.setCellData(xlfile, xlSheetName, 918, 7, actR2Rate);
		
		String actR2Gross = select2ndRow_7thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 919, 6);
		excelReader.setCellData(xlfile, xlSheetName, 919, 7, actR2Gross);
		
		String actR2Batch = select2ndRow_8thColumn.getText();
		String expR2Batch = excelReader.getCellData(xlSheetName, 920, 6);
		excelReader.setCellData(xlfile, xlSheetName, 920, 7, actR2Batch);

		

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		String actR3Warehouse = select3rdRow_1stColumn.getText();
		String expR3Warehouse = excelReader.getCellData(xlSheetName, 921, 6);
		excelReader.setCellData(xlfile, xlSheetName, 921, 7, actR3Warehouse);

		
		String actR3Item = select3rdRow_2ndColumn.getText();
		String expR3Item = excelReader.getCellData(xlSheetName, 922, 6);
		excelReader.setCellData(xlfile, xlSheetName, 922, 7, actR3Item);

		
		String actR3Units = select3rdRow_3rdColumn.getText();
		String expR3Units = excelReader.getCellData(xlSheetName, 923, 6);
		excelReader.setCellData(xlfile, xlSheetName, 923, 7, actR3Units);

		
		String actR3Quantity = select3rdRow_4thColumn.getText();
		String expR3Quantity = excelReader.getCellData(xlSheetName, 924, 6);
		excelReader.setCellData(xlfile, xlSheetName, 924, 7, actR3Quantity);

		
		String actR3Rate = select3rdRow_6thColumn.getText();
		String expR3Rate = excelReader.getCellData(xlSheetName, 925, 6);
		excelReader.setCellData(xlfile, xlSheetName, 925, 7, actR3Rate);

		
		String actR3Gross = select3rdRow_7thColumn.getText();
		String expR3Gross = excelReader.getCellData(xlSheetName, 926, 6);
		excelReader.setCellData(xlfile, xlSheetName, 926, 7, actR3Gross);

		
		String actR3Bins = select3rdRow_9thColumn.getText();
		String expR3Bins = excelReader.getCellData(xlSheetName, 927, 6);
		excelReader.setCellData(xlfile, xlSheetName, 927, 7, actR3Bins);

		
		String actR3Bins2 = select3rdRow_10thColumn.getText();
		String expR3Bins2 = excelReader.getCellData(xlSheetName, 928, 6);
		excelReader.setCellData(xlfile, xlSheetName, 928, 7, actR3Bins2);

		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
		String actR4Warehouse = select4thRow_1stColumn.getText();
		String expR4Warehouse = excelReader.getCellData(xlSheetName, 929, 6);
		excelReader.setCellData(xlfile, xlSheetName, 929, 7, actR4Warehouse);
		
		String actR4Item = select4thRow_2ndColumn.getText();
		String expR4Item =  excelReader.getCellData(xlSheetName, 930, 6);
		excelReader.setCellData(xlfile, xlSheetName, 930, 7, actR4Item);
		
		String actR4Units = select4thRow_3rdColumn.getText();
		String expR4Units =  excelReader.getCellData(xlSheetName, 931, 6);
		excelReader.setCellData(xlfile, xlSheetName, 931, 7, actR4Units);
		
		String actR4Quantity = select4thRow_4thColumn.getText();
		String expR4Quantity =  excelReader.getCellData(xlSheetName, 932, 6);
		excelReader.setCellData(xlfile, xlSheetName, 932, 7, actR4Quantity);
		
		String actR4Rate = select4thRow_6thColumn.getText();
		String expR4Rate =  excelReader.getCellData(xlSheetName, 933, 6);
		excelReader.setCellData(xlfile, xlSheetName, 933, 7, actR4Rate);
		
		String actR4Gross = select4thRow_7thColumn.getText();
		String expR4Gross =  excelReader.getCellData(xlSheetName, 934, 6);
		excelReader.setCellData(xlfile, xlSheetName, 934, 7, actR4Gross);
		
		String actR4RMA = select4thRow_11thColumn.getText();
		String expR4RMA =  excelReader.getCellData(xlSheetName, 935, 6);
		excelReader.setCellData(xlfile, xlSheetName, 935, 7, actR4RMA);
		
		
		System.out.println("**********************checkAgeingAnalysisByRMAReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs       : "+actOpenWindowsCount +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("From Warehouse  : "+actFromWarehouse	+"  Value Expected  "+expFromWarehouse);
		
		System.out.println("*******************************ROW1**********************************");
		
		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		
		
		System.out.println("*******************************ROW2**********************************");
		
		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Bins            : "+actR2Batch				+"  Value Expected  "+expR2Batch);
		
		System.out.println("*******************************ROW3**********************************");
		
		System.out.println("Warehouse       : "+actR3Warehouse			+"  Value Expected  "+expR3Warehouse);
		System.out.println("Item            : "+actR3Item				+"  Value Expected  "+expR3Item);
		System.out.println("Units           : "+actR3Units				+"  Value Expected  "+expR3Units);
		System.out.println("Quantity        : "+actR3Quantity			+"  Value Expected  "+expR3Quantity);
		System.out.println("Rate            : "+actR3Rate				+"  Value Expected  "+expR3Rate);
		System.out.println("Gross           : "+actR3Gross				+"  Value Expected  "+expR3Gross);
		System.out.println("Bins            : "+actR3Bins				+"  Value Expected  "+expR3Bins);
		System.out.println("Bins2           : "+actR3Bins2				+"  Value Expected  "+expR3Bins2);
		
		System.out.println("*******************************ROW4**********************************");
		
		System.out.println("Warehouse       : "+actR4Warehouse			+"  Value Expected  "+expR4Warehouse);
		System.out.println("Item            : "+actR4Item				+"  Value Expected  "+expR4Item);
		System.out.println("Units           : "+actR4Units				+"  Value Expected  "+expR4Units);
		System.out.println("Quantity        : "+actR4Quantity			+"  Value Expected  "+expR4Quantity);
		System.out.println("Rate            : "+actR4Rate				+"  Value Expected  "+expR4Rate);
		System.out.println("Gross           : "+actR4Gross				+"  Value Expected  "+expR4Gross);
		System.out.println("RMA             : "+actR4RMA				+"  Value Expected  "+expR4RMA);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
			   && actDocNo.equalsIgnoreCase(expDocNo) && actFromWarehouse.equalsIgnoreCase(expFromWarehouse)
			   
			   
			   && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) && actR1Item.equalsIgnoreCase(expR1Item) && actR1Units.equalsIgnoreCase(expR1Units) 
			   && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) && actR1Gross.equalsIgnoreCase(expR1Gross) 
			   
			   && actR2Warehouse.equalsIgnoreCase(expR2Warehouse) && actR2Item.equalsIgnoreCase(expR2Item) && actR2Units.equalsIgnoreCase(expR2Units) 
			   && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Batch.equalsIgnoreCase(expR2Batch) 
			   
			   && actR3Warehouse.equalsIgnoreCase(expR3Warehouse) && actR3Item.equalsIgnoreCase(expR3Item) && actR3Units.equalsIgnoreCase(expR3Units) 
			   && actR3Quantity.equalsIgnoreCase(expR3Quantity) && actR3Rate.equalsIgnoreCase(expR3Rate) && actR3Gross.equalsIgnoreCase(expR3Gross) 
			   && actR3Bins.equalsIgnoreCase(expR3Bins) && actR3Bins2.equalsIgnoreCase(expR3Bins2)
			   
			   && actR4Warehouse.equalsIgnoreCase(expR4Warehouse) && actR4Item.equalsIgnoreCase(expR4Item) && actR4Units.equalsIgnoreCase(expR4Units) 
			   && actR4Quantity.equalsIgnoreCase(expR4Quantity) && actR4Rate.equalsIgnoreCase(expR4Rate) && actR4Gross.equalsIgnoreCase(expR4Gross) 
			   && actR4RMA.equalsIgnoreCase(expR4RMA)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 900, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 900, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	@FindBy (xpath="//tbody[@id='LandingGridBody']/tr/td[10]")
	private static List<WebElement> ageingAnalysisByBinHomeTableItemNamesList;
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[16]/ul/li[5]")
	private static WebElement ageingStockByBins;
	
	public boolean checkAgeingStockByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);

		getAction().sendKeys(Keys.END).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
		stockAgeingAnalysisMenu.click();

		Thread.sleep(2000);
		
		getAction().moveToElement(ageingStockByBins).build().perform();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingStockByBins));
		ageingStockByBins.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = ageingAnalysisByBinHomeTableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 940, 5) )
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 941, 5)) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 942, 5)) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 943, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 943, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 943, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 944, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 944, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 945, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 945, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 946, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 946, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = excelReader.getCellData(xlSheetName, 947, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 947, 7, actRow5List);
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List = excelReader.getCellData(xlSheetName, 948, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 948, 7, actRow6List);
		
		
		
		int reportsRow7ListCount = reportsRow7List.size();
		ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow7ListCount;i++)
		{
			String data = reportsRow7List.get(i).getText();
			reportsRow7ListArray.add(data);
		}
		String actRow7List = reportsRow7ListArray.toString();
		String expRow7List = excelReader.getCellData(xlSheetName, 949, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 949, 7, actRow7List);
		
		
		System.out.println("*****************************************checkAgeingStockByBinsReports***************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		System.out.println(actRow7List);
		System.out.println(expRow7List);
		
			
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
				&& actRow7List.equalsIgnoreCase(expRow7List))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 940, 8, resPass);

			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 940, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_591_0_DefaultFilter_0']")
	private static WebElement stockAnalysisByBin_DefaultFilterTxt;
	
	@FindBy(xpath="//*[@id='FilterFields_591_0']/ul[5]/li[1]/div/label/span")
	private static WebElement sabins_report_FilterWarehouseNameChkbox;
	
	public boolean checkStockAnalysisByBinReportWithFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());

		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
			report_FilterExpandBtn.click();
			Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterWarehouseExpandBtn));
		report_FilterWarehouseExpandBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sabins_report_FilterWarehouseNameChkbox));
		
		if (report_FilterWarehouseNameChkbox.isSelected()==false) 
		{
			sabins_report_FilterWarehouseNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAnalysisByBin_DefaultFilterTxt));
		stockAnalysisByBin_DefaultFilterTxt.click();
		stockAnalysisByBin_DefaultFilterTxt.sendKeys(excelReader.getCellData(xlSheetName, 955, 5));
		
		Thread.sleep(2000);
		
		stockAnalysisByBin_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(4000);
	
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = "[Bin Update Bin Update]";
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = "[FIFO COGS ITEM, , , , , , , , , , , , , , , , , , ]";
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = "[Grand Total, , , , , , , , , , , , , , , , , , ]";*/
		
		
		
		String actReportsTable=String.valueOf(reportsTable.getText().isEmpty());
		String expReportsTable = excelReader.getCellData(xlSheetName, 955, 6);
		excelReader.setCellData(xlfile, xlSheetName, 955, 7, actReportsTable);
		
		System.out.println("****************************checkStockAnalysisByBinReportWithFilterOption******************************");
		
		System.out.println("Reports Table isEmpty : "+actReportsTable+"  Value Expected  "+expReportsTable);
		
		/*System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);*/
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sabins_report_FilterWarehouseNameChkbox));
		
		if (report_FilterWarehouseNameChkbox.isSelected()==true) 
		{
			sabins_report_FilterWarehouseNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
			
		Thread.sleep(2000);
		
		
		String actRowCount=String.valueOf(reportsRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 956, 6);
		excelReader.setCellData(xlfile, xlSheetName, 956, 7, actRowCount);
		
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		if(actReportsTable.equals(expReportsTable)	&& actRowCount.equals(expRowCount))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 955, 8, resPass);
			return true;
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 955, 8, resFail);

			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	public boolean checkAgeingAnalysisByBinReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_WarehouseTxt));
		sl_WarehouseTxt.click();
		sl_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 960, 5));
		Thread.sleep(2000);
		sl_WarehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ignoreInternalTransferForAgeingChkbox));
		ignoreInternalTransferForAgeingChkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 960, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 960, 7, actRow1List);

		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 961, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 961, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 962, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 962, 7, actRow3List);
		
		
		System.out.println("*****************************************checkAgeingAnalysisByBinReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions***************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 960, 8, resPass);

			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 960, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	public boolean checkAgeingAnalysisByBinReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
		report_CloseBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_WarehouseTxt));
		sl_WarehouseTxt.click();
		sl_WarehouseTxt.sendKeys(Keys.END, Keys.SHIFT, Keys.HOME);
		sl_WarehouseTxt.sendKeys( excelReader.getCellData(xlSheetName, 967, 5));
		Thread.sleep(2000);
		sl_WarehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		int reportItemGridListCount=reportTableColumn1RowsList.size();
		   
		for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 968, 5)))
             {
                 int count =i+2;
                 
                 for (int j = i; j < count; j++) 
     		     {
                	 String data1=reportTableColumn1RowsList.get(j).getText();
		             
                	 if (data1.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 969, 5)))
                     {
                		 System.err.println("Item Name : "+data1);
                		 reportTableColumn1RowsList.get(j).click();
                		 Thread.sleep(1000);
                		 break;
                     }
     		     }
             }
	      }
	    
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 969, 6);
		excelReader.setCellData(xlfile, xlSheetName, 969, 7, actMonthList);

		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 970, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 970, 6);
		excelReader.setCellData(xlfile, xlSheetName, 970, 7, actDay);
		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 971, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 971, 6);
		excelReader.setCellData(xlfile, xlSheetName, 971, 7, actVoucherList);


		
		System.out.println(actVoucherList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 972, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 972, 6);
		excelReader.setCellData(xlfile, xlSheetName, 972, 7, actOpenWindowsCount);


	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 973, 6);
		excelReader.setCellData(xlfile, xlSheetName, 973, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferheaderWarehouse1Txt));
		String actFromWarehouse = stockTransferheaderWarehouse1Txt.getAttribute("value");
		String expFromWarehouse = excelReader.getCellData(xlSheetName, 974, 6);
		excelReader.setCellData(xlfile, xlSheetName, 974, 7, actFromWarehouse);
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 975, 6);
		excelReader.setCellData(xlfile, xlSheetName, 975, 7, actR1Warehouse);
		
		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 976, 6);
		excelReader.setCellData(xlfile, xlSheetName, 976, 7, actR1Item);
		
		String actR1Units = select1stRow_3rdColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 977, 6);
		excelReader.setCellData(xlfile, xlSheetName, 977, 7, actR1Units);
		
		String actR1Quantity = select1stRow_4thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 978, 6);
		excelReader.setCellData(xlfile, xlSheetName, 978, 7, actR1Quantity);
		
		String actR1Rate = select1stRow_6thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 979, 6);
		excelReader.setCellData(xlfile, xlSheetName, 979, 7, actR1Rate);
		
		String actR1Gross = select1stRow_7thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 980, 6);
		excelReader.setCellData(xlfile, xlSheetName, 980, 7, actR1Gross);
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 981, 6);
		excelReader.setCellData(xlfile, xlSheetName, 981, 7, actR2Warehouse);
		
		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 982, 6);
		excelReader.setCellData(xlfile, xlSheetName, 982, 7, actR2Item);
		
		String actR2Units = select2ndRow_3rdColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 983, 6);
		excelReader.setCellData(xlfile, xlSheetName, 983, 7, actR2Units);
		
		String actR2Quantity = select2ndRow_4thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 984, 6);
		excelReader.setCellData(xlfile, xlSheetName, 984, 7, actR2Quantity);
		
		String actR2Rate = select2ndRow_6thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 985, 6);
		excelReader.setCellData(xlfile, xlSheetName, 985, 7, actR2Rate);
		
		String actR2Gross = select2ndRow_7thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 986, 6);
		excelReader.setCellData(xlfile, xlSheetName, 986, 7, actR2Gross);
		
		String actR2Batch = select2ndRow_8thColumn.getText();
		String expR2Batch = excelReader.getCellData(xlSheetName, 987, 6);
		excelReader.setCellData(xlfile, xlSheetName, 987, 7, actR2Batch);

		

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		String actR3Warehouse = select3rdRow_1stColumn.getText();
		String expR3Warehouse = excelReader.getCellData(xlSheetName, 988, 6);
		excelReader.setCellData(xlfile, xlSheetName, 988, 7, actR3Warehouse);

		
		String actR3Item = select3rdRow_2ndColumn.getText();
		String expR3Item = excelReader.getCellData(xlSheetName, 989, 6);
		excelReader.setCellData(xlfile, xlSheetName, 989, 7, actR3Item);

		
		String actR3Units = select3rdRow_3rdColumn.getText();
		String expR3Units = excelReader.getCellData(xlSheetName, 990, 6);
		excelReader.setCellData(xlfile, xlSheetName, 990, 7, actR3Units);

		
		String actR3Quantity = select3rdRow_4thColumn.getText();
		String expR3Quantity = excelReader.getCellData(xlSheetName, 991, 6);
		excelReader.setCellData(xlfile, xlSheetName, 991, 7, actR3Quantity);

		
		String actR3Rate = select3rdRow_6thColumn.getText();
		String expR3Rate = excelReader.getCellData(xlSheetName, 992, 6);
		excelReader.setCellData(xlfile, xlSheetName, 992, 7, actR3Rate);

		
		String actR3Gross = select3rdRow_7thColumn.getText();
		String expR3Gross = excelReader.getCellData(xlSheetName, 993, 6);
		excelReader.setCellData(xlfile, xlSheetName, 993, 7, actR3Gross);

		
		String actR3Bins = select3rdRow_9thColumn.getText();
		String expR3Bins = excelReader.getCellData(xlSheetName, 994, 6);
		excelReader.setCellData(xlfile, xlSheetName, 994, 7, actR3Bins);

		
		String actR3Bins2 = select3rdRow_10thColumn.getText();
		String expR3Bins2 = excelReader.getCellData(xlSheetName, 995, 6);
		excelReader.setCellData(xlfile, xlSheetName, 995, 7, actR3Bins2);

		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
		String actR4Warehouse = select4thRow_1stColumn.getText();
		String expR4Warehouse = excelReader.getCellData(xlSheetName, 996, 6);
		excelReader.setCellData(xlfile, xlSheetName, 996, 7, actR4Warehouse);
		
		String actR4Item = select4thRow_2ndColumn.getText();
		String expR4Item =  excelReader.getCellData(xlSheetName, 997, 6);
		excelReader.setCellData(xlfile, xlSheetName, 997, 7, actR4Item);
		
		String actR4Units = select4thRow_3rdColumn.getText();
		String expR4Units =  excelReader.getCellData(xlSheetName, 998, 6);
		excelReader.setCellData(xlfile, xlSheetName, 998, 7, actR4Units);
		
		String actR4Quantity = select4thRow_4thColumn.getText();
		String expR4Quantity =  excelReader.getCellData(xlSheetName, 999, 6);
		excelReader.setCellData(xlfile, xlSheetName, 999, 7, actR4Quantity);
		
		String actR4Rate = select4thRow_6thColumn.getText();
		String expR4Rate =  excelReader.getCellData(xlSheetName, 1000, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1000, 7, actR4Rate);
		
		String actR4Gross = select4thRow_7thColumn.getText();
		String expR4Gross =  excelReader.getCellData(xlSheetName, 1001, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1001, 7, actR4Gross);
		
		String actR4RMA = select4thRow_11thColumn.getText();
		String expR4RMA =  excelReader.getCellData(xlSheetName, 1002, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1002, 7, actR4RMA);
		
		System.out.println("**********************checkAgeingAnalysisByBinReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs       : "+actOpenWindowsCount +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("From Warehouse  : "+actFromWarehouse	+"  Value Expected  "+expFromWarehouse);
		
		System.out.println("*******************************ROW1**********************************");
		
		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		
		
		System.out.println("*******************************ROW2**********************************");
		
		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Bins            : "+actR2Batch				+"  Value Expected  "+expR2Batch);
		
		System.out.println("*******************************ROW3**********************************");
		
		System.out.println("Warehouse       : "+actR3Warehouse			+"  Value Expected  "+expR3Warehouse);
		System.out.println("Item            : "+actR3Item				+"  Value Expected  "+expR3Item);
		System.out.println("Units           : "+actR3Units				+"  Value Expected  "+expR3Units);
		System.out.println("Quantity        : "+actR3Quantity			+"  Value Expected  "+expR3Quantity);
		System.out.println("Rate            : "+actR3Rate				+"  Value Expected  "+expR3Rate);
		System.out.println("Gross           : "+actR3Gross				+"  Value Expected  "+expR3Gross);
		System.out.println("Bins            : "+actR3Bins				+"  Value Expected  "+expR3Bins);
		System.out.println("Bins2           : "+actR3Bins2				+"  Value Expected  "+expR3Bins2);
		
		System.out.println("*******************************ROW4**********************************");
		
		System.out.println("Warehouse       : "+actR4Warehouse			+"  Value Expected  "+expR4Warehouse);
		System.out.println("Item            : "+actR4Item				+"  Value Expected  "+expR4Item);
		System.out.println("Units           : "+actR4Units				+"  Value Expected  "+expR4Units);
		System.out.println("Quantity        : "+actR4Quantity			+"  Value Expected  "+expR4Quantity);
		System.out.println("Rate            : "+actR4Rate				+"  Value Expected  "+expR4Rate);
		System.out.println("Gross           : "+actR4Gross				+"  Value Expected  "+expR4Gross);
		System.out.println("RMA             : "+actR4RMA				+"  Value Expected  "+expR4RMA);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
			   && actDocNo.equalsIgnoreCase(expDocNo) && actFromWarehouse.equalsIgnoreCase(expFromWarehouse)
			   
			   
			   && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) && actR1Item.equalsIgnoreCase(expR1Item) && actR1Units.equalsIgnoreCase(expR1Units) 
			   && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) && actR1Gross.equalsIgnoreCase(expR1Gross) 
			   
			   && actR2Warehouse.equalsIgnoreCase(expR2Warehouse) && actR2Item.equalsIgnoreCase(expR2Item) && actR2Units.equalsIgnoreCase(expR2Units) 
			   && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Batch.equalsIgnoreCase(expR2Batch) 
			   
			   && actR3Warehouse.equalsIgnoreCase(expR3Warehouse) && actR3Item.equalsIgnoreCase(expR3Item) && actR3Units.equalsIgnoreCase(expR3Units) 
			   && actR3Quantity.equalsIgnoreCase(expR3Quantity) && actR3Rate.equalsIgnoreCase(expR3Rate) && actR3Gross.equalsIgnoreCase(expR3Gross) 
			   && actR3Bins.equalsIgnoreCase(expR3Bins) && actR3Bins2.equalsIgnoreCase(expR3Bins2)
			   
			   && actR4Warehouse.equalsIgnoreCase(expR4Warehouse) && actR4Item.equalsIgnoreCase(expR4Item) && actR4Units.equalsIgnoreCase(expR4Units) 
			   && actR4Quantity.equalsIgnoreCase(expR4Quantity) && actR4Rate.equalsIgnoreCase(expR4Rate) && actR4Gross.equalsIgnoreCase(expR4Gross) 
			   && actR4RMA.equalsIgnoreCase(expR4RMA)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 967, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 967, 8, resFail);
			return false;
		}
	}
	
	
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[16]/ul/li[6]")
	private static WebElement ageingByItemByBins;
	
	
	
	
	public boolean checkAgeingByItemByBinsReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());

		 LogoutandLoginwithSU();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);

		getAction().sendKeys(Keys.END).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockAgeingAnalysisMenu));
		stockAgeingAnalysisMenu.click();
		
		Thread.sleep(2000);
		
		getAction().moveToElement(ageingByItemByBins).build().perform();
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ageingByItemByBins));
		ageingByItemByBins.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1009, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 1009, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 1009, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1010, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 1010, 7, actRow2List);
			
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 1011, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 1011, 7, actRow3List);
			
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 1012, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 1012, 7, actRow4List);
			
		
		
		
		System.out.println("*****************************************checkAgeingByItemByBinsReports***************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
			
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1009, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1009, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_646_0_DefaultFilter_0']")
	private static WebElement stockByItemByBin_DefaultFilterTxt;
	
	@FindBy(xpath="//*[@id='FilterFields_646_0']/ul[5]/li[1]/div/label/span")
	private static WebElement stockByItem_report_FilterWarehouseNameChkbox;
	
	public boolean checkStockByItemByBinReportWithFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
			report_FilterExpandBtn.click();
			Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterWarehouseExpandBtn));
		report_FilterWarehouseExpandBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockByItem_report_FilterWarehouseNameChkbox));
		
		if (report_FilterWarehouseNameChkbox.isSelected()==false) 
		{
			stockByItem_report_FilterWarehouseNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockByItemByBin_DefaultFilterTxt));
		stockByItemByBin_DefaultFilterTxt.click();
		stockByItemByBin_DefaultFilterTxt.sendKeys( excelReader.getCellData(xlSheetName, 1019, 5));
		
		Thread.sleep(2000);
		
		stockByItemByBin_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(4000);
	
		/*getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = "[FIFO COGS ITEM FIFO COGS ITEM]";
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = "[Bin Update, , , , , , , , , , , , , , , , , , ]";
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = "[Total, , , , , , , , , , , , , , , , , , ]";*/
		
	
		
		String actReportsTable=String.valueOf(reportsTable.getText().isEmpty());
		String expReportsTable = excelReader.getCellData(xlSheetName, 1019, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1019, 7, actReportsTable);
		
		System.out.println("****************************checkStockByItemByBinReportWithFilterOption******************************");
		
		System.out.println("Reports Table isEmpty : "+actReportsTable+"  Value Expected  "+expReportsTable);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockByItem_report_FilterWarehouseNameChkbox));
		
		if (report_FilterWarehouseNameChkbox.isSelected()==true) 
		{
			stockByItem_report_FilterWarehouseNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
			
		Thread.sleep(2000);
		
		
		
		String actRowCount=String.valueOf(reportsRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 1020, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1020, 7, actRowCount);
		
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		if(actReportsTable.equals(expReportsTable) && actRowCount.equals(expRowCount))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			  excelReader.setCellData(xlfile, xlSheetName, 1019, 8, resPass);

			return true;
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1019, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	public boolean checkAgeingByItemByBinReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		  excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_WarehouseTxt));
		sl_WarehouseTxt.click();
		sl_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 1024, 5));
		Thread.sleep(2000);
		sl_WarehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ignoreInternalTransferForAgeingChkbox));
		ignoreInternalTransferForAgeingChkbox.click();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 1024, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1024, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1025, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1025, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 1026, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1026, 7, actRow3List);
		
		
		
		System.out.println("*****************************************checkAgeingByItemByBinReportWithWarehouseAndIgnoreInternalTransferForAgeingOptions***************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1024, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1024, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	public boolean checkAgeingByItemByBinReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
		report_CloseBtn.click();
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_WarehouseTxt));
		sl_WarehouseTxt.click();
		sl_WarehouseTxt.sendKeys(Keys.END,Keys.SHIFT,Keys.HOME);
		sl_WarehouseTxt.sendKeys(excelReader.getCellData(xlSheetName, 1030, 5));
		Thread.sleep(2000);
		sl_WarehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
		
		int reportItemGridListCount=reportTableColumn1RowsList.size();
		   
		for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1031, 5)))
             {
                 int count =i+2;
                 
                 for (int j = i; j < count; j++) 
     		     {
                	 String data1=reportTableColumn1RowsList.get(j).getText();
		             
                	 if (data1.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1032, 5)))
                     {
                		 System.err.println("Item Name : "+data1);
                		 reportTableColumn1RowsList.get(j).click();
                		 Thread.sleep(1000);
                		 break;
                     }
     		     }
             }
	      }
	    
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 1032, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1032, 7, actMonthList);
		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1033, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 1033, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1033, 7, actDay);
		
		System.out.println(actDayList);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1034, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 1034, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1034, 7, actVoucherList);
		
		
		System.out.println(actVoucherList);
		System.out.println(expVouchersList);

		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1035, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 1035, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1035, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 1036, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1036, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferheaderWarehouse1Txt));
		String actFromWarehouse = stockTransferheaderWarehouse1Txt.getAttribute("value");
		String expFromWarehouse = excelReader.getCellData(xlSheetName, 1037, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1037, 7, actFromWarehouse);
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 1038, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1038, 7, actR1Warehouse);
		
		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 1039, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1039, 7, actR1Item);
		
		String actR1Units = select1stRow_3rdColumn.getText();
		String expR1Units = excelReader.getCellData(xlSheetName, 1040, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1040, 7, actR1Units);
		
		String actR1Quantity = select1stRow_4thColumn.getText();
		String expR1Quantity = excelReader.getCellData(xlSheetName, 1041, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1041, 7, actR1Quantity);
		
		String actR1Rate = select1stRow_6thColumn.getText();
		String expR1Rate = excelReader.getCellData(xlSheetName, 1042, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1042, 7, actR1Rate);
		
		String actR1Gross = select1stRow_7thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 1043, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1043, 7, actR1Gross);
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 1044, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1044, 7, actR2Warehouse);
		
		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item = excelReader.getCellData(xlSheetName, 1045, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1045, 7, actR2Item);
		
		String actR2Units = select2ndRow_3rdColumn.getText();
		String expR2Units = excelReader.getCellData(xlSheetName, 1046, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1046, 7, actR2Units);
		
		String actR2Quantity = select2ndRow_4thColumn.getText();
		String expR2Quantity = excelReader.getCellData(xlSheetName, 1047, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1047, 7, actR2Quantity);
		
		String actR2Rate = select2ndRow_6thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 1048, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1048, 7, actR2Rate);
		
		String actR2Gross = select2ndRow_7thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 1049, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1049, 7, actR2Gross);
		
		String actR2Batch = select2ndRow_8thColumn.getText();
		String expR2Batch = excelReader.getCellData(xlSheetName, 1050, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1050, 7, actR2Batch);

		

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		String actR3Warehouse = select3rdRow_1stColumn.getText();
		String expR3Warehouse = excelReader.getCellData(xlSheetName, 1051, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1051, 7, actR3Warehouse);

		
		String actR3Item = select3rdRow_2ndColumn.getText();
		String expR3Item = excelReader.getCellData(xlSheetName, 1052, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1052, 7, actR3Item);

		
		String actR3Units = select3rdRow_3rdColumn.getText();
		String expR3Units = excelReader.getCellData(xlSheetName, 1053, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1053, 7, actR3Units);

		
		String actR3Quantity = select3rdRow_4thColumn.getText();
		String expR3Quantity = excelReader.getCellData(xlSheetName, 1054, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1054, 7, actR3Quantity);

		
		String actR3Rate = select3rdRow_6thColumn.getText();
		String expR3Rate = excelReader.getCellData(xlSheetName, 1055, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1055, 7, actR3Rate);

		
		String actR3Gross = select3rdRow_7thColumn.getText();
		String expR3Gross = excelReader.getCellData(xlSheetName, 1056, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1056, 7, actR3Gross);

		
		String actR3Bins = select3rdRow_9thColumn.getText();
		String expR3Bins = excelReader.getCellData(xlSheetName, 1057, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1057, 7, actR3Bins);

		
		String actR3Bins2 = select3rdRow_10thColumn.getText();
		String expR3Bins2 = excelReader.getCellData(xlSheetName, 1058, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1058, 7, actR3Bins2);

		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
		String actR4Warehouse = select4thRow_1stColumn.getText();
		String expR4Warehouse = excelReader.getCellData(xlSheetName, 1059, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1059, 7, actR4Warehouse);
		
		String actR4Item = select4thRow_2ndColumn.getText();
		String expR4Item =  excelReader.getCellData(xlSheetName, 1060, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1060, 7, actR4Item);
		
		String actR4Units = select4thRow_3rdColumn.getText();
		String expR4Units =  excelReader.getCellData(xlSheetName, 1061, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1061, 7, actR4Units);
		
		String actR4Quantity = select4thRow_4thColumn.getText();
		String expR4Quantity =  excelReader.getCellData(xlSheetName, 1062, 6);
		excelReader.setCellData(xlfile, xlSheetName,1062, 7, actR4Quantity);
		
		String actR4Rate = select4thRow_6thColumn.getText();
		String expR4Rate =  excelReader.getCellData(xlSheetName, 1063, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1063, 7, actR4Rate);
		
		String actR4Gross = select4thRow_7thColumn.getText();
		String expR4Gross =  excelReader.getCellData(xlSheetName, 1064, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1064, 7, actR4Gross);
		
		String actR4RMA = select4thRow_11thColumn.getText();
		String expR4RMA =  excelReader.getCellData(xlSheetName, 1065, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1065, 7, actR4RMA);
		
		System.out.println("**********************checkAgeingByItemByBinReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs       : "+actOpenWindowsCount +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("From Warehouse  : "+actFromWarehouse	+"  Value Expected  "+expFromWarehouse);
		
		System.out.println("*******************************ROW1**********************************");
		
		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		
		
		System.out.println("*******************************ROW2**********************************");
		
		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Bins            : "+actR2Batch				+"  Value Expected  "+expR2Batch);
		
		System.out.println("*******************************ROW3**********************************");
		
		System.out.println("Warehouse       : "+actR3Warehouse			+"  Value Expected  "+expR3Warehouse);
		System.out.println("Item            : "+actR3Item				+"  Value Expected  "+expR3Item);
		System.out.println("Units           : "+actR3Units				+"  Value Expected  "+expR3Units);
		System.out.println("Quantity        : "+actR3Quantity			+"  Value Expected  "+expR3Quantity);
		System.out.println("Rate            : "+actR3Rate				+"  Value Expected  "+expR3Rate);
		System.out.println("Gross           : "+actR3Gross				+"  Value Expected  "+expR3Gross);
		System.out.println("Bins            : "+actR3Bins				+"  Value Expected  "+expR3Bins);
		System.out.println("Bins2           : "+actR3Bins2				+"  Value Expected  "+expR3Bins2);
		
		System.out.println("*******************************ROW4**********************************");
		
		System.out.println("Warehouse       : "+actR4Warehouse			+"  Value Expected  "+expR4Warehouse);
		System.out.println("Item            : "+actR4Item				+"  Value Expected  "+expR4Item);
		System.out.println("Units           : "+actR4Units				+"  Value Expected  "+expR4Units);
		System.out.println("Quantity        : "+actR4Quantity			+"  Value Expected  "+expR4Quantity);
		System.out.println("Rate            : "+actR4Rate				+"  Value Expected  "+expR4Rate);
		System.out.println("Gross           : "+actR4Gross				+"  Value Expected  "+expR4Gross);
		System.out.println("RMA             : "+actR4RMA				+"  Value Expected  "+expR4RMA);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) &&actOpenWindowsCount.equals(expOpenWindowsCount)
			   && actDocNo.equalsIgnoreCase(expDocNo) && actFromWarehouse.equalsIgnoreCase(expFromWarehouse)
			   
			   
			   && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) && actR1Item.equalsIgnoreCase(expR1Item) && actR1Units.equalsIgnoreCase(expR1Units) 
			   && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) && actR1Gross.equalsIgnoreCase(expR1Gross) 
			   
			   && actR2Warehouse.equalsIgnoreCase(expR2Warehouse) && actR2Item.equalsIgnoreCase(expR2Item) && actR2Units.equalsIgnoreCase(expR2Units) 
			   && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Batch.equalsIgnoreCase(expR2Batch) 
			   
			   && actR3Warehouse.equalsIgnoreCase(expR3Warehouse) && actR3Item.equalsIgnoreCase(expR3Item) && actR3Units.equalsIgnoreCase(expR3Units) 
			   && actR3Quantity.equalsIgnoreCase(expR3Quantity) && actR3Rate.equalsIgnoreCase(expR3Rate) && actR3Gross.equalsIgnoreCase(expR3Gross) 
			   && actR3Bins.equalsIgnoreCase(expR3Bins) && actR3Bins2.equalsIgnoreCase(expR3Bins2)
			   
			   && actR4Warehouse.equalsIgnoreCase(expR4Warehouse) && actR4Item.equalsIgnoreCase(expR4Item) && actR4Units.equalsIgnoreCase(expR4Units) 
			   && actR4Quantity.equalsIgnoreCase(expR4Quantity) && actR4Rate.equalsIgnoreCase(expR4Rate) && actR4Gross.equalsIgnoreCase(expR4Gross) 
			   && actR4RMA.equalsIgnoreCase(expR4RMA)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 1030, 8, resPass);

			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 1030, 8, resFail);
			return false;
		}
	}
	
	

	
	
	
	
	

	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[17]")
	private static WebElement reorderReport;
	
	public boolean checkReorderReports() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		  excelReader=new ExcelReader(POJOUtility.getExcelPath());
		  LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);
		
		getAction().sendKeys(Keys.END).build().perform();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reorderReport));
		reorderReport.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
		sl_SelectAllItemsChkBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		
		
		String actReportsTable=String.valueOf(reportsTable.getText().isEmpty());
		String expReportsTable = excelReader.getCellData(xlSheetName, 1070, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1070, 7, actReportsTable);

		
		System.out.println("*********************************checkReorderReportReport*****************************************");
		
		System.out.println("Reports Table isEmpty : "+actReportsTable+"  Value Expected  "+expReportsTable);
		
		if(actReportsTable.equals(expReportsTable))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1070, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1070, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public boolean checkReorderReportPrintOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ReportPrintBtn));
		sl_ReportPrintBtn.click();
		
		getWaitForAlert();
		Thread.sleep(1000);
		getAlert().accept();
		
		
		String expMessage = excelReader.getCellData(xlSheetName, 1074, 6);
		String actMessage = checkValidationMessage(expMessage);
      excelReader.setCellData(xlfile, xlSheetName, 1074, 7, actMessage); 
		
		
		
		System.out.println("*********************************checkReorderReportPrintOption*****************************************");
		
		System.out.println("Error Message : "+actMessage+"  Value Expected  "+expMessage);
		
		if(actMessage.equalsIgnoreCase(expMessage))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1074, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1074, 8, resFail);
			return false;
		}
	}
	
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[18]")
	private static WebElement fastMovingItem;
	
	
	
	public boolean checkFastMovingItemsReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();

		getAction().sendKeys(Keys.END).build().perform();
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)", "");
		//getAction().moveToElement(fastMovingItem).build().perform();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fastMovingItem));
		fastMovingItem.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1079, 5)) || 
					actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1080, 5)) 
					|| actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1081, 5)) || 
					actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1082, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().equals(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 1082, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1082, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1083, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1083, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =excelReader.getCellData(xlSheetName, 1084, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1084, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 1085, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1085, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = excelReader.getCellData(xlSheetName, 1086, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1086, 7, actRow5List);
		
		
		System.out.println("*********************************checkFastMovingItemsReport*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1079, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1079, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_584_0_DefaultFilter_0']")
	private static WebElement fastMovingItems_DefaultFilterTxt;
	
	
	@FindBy(xpath="//*[@id='FilterFields_584_0']/ul[4]/li[1]/div/label/span")
	private static WebElement fastMoving_report_FilterItemNameChkbox;
	
	public boolean checkFastMovingItemsReportWithFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
			report_FilterExpandBtn.click();
			Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterItemExpansion));
		report_FilterItemExpansion.click();
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fastMoving_report_FilterItemNameChkbox));
		
		
		if (report_FilterItemNameChkbox.isSelected()==false) 
		{
			fastMoving_report_FilterItemNameChkbox.click();
		}
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fastMovingItems_DefaultFilterTxt));
		fastMovingItems_DefaultFilterTxt.click();
		fastMovingItems_DefaultFilterTxt.sendKeys(excelReader.getCellData(xlSheetName, 1091, 5));
		
		Thread.sleep(2000);
		
		fastMovingItems_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 1091, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1091, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =  excelReader.getCellData(xlSheetName, 1092, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1092, 7, actRow2List);
		
		System.out.println("****************************checkFastMovingItemsReportWithFilterOption******************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fastMoving_report_FilterItemNameChkbox));
		
		if (report_FilterItemNameChkbox.isSelected()==true) 
		{
			fastMoving_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
			
		Thread.sleep(4000);
		
		
		String actRowCount=String.valueOf(reportsRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 1093, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1093, 7, actRowCount);

		
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		Thread.sleep(2000);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRowCount.equals(expRowCount))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1091, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 1091, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	public boolean checkFastMovingItemsReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		int reportItemGridListCount=reportTableColumn1RowsList.size();
	   
		for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1097, 5)))
             {
         		 System.err.println("Item Name : "+data);
         		 reportTableColumn1RowsList.get(i).click();
         		 Thread.sleep(1000);
         		 break;
              }
  		}
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 1097, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1097, 7, actMonthList);

		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1098, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 1098, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1098, 7, actDay);

		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1099, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 1099, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1099, 7, actVoucherList);

		System.out.println(actVoucherList);
		System.out.println(expVouchersList);

		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1100, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 1100, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1100, 7, actOpenWindowsCount);


	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 1101, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1101, 7, actDocNo);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomer=customerAccountTxt.getAttribute("value");
		String expCustomer=excelReader.getCellData(xlSheetName, 1102, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1102, 7, actCustomer);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	String actWarehouse = select1stRow_1stColumn.getText();
	 	String expWarehouse = excelReader.getCellData(xlSheetName, 1103, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1103, 7, actWarehouse);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	 	String actItem = select1stRow_2ndColumn.getText();
	 	String expItem = excelReader.getCellData(xlSheetName, 1104, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1104, 7, actItem);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
	 	String actTaxCode = select1stRow_3rdColumn.getText();
	 	String expTaxCode = excelReader.getCellData(xlSheetName, 1105, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1105, 7, actTaxCode);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
	 	String actSalesAccount = select1stRow_4thColumn.getText();
	 	String expSalesAccount = excelReader.getCellData(xlSheetName, 1106, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1106, 7, actSalesAccount);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	 	String actUnits = select1stRow_5thColumn.getText();
	 	String expUnits = excelReader.getCellData(xlSheetName, 1107, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1107, 7, actUnits);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	 	String actQuantity = select1stRow_11thColumn.getText();
	 	String expQuantity = excelReader.getCellData(xlSheetName, 1108, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1108, 7, actQuantity);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
	 	String actLSalesOrder = select1stRow_12thColumn.getText();
	 	String expLSalesOrder = excelReader.getCellData(xlSheetName, 1109, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1109, 7, actLSalesOrder);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	 	String actRate = select1stRow_14thColumn.getText();
	 	String expRate = excelReader.getCellData(xlSheetName, 1110, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1110, 7, actRate);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
	 	String actGross = select1stRow_15thColumn.getText();
	 	String expGross = excelReader.getCellData(xlSheetName, 1111, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1111, 7, actGross);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
	 	String actBatch = select1stRow_19thColumn.getText();
	 	String expBatch = excelReader.getCellData(xlSheetName, 1112, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1112, 7, actBatch);

	 	
		System.out.println("**********************checkFastMovingItemsReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
		System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
		System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
		System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
		System.out.println("Tax Code         : "+actTaxCode       +"  Value Expected  "+expTaxCode);
		System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
		System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
		System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
		System.out.println("L Sales Order    : "+actLSalesOrder   +"  Value Expected  "+expLSalesOrder);
		System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
		System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
		System.out.println("Batch            : "+actBatch         +"  Value Expected  "+expBatch);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
				&& actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
				   && actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
				   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesOrder.equalsIgnoreCase(expLSalesOrder)
				   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1097, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1097, 8, resFail);
			return false;
		}
	}
	
	
	
	@FindBy(xpath="(//span[contains(text(),'Slow Moving Item')])[1]")
	private static WebElement slowMovingItem;
	
	
	public boolean checkSlowMovingItemsReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();

		getAction().sendKeys(Keys.END).build().perform();
		//((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,500)");
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(slowMovingItem));
		slowMovingItem.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1118, 5)) 
					|| actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1119, 5)) 
					|| actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1120, 5)) 
					|| actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1121, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().equals(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 1121, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1121, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1122, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1122, 7, actRow2List);
		
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 1123, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1123, 7, actRow3List);
		
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 1124, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1124, 7, actRow4List);
		
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =excelReader.getCellData(xlSheetName, 1125, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1125, 7, actRow5List);
		
		
		
		System.out.println("*********************************checkSlowMovingItemsReport*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1118, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1118, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_585_0_DefaultFilter_0']")
	private static WebElement slowMovingItems_DefaultFilterTxt;
	
	@FindBy(xpath="//*[@id='FilterFields_585_0']/ul[4]/li[1]/div/label/span")
	
	private static WebElement slm_report_FilterItemNameChkbox;
	
	public boolean checkSlowMovingItemsReportWithFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterItemExpansion));
		report_FilterItemExpansion.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(slm_report_FilterItemNameChkbox));
		
		if (report_FilterItemNameChkbox.isSelected()==false) 
		{
			slm_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(slowMovingItems_DefaultFilterTxt));
		slowMovingItems_DefaultFilterTxt.click();
		slowMovingItems_DefaultFilterTxt.sendKeys( excelReader.getCellData(xlSheetName, 1130, 5));
		
		Thread.sleep(2000);
		
		slowMovingItems_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 1130, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1130, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1131, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1131, 7, actRow2List);
		
		System.out.println("****************************checkSlowMovingItemsReportWithFilterOption******************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(slm_report_FilterItemNameChkbox));
		
		if (slm_report_FilterItemNameChkbox.isSelected()==true) 
		{
			slm_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
			
		Thread.sleep(2000);
		
		
		
		String actRowCount=String.valueOf(reportsRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 1132, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1132, 7, actRowCount);

		
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRowCount.equals(expRowCount))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1130, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 1130, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public boolean checkSlowMovingItemsReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		int reportItemGridListCount=reportTableColumn1RowsList.size();
	   
		for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1138, 5)))
             {
         		 System.err.println("Item Name : "+data);
         		 reportTableColumn1RowsList.get(i).click();
         		 Thread.sleep(1000);
         		 break;
              }
  		}
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 1138, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1138, 7, actMonthList);
		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1139, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 1139, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1139, 7, actDay);

		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1140, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 1140, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1140, 7, actVoucherList);
		
		System.out.println(actVouchersList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1141, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 1141, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1141, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 1142, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1142, 7, actDocNo);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomer=customerAccountTxt.getAttribute("value");
		String expCustomer=excelReader.getCellData(xlSheetName, 1143, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1143, 7, actCustomer);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	String actWarehouse = select1stRow_1stColumn.getText();
	 	String expWarehouse = excelReader.getCellData(xlSheetName, 1144, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1144, 7, actWarehouse);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	 	String actItem = select1stRow_2ndColumn.getText();
	 	String expItem = excelReader.getCellData(xlSheetName, 1145, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1145, 7, actItem);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
	 	String actTaxCode = select1stRow_3rdColumn.getText();
	 	String expTaxCode = excelReader.getCellData(xlSheetName, 1146, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1146, 7, actTaxCode);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
	 	String actSalesAccount = select1stRow_4thColumn.getText();
	 	String expSalesAccount =excelReader.getCellData(xlSheetName, 1147, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1147, 7, actSalesAccount);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	 	String actUnits = select1stRow_5thColumn.getText();
	 	String expUnits = excelReader.getCellData(xlSheetName, 1148, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1148, 7, actUnits);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	 	String actQuantity = select1stRow_11thColumn.getText();
	 	String expQuantity = excelReader.getCellData(xlSheetName, 1149, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1149, 7, actQuantity);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
	 	String actLSalesOrder = select1stRow_12thColumn.getText();
	 	String expLSalesOrder = excelReader.getCellData(xlSheetName, 1150, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1150, 7, actLSalesOrder);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	 	String actRate = select1stRow_14thColumn.getText();
	 	String expRate = excelReader.getCellData(xlSheetName, 1151, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1151, 7, actRate);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
	 	String actGross = select1stRow_15thColumn.getText();
	 	String expGross = excelReader.getCellData(xlSheetName, 1152, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1152, 7, actGross);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
	 	String actBatch = select1stRow_19thColumn.getText();
	 	String expBatch = excelReader.getCellData(xlSheetName, 1153, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1153, 7, actBatch);

	 	
		System.out.println("**********************checkSlowMovingItemsReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
		System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
		System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
		System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
		System.out.println("Tax Code         : "+actTaxCode       +"  Value Expected  "+expTaxCode);
		System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
		System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
		System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
		System.out.println("L Sales Order    : "+actLSalesOrder   +"  Value Expected  "+expLSalesOrder);
		System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
		System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
		System.out.println("Batch            : "+actBatch         +"  Value Expected  "+expBatch);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);

		
		if (actMonthlyList.equals(expMonthlyList) && actDayList.equals(expDayList) && actVouchersList.equals(expVouchersList) && actOpenWindowsCount==expOpenWindowsCount
				&& actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
				   && actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
				   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesOrder.equalsIgnoreCase(expLSalesOrder)
				   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1138, 8, resPass);

			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1138, 8, resFail);

			return false;
		}
	}
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[20]")
	private static WebElement peakORLowBalances;
	
	
	public boolean checkPeakAndLowBalancesReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(2000);

		getAction().sendKeys(Keys.END).build().perform();
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,1000)");
		Thread.sleep(2000);
		
		getAction().moveToElement(peakORLowBalances).build().perform();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(peakORLowBalances));
		peakORLowBalances.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1159, 5) )
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1160, 5) ) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1161, 5) ) || 
					actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1162, 5) ))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().equals(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			if (i!=4 && i!=6) 
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 1162, 6); 
		excelReader.setCellData(xlfile, xlSheetName, 1162, 7, actRow1List);
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			if (i!=4 && i!=6) 
			{
				String data = reportsRow2List.get(i).getText();
				reportsRow2ListArray.add(data);
			}
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1163, 6); 
		excelReader.setCellData(xlfile, xlSheetName, 1163, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			if (i!=4 && i!=6) 
			{
				String data = reportsRow3List.get(i).getText();
				reportsRow3ListArray.add(data);
			}
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 1164, 6); 
		excelReader.setCellData(xlfile, xlSheetName, 1164, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			if (i!=4 && i!=6) 
			{
				String data = reportsRow4List.get(i).getText();
				reportsRow4ListArray.add(data);
			}
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List =excelReader.getCellData(xlSheetName, 1165, 6); 
		excelReader.setCellData(xlfile, xlSheetName, 1165, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List = excelReader.getCellData(xlSheetName, 1166, 6); 
		excelReader.setCellData(xlfile, xlSheetName, 1166, 7, actRow5List);
		
		System.out.println("*********************************checkPeakAndLowBalancesReport*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
		{
			   excelReader.setCellData(xlfile, xlSheetName, 1159, 8, resPass);
			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1159, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_586_0_DefaultFilter_0']")
	private static WebElement peakAndLowBalance_DefaultFilterTxt;
	
	@FindBy (xpath="//*[@id='5021']/following-sibling::span")
	private static WebElement palb_report_FilterItemNameChkbox;
	
	
	public boolean checkPeakAndLowBalancesReportWithFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
			report_FilterExpandBtn.click();
			Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterItemExpansion));
		report_FilterItemExpansion.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(palb_report_FilterItemNameChkbox));
		
		if (report_FilterItemNameChkbox.isSelected()==false) 
		{
			palb_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(peakAndLowBalance_DefaultFilterTxt));
		peakAndLowBalance_DefaultFilterTxt.click();
		peakAndLowBalance_DefaultFilterTxt.sendKeys(excelReader.getCellData(xlSheetName, 1172, 5));
		
		Thread.sleep(2000);
		
		peakAndLowBalance_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			if (i!=4 && i!=6) 
			{
				String data = reportsRow1List.get(i).getText();
				reportsRow1ListArray.add(data);
			}
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 1172, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1172, 7, actRow1List);

		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1173, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1173, 7, actRow1List);
		
		System.out.println("****************************checkPeakAndLowBalancesReportWithFilterOption******************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(palb_report_FilterItemNameChkbox));
		
		if (report_FilterItemNameChkbox.isSelected()==true) 
		{
			palb_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
			
		Thread.sleep(2000);
		
	
		String actRowCount=String.valueOf( reportsRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 1174, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1174, 7, actRowCount);
		
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRowCount.equals(expRowCount))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1172, 8, resPass);

			return true;
		} 
		else 
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1172, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	public boolean checkPeakAndLowBalancesReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());

		int reportItemGridListCount=reportTableColumn1RowsList.size();
	   
		for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1178, 5)))
             {
        		 reportTableColumn1RowsList.get(i).click();
        		 Thread.sleep(1000);
        		 break;
             }
		     }
	    
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 1178, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1178, 7, actMonthList);
		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1179, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 1179, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1179, 7, actDay);
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1180, 5)) )
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		

		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 1180, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1180, 7, actVoucherList);


		
		System.out.println(actVoucherList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1181, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 1181, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1181, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 1182, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1182, 7, actDocNo);

		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomer=customerAccountTxt.getAttribute("value");
		String expCustomer=excelReader.getCellData(xlSheetName, 1183, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1183, 7, actCustomer);

		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	String actWarehouse = select1stRow_1stColumn.getText();
	 	String expWarehouse = excelReader.getCellData(xlSheetName, 1184, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1184, 7, actWarehouse);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	 	String actItem = select1stRow_2ndColumn.getText();
	 	String expItem = excelReader.getCellData(xlSheetName, 1185, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1185, 7, actItem);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
	 	String actTaxCode = select1stRow_3rdColumn.getText();
	 	String expTaxCode = excelReader.getCellData(xlSheetName, 1186, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1186, 7, actTaxCode);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
	 	String actSalesAccount = select1stRow_4thColumn.getText();
	 	String expSalesAccount = excelReader.getCellData(xlSheetName, 1187, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1187, 7, actSalesAccount);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	 	String actUnits = select1stRow_5thColumn.getText();
	 	String expUnits = excelReader.getCellData(xlSheetName, 1188, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1188, 7, actUnits);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	 	String actQuantity = select1stRow_11thColumn.getText();
	 	String expQuantity =excelReader.getCellData(xlSheetName, 1189, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1189, 7, actQuantity);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
	 	String actLSalesOrder = select1stRow_12thColumn.getText();
	 	String expLSalesOrder = excelReader.getCellData(xlSheetName, 1190, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1190, 7, actLSalesOrder);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	 	String actRate = select1stRow_14thColumn.getText();
	 	String expRate = excelReader.getCellData(xlSheetName, 1191, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1191, 7, actRate);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
	 	String actGross = select1stRow_15thColumn.getText();
	 	String expGross = excelReader.getCellData(xlSheetName, 1192, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1192, 7, actGross);

	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_19thColumn));
	 	String actBatch = select1stRow_19thColumn.getText();
	 	String expBatch = excelReader.getCellData(xlSheetName, 1193, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1193, 7, actBatch);

	 	
		System.out.println("**********************checkPeakAndLowBalancesReportWithBackTrackOption*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
		System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
		System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
		System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
		System.out.println("Tax Code         : "+actTaxCode       +"  Value Expected  "+expTaxCode);
		System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
		System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
		System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
		System.out.println("L Sales Order    : "+actLSalesOrder   +"  Value Expected  "+expLSalesOrder);
		System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
		System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
		System.out.println("Batch            : "+actBatch         +"  Value Expected  "+expBatch);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);

		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
				&& actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
				   && actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
				   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesOrder.equalsIgnoreCase(expLSalesOrder)
				   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) && actBatch.equalsIgnoreCase(expBatch)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1178, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1178, 8, resFail);
			return false;
		}
	}
	

	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[21]")
	private static WebElement bestSellingItem;
	
	public boolean checkBestSellingItemReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(1000);

		getAction().sendKeys(Keys.END).build().perform();
		Thread.sleep(1000);
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,1000)");
		Thread.sleep(2000);
		
		getAction().moveToElement(bestSellingItem).build().perform();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(bestSellingItem));
		bestSellingItem.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 1199, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1199, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =excelReader.getCellData(xlSheetName, 1200, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1200, 7, actRow2List);
		
		

		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 1201, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1201, 7, actRow3List);
		
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 1202, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1202, 7, actRow4List);
		
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =excelReader.getCellData(xlSheetName, 1203, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1203, 7, actRow1List);
		
		
		
		System.out.println("*********************************checkBestSellingItemReport*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1199, 8, resPass);

			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1199, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	@FindBy (xpath="//select[@id='RITCombobox__2']")
	private static WebElement bestSellingItemsSortingOnDrpdwn;
	
	@FindBy (xpath="//*[@id='dvReportInputs']/div[6]//span")
	private static WebElement includeAllVoucherChkBox;
	
	
	
	public boolean checkBestSellingItemReportWithSortingOnAndIncludeAllVouchersChkBox() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		  excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(bestSellingItemsSortingOnDrpdwn));
		Select s=new Select(bestSellingItemsSortingOnDrpdwn);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(includeAllVoucherChkBox));
		includeAllVoucherChkBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 1209, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 1209, 7, actRow1List);

		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =  excelReader.getCellData(xlSheetName, 1210, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 1210, 7, actRow2List);

		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =  excelReader.getCellData(xlSheetName, 1211, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 1211, 7, actRow3List);

		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List =  excelReader.getCellData(xlSheetName, 1212, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 1212, 7, actRow4List);

		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =  excelReader.getCellData(xlSheetName, 1213, 6);
		 excelReader.setCellData(xlfile, xlSheetName, 1213, 7, actRow5List);

		
		
		System.out.println("*********************************checkBestSellingItemReportWithSortingOnAndIncludeAllVouchersChkBox*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1209, 8, resPass);
			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1209, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	public boolean checkBestSellingItemReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{	
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		int reportItemGridListCount=reportTableColumn1RowsList.size();
	    for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase(  excelReader.getCellData(xlSheetName, 1219, 5)))
             {
        		 reportTableColumn1RowsList.get(i).click();
        		 Thread.sleep(1000);
        		 break;
             }
		     }
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 1219, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1219, 7, actMonthList);
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1220, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 1220, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1220, 7, actDay);

		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1221, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 1221, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1221, 7, actVoucherList);

		
		System.out.println(actVoucherList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1222, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 1222, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1222, 7, actOpenWindowsCount);
		
	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 1223, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1223, 7, actDocNo);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(warehouseTxt));
		String actWarehouse=warehouseTxt.getAttribute("value");
		String expWarehouse=excelReader.getCellData(xlSheetName, 1224, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1224, 7, actWarehouse);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actItem=select1stRow_1stColumn.getText();
		String expItem=excelReader.getCellData(xlSheetName, 1225, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1225, 7, actItem);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
		String actUnits=select1stRow_2ndColumn.getText();
		String expUnits=excelReader.getCellData(xlSheetName, 1226, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1226, 7, actUnits);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
		String actQuantity=select1stRow_3rdColumn.getText();
		String expQuantity=excelReader.getCellData(xlSheetName, 1227, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1227, 7, actQuantity);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
		String actRate=select1stRow_4thColumn.getText();
		String expRate=excelReader.getCellData(xlSheetName, 1228, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1228, 7, actRate);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
		String actGross=select1stRow_5thColumn.getText();
		String expGross=excelReader.getCellData(xlSheetName, 1229, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1229, 7, actGross);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
		String actRMA=select1stRow_9thColumn.getText();
		String expRMA=excelReader.getCellData(xlSheetName, 1230, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1230, 7, actRMA);
		

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actItem2=select2ndRow_1stColumn.getText();
		String expItem2=excelReader.getCellData(xlSheetName, 1231, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1231, 7, actItem2);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_2ndColumn));
		String actUnits2=select2ndRow_2ndColumn.getText();
		String expUnits2=excelReader.getCellData(xlSheetName, 1232, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1232, 7, actUnits2);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_3rdColumn));
		String actQuantity2=select2ndRow_3rdColumn.getText();
		String expQuantity2=excelReader.getCellData(xlSheetName, 1233, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1233, 7, actQuantity2);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_4thColumn));
		String actRate2=select2ndRow_4thColumn.getText();
		String expRate2=excelReader.getCellData(xlSheetName, 1234, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1234, 7, actRate2);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_5thColumn));
		String actGross2=select2ndRow_5thColumn.getText();
		String expGross2=excelReader.getCellData(xlSheetName, 1235, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1235, 7, actGross2);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_9thColumn));
		String actRMA2=select2ndRow_9thColumn.getText();
		String expRMA2=excelReader.getCellData(xlSheetName, 1236, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1236, 7, actRMA2);
		
		System.out.println("**********************checkMultiLevelStockMovementReportBackTrackOption*********************");
		
		System.out.println("*********** Voucher No : "+actDocNo      +"  value expected  "+expDocNo);
		System.out.println("*********** Warehouse  : "+actWarehouse  +"  value expected  "+expWarehouse);
		System.out.println("*********** Item       : "+actItem       +"  value expected  "+expItem);
		System.out.println("*********** Units      : "+actUnits      +"  value expected  "+expUnits);
		System.out.println("*********** Quantity   : "+actQuantity   +"  value expected  "+expQuantity);
		System.out.println("*********** Rate       : "+actRate       +"  value expected  "+expRate);
		System.out.println("*********** Gross      : "+actGross      +"  value expected  "+expGross);
		System.out.println(actRMA);
		System.out.println(expRMA);
		System.out.println("*********** Item2       : "+actItem2       +"  value expected  "+expItem2);
		System.out.println("*********** Units2      : "+actUnits2      +"  value expected  "+expUnits2);
		System.out.println("*********** Quantity2   : "+actQuantity2   +"  value expected  "+expQuantity2);
		System.out.println("*********** Rate2       : "+actRate2       +"  value expected  "+expRate2);
		System.out.println("*********** Gross2      : "+actGross2      +"  value expected  "+expGross2);
		System.out.println(actRMA2);
		System.out.println(expRMA2);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVouchersList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
			 && actDocNo.equalsIgnoreCase(expDocNo) && actWarehouse.equalsIgnoreCase(expWarehouse) && actItem.equalsIgnoreCase(expItem) 
			 && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actRate.equalsIgnoreCase(expRate)
			 && actGross.equalsIgnoreCase(expGross) && actRMA.equalsIgnoreCase(expRMA)
			 && actItem2.equalsIgnoreCase(expItem2) && actUnits2.equalsIgnoreCase(expUnits2) && actQuantity2.equalsIgnoreCase(expQuantity2) 
			 && actRate2.equalsIgnoreCase(expRate2) && actGross2.equalsIgnoreCase(expGross2) && actRMA2.equalsIgnoreCase(expRMA2)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1219, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1219, 8, resFail);
			return false;
		}
	}
	
	
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[23]")
	private static WebElement stockTransferReport;
	

	public boolean checkStockTransferRegisterReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(1000);

		getAction().sendKeys(Keys.END).build().perform();
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,1000)");
		Thread.sleep(2000);
		
		getAction().moveToElement(stockTransferReport).build().perform();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferReport));
		stockTransferReport.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1244, 5)) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1245, 5)) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1246, 5)) || 
					actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1247, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 1247, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1247, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1248, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1248, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =excelReader.getCellData(xlSheetName, 1249, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1249, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 1250, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1250, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =excelReader.getCellData(xlSheetName, 1251, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1251, 7, actRow5List);
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List = excelReader.getCellData(xlSheetName, 1252, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1252, 7, actRow6List);
		
		
		
		int reportsRow7ListCount = reportsRow7List.size();
		ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow7ListCount;i++)
		{
			String data = reportsRow7List.get(i).getText();
			reportsRow7ListArray.add(data);
		}
		String actRow7List = reportsRow7ListArray.toString();
		String expRow7List =excelReader.getCellData(xlSheetName, 1253, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1253, 7, actRow7List);
		
		
		
		int reportsRow8ListCount = reportsRow8List.size();
		ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow8ListCount;i++)
		{
			String data = reportsRow8List.get(i).getText();
			reportsRow8ListArray.add(data);
		}
		String actRow8List = reportsRow8ListArray.toString();
		String expRow8List =excelReader.getCellData(xlSheetName, 1254, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1254, 7, actRow8List);
		
		
		
		int reportsRow9ListCount = reportsRow9List.size();
		ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow9ListCount;i++)
		{
			String data = reportsRow9List.get(i).getText();
			reportsRow9ListArray.add(data);
		}
		String actRow9List = reportsRow9ListArray.toString();
		String expRow9List =excelReader.getCellData(xlSheetName, 1255, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1255, 7, actRow9List);
		
		
		
		int reportsRow10ListCount = reportsRow10List.size();
		ArrayList<String> reportsRow10ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow10ListCount;i++)
		{
			String data = reportsRow10List.get(i).getText();
			reportsRow10ListArray.add(data);
		}
		String actRow10List = reportsRow10ListArray.toString();
		String expRow10List =excelReader.getCellData(xlSheetName, 1256, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1256, 7, actRow10List);
		
		
		
		int reportsRow11ListCount = reportsRow11List.size();
		ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow11ListCount;i++)
		{
			String data = reportsRow11List.get(i).getText();
			reportsRow11ListArray.add(data);
		}
		String actRow11List = reportsRow11ListArray.toString();
		String expRow11List = excelReader.getCellData(xlSheetName, 1257, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1257, 7, actRow11List);
		
		
		
		int reportsRow12ListCount = reportsRow12List.size();
		ArrayList<String> reportsRow12ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow12ListCount;i++)
		{
			String data = reportsRow12List.get(i).getText();
			reportsRow12ListArray.add(data);
		}
		String actRow12List = reportsRow12ListArray.toString();
		String expRow12List =excelReader.getCellData(xlSheetName, 1258, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1258, 7, actRow12List);
		
		
		
		int reportsRow13ListCount = reportsRow13List.size();
		ArrayList<String> reportsRow13ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow13ListCount;i++)
		{
			String data = reportsRow13List.get(i).getText();
			reportsRow13ListArray.add(data);
		}
		String actRow13List = reportsRow13ListArray.toString();
		String expRow13List =excelReader.getCellData(xlSheetName, 1259, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1259, 7, actRow13List);
		
		
		System.out.println("****************************************checkStockTransferRegisterReport******************************************");

		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		System.out.println(actRow7List);
		System.out.println(expRow7List);
		
		System.out.println(actRow8List);
		System.out.println(expRow8List);
		
		System.out.println(actRow9List);
		System.out.println(expRow9List);
		
		System.out.println(actRow10List);
		System.out.println(expRow10List);
		
		System.out.println(actRow11List);
		System.out.println(expRow11List);
		
		System.out.println(actRow12List);
		System.out.println(expRow12List);
		
		System.out.println(actRow13List);
		System.out.println(expRow13List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
			&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
			&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
			&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List) && actRow12List.equalsIgnoreCase(expRow12List) 
			&& actRow13List.equalsIgnoreCase(expRow13List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1244, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1244, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	public boolean checkStockTransferRegisterReportPrintOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		   excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_ReportPrintBtn));
		sl_ReportPrintBtn.click();
		getWaitForAlert();
		Thread.sleep(1000);
		getAlert().accept();
		
		
		Thread.sleep(3000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 1267, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1267, 7, actOpenWindowsCount);


	 	getDriver().switchTo().window(openTabs.get(0));
	 	
	 	Thread.sleep(1000);
	 	
	 	getDriver().switchTo().window(openTabs.get(1)).close();
	 	
	 	getDriver().switchTo().window(openTabs.get(0));
		
		System.out.println("***************************checkStockTransferRegisterReportPrintOption*********************************");
		
		System.out.println("openTabs"+openTabs);
		System.out.println("Open Tabs Count : "+actOpenWindowsCount+"  Value Expected  "+expOpenWindowsCount);
		
		Thread.sleep(1000);
		
		if(actOpenWindowsCount==expOpenWindowsCount)
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			   excelReader.setCellData(xlfile, xlSheetName, 1267, 8, resPass);

			return true;
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 1267, 8, resFail);
			return false;
		}
	}
	
	@FindBy (xpath="//input[@id='MasterSingle__1']")
	private static WebElement fromWarehouseTxt;
	
	@FindBy (xpath="//input[@id='MasterSingle__2']")
	private static WebElement toWarehouseTxt;

	@FindBy (xpath="//select[@id='RITCombobox__3']")
	private static WebElement summaryOrDetailsDrpdwn;
	
	
	public boolean checkStockTransferRegisterReportWithFromAndToWarehouseAndDetailsOptions() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(fromWarehouseTxt));
		fromWarehouseTxt.sendKeys( excelReader.getCellData(xlSheetName, 1272, 5));
		Thread.sleep(2000);
		fromWarehouseTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(toWarehouseTxt));
		toWarehouseTxt.sendKeys( excelReader.getCellData(xlSheetName, 1273, 5));
		Thread.sleep(2000);
		toWarehouseTxt.sendKeys(Keys.TAB);
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(summaryOrDetailsDrpdwn));
		Select s=new Select(summaryOrDetailsDrpdwn);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 1273, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1273, 7, actRow1List);

		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =excelReader.getCellData(xlSheetName, 1274, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1274, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =excelReader.getCellData(xlSheetName, 1275, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1275, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List =excelReader.getCellData(xlSheetName, 1276, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1276, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =excelReader.getCellData(xlSheetName, 1277, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1277, 7, actRow5List);
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List =excelReader.getCellData(xlSheetName, 1278, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1278, 7, actRow6List);
		
		
		
		int reportsRow7ListCount = reportsRow7List.size();
		ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow7ListCount;i++)
		{
			String data = reportsRow7List.get(i).getText();
			reportsRow7ListArray.add(data);
		}
		String actRow7List = reportsRow7ListArray.toString();
		String expRow7List =excelReader.getCellData(xlSheetName, 1279, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1279, 7, actRow7List);
		
		
		
		int reportsRow8ListCount = reportsRow8List.size();
		ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow8ListCount;i++)
		{
			String data = reportsRow8List.get(i).getText();
			reportsRow8ListArray.add(data);
		}
		String actRow8List = reportsRow8ListArray.toString();
		String expRow8List =excelReader.getCellData(xlSheetName, 1280, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1280, 7, actRow8List);
		
		
		
		int reportsRow9ListCount = reportsRow9List.size();
		ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow9ListCount;i++)
		{
			String data = reportsRow9List.get(i).getText();
			reportsRow9ListArray.add(data);
		}
		String actRow9List = reportsRow9ListArray.toString();
		String expRow9List =excelReader.getCellData(xlSheetName, 1281, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1281, 7, actRow9List);
		
		
		
		int reportsRow10ListCount = reportsRow10List.size();
		ArrayList<String> reportsRow10ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow10ListCount;i++)
		{
			String data = reportsRow10List.get(i).getText();
			reportsRow10ListArray.add(data);
		}
		String actRow10List = reportsRow10ListArray.toString();
		String expRow10List =excelReader.getCellData(xlSheetName, 1282, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1282, 7, actRow10List);
		
		
		
		int reportsRow11ListCount = reportsRow11List.size();
		ArrayList<String> reportsRow11ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow11ListCount;i++)
		{
			String data = reportsRow11List.get(i).getText();
			reportsRow11ListArray.add(data);
		}
		String actRow11List = reportsRow11ListArray.toString();
		String expRow11List =excelReader.getCellData(xlSheetName, 1283, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1283, 7, actRow11List);
		
		
		
		int reportsRow12ListCount = reportsRow12List.size();
		ArrayList<String> reportsRow12ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow12ListCount;i++)
		{
			String data = reportsRow12List.get(i).getText();
			reportsRow12ListArray.add(data);
		}
		String actRow12List = reportsRow12ListArray.toString();
		String expRow12List =excelReader.getCellData(xlSheetName, 1284, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1284, 7, actRow12List);
		
		
		
		int reportsRow13ListCount = reportsRow13List.size();
		ArrayList<String> reportsRow13ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow13ListCount;i++)
		{
			String data = reportsRow13List.get(i).getText();
			reportsRow13ListArray.add(data);
		}
		String actRow13List = reportsRow13ListArray.toString();
		String expRow13List =excelReader.getCellData(xlSheetName, 1285, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1285, 7, actRow13List);
		
		
		System.out.println("****************************************checkStockTransferRegisterReportWithFromAndToWarehouseAndDetailsOptions******************************************");

		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		System.out.println(actRow7List);
		System.out.println(expRow7List);
		
		System.out.println(actRow8List);
		System.out.println(expRow8List);
		
		System.out.println(actRow9List);
		System.out.println(expRow9List);
		
		System.out.println(actRow10List);
		System.out.println(expRow10List);
		
		System.out.println(actRow11List);
		System.out.println(expRow11List);
		
		System.out.println(actRow12List);
		System.out.println(expRow12List);
		
		System.out.println(actRow13List);
		System.out.println(expRow13List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
			&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
			&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List) 
			&& actRow10List.equalsIgnoreCase(expRow10List) && actRow11List.equalsIgnoreCase(expRow11List) && actRow12List.equalsIgnoreCase(expRow12List) 
			&& actRow13List.equalsIgnoreCase(expRow13List))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1272, 8, resPass);
			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1272, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public boolean checkStockTransferRegisterReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		  excelReader=new ExcelReader(POJOUtility.getExcelPath());

		
		int Count=reportTableColumn1RowsList.size();
		   
		for (int i = 0; i < Count; i++) 
	    {
             String data=reportTableColumn2RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1292, 5)))
             {
            	 reportTableColumn2RowsList.get(i).click();
        		 Thread.sleep(1000);
        		 break;
             }
		     }
		
		
	    getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 1292, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1292, 7, actMonthList);

		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1293, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 1293, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1293, 7, actDay);
		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1294, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 1294, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1294, 7, actVoucherList);

		System.out.println(actVoucherList);
		System.out.println(expVouchersList);

		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1295, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 1295, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1295, 7, actOpenWindowsCount);
	 	
		getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo = documentNumberTxt.getAttribute("value");
		String expDocNo = excelReader.getCellData(xlSheetName, 1296, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1296, 7, actDocNo);
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockTransferheaderWarehouse1Txt));
		String actFromWarehouse = stockTransferheaderWarehouse1Txt.getAttribute("value");
		String expFromWarehouse = excelReader.getCellData(xlSheetName, 1297, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1297, 7, actFromWarehouse);
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
		String actR1Warehouse = select1stRow_1stColumn.getText();
		String expR1Warehouse = excelReader.getCellData(xlSheetName, 1298, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1298, 7, actR1Warehouse);
		
		String actR1Item = select1stRow_2ndColumn.getText();
		String expR1Item = excelReader.getCellData(xlSheetName, 1299, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1299, 7, actR1Item);
		
		String actR1Units = select1stRow_3rdColumn.getText();
		String expR1Units =excelReader.getCellData(xlSheetName, 1300, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1300, 7, actR1Units);
		
		String actR1Quantity = select1stRow_4thColumn.getText();
		String expR1Quantity =excelReader.getCellData(xlSheetName, 1301, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1301, 7, actR1Quantity);
		
		String actR1Rate = select1stRow_6thColumn.getText();
		String expR1Rate =excelReader.getCellData(xlSheetName, 1302, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1302, 7, actR1Rate);
		
		String actR1Gross = select1stRow_7thColumn.getText();
		String expR1Gross = excelReader.getCellData(xlSheetName, 1303, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1303, 7, actR1Gross);
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select2ndRow_1stColumn));
		String actR2Warehouse = select2ndRow_1stColumn.getText();
		String expR2Warehouse = excelReader.getCellData(xlSheetName, 1304, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1304, 7, actR2Warehouse);
		
		String actR2Item = select2ndRow_2ndColumn.getText();
		String expR2Item =excelReader.getCellData(xlSheetName, 1305, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1305, 7, actR2Item);
		
		String actR2Units = select2ndRow_3rdColumn.getText();
		String expR2Units =excelReader.getCellData(xlSheetName, 1306, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1306, 7, actR2Units);
		
		String actR2Quantity = select2ndRow_4thColumn.getText();
		String expR2Quantity =excelReader.getCellData(xlSheetName, 1307, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1307, 7, actR2Quantity);
		
		String actR2Rate = select2ndRow_6thColumn.getText();
		String expR2Rate = excelReader.getCellData(xlSheetName, 1308, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1308, 7, actR2Rate);
		
		String actR2Gross = select2ndRow_7thColumn.getText();
		String expR2Gross = excelReader.getCellData(xlSheetName, 1309, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1309, 7, actR2Gross);
		
		String actR2Batch = select2ndRow_8thColumn.getText();
		String expR2Batch = excelReader.getCellData(xlSheetName, 1310, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1310, 7, actR2Batch);

		

		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select3rdRow_1stColumn));
		String actR3Warehouse = select3rdRow_1stColumn.getText();
		String expR3Warehouse =excelReader.getCellData(xlSheetName, 1311, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1311, 7, actR3Warehouse);
		
		String actR3Item = select3rdRow_2ndColumn.getText();
		String expR3Item =excelReader.getCellData(xlSheetName, 1312, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1312, 7, actR3Item);
		
		String actR3Units = select3rdRow_3rdColumn.getText();
		String expR3Units =excelReader.getCellData(xlSheetName, 1313, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1313, 7, actR3Units);
		
		String actR3Quantity = select3rdRow_4thColumn.getText();
		String expR3Quantity =excelReader.getCellData(xlSheetName, 1314, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1314, 7, actR3Quantity);
		
		String actR3Rate = select3rdRow_6thColumn.getText();
		String expR3Rate =excelReader.getCellData(xlSheetName, 1315, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1315, 7, actR3Rate);
		
		String actR3Gross = select3rdRow_7thColumn.getText();
		String expR3Gross =excelReader.getCellData(xlSheetName, 1316, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1316, 7, actR3Gross);
		
		String actR3Bins = select3rdRow_9thColumn.getText();
		String expR3Bins =excelReader.getCellData(xlSheetName, 1317, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1317, 7, actR3Bins);
		
		String actR3Bins2 = select3rdRow_10thColumn.getText();
		String expR3Bins2 =excelReader.getCellData(xlSheetName, 1318, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1318, 7, actR3Bins2);
		
		
		
		getWebDriverWait().until(ExpectedConditions.elementToBeClickable(select4thRow_1stColumn));
		String actR4Warehouse = select4thRow_1stColumn.getText();
		String expR4Warehouse =excelReader.getCellData(xlSheetName, 1319, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1319, 7, actR4Warehouse);
		
		String actR4Item = select4thRow_2ndColumn.getText();
		String expR4Item = excelReader.getCellData(xlSheetName, 1320, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1320, 7, actR4Item);
		
		String actR4Units = select4thRow_3rdColumn.getText();
		String expR4Units = excelReader.getCellData(xlSheetName, 1321, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1321, 7, actR4Units);
		
		String actR4Quantity = select4thRow_4thColumn.getText();
		String expR4Quantity = excelReader.getCellData(xlSheetName, 1322, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1322, 7, actR4Quantity);
		
		String actR4Rate = select4thRow_6thColumn.getText();
		String expR4Rate =excelReader.getCellData(xlSheetName, 1323, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1323, 7, actR4Rate);
		
		String actR4Gross = select4thRow_7thColumn.getText();
		String expR4Gross =excelReader.getCellData(xlSheetName, 1324, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1324, 7, actR4Gross);
		
		String actR4RMA = select4thRow_11thColumn.getText();
		String expR4RMA = excelReader.getCellData(xlSheetName, 1325, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1325, 7, actR4RMA);
		
		
		System.out.println("**********************checkStockTransferRegisterReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs       : "+actOpenWindowsCount +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No     : "+actDocNo			+"  Value Expected  "+expDocNo);
		System.out.println("From Warehouse  : "+actFromWarehouse	+"  Value Expected  "+expFromWarehouse);
		
		System.out.println("*******************************ROW1**********************************");
		
		System.out.println("Warehouse       : "+actR1Warehouse			+"  Value Expected  "+expR1Warehouse);
		System.out.println("Item            : "+actR1Item				+"  Value Expected  "+expR1Item);
		System.out.println("Units           : "+actR1Units				+"  Value Expected  "+expR1Units);
		System.out.println("Quantity        : "+actR1Quantity			+"  Value Expected  "+expR1Quantity);
		System.out.println("Rate            : "+actR1Rate				+"  Value Expected  "+expR1Rate);
		System.out.println("Gross           : "+actR1Gross				+"  Value Expected  "+expR1Gross);
		
		
		System.out.println("*******************************ROW2**********************************");
		
		System.out.println("Warehouse       : "+actR2Warehouse			+"  Value Expected  "+expR2Warehouse);
		System.out.println("Item            : "+actR2Item				+"  Value Expected  "+expR2Item);
		System.out.println("Units           : "+actR2Units				+"  Value Expected  "+expR2Units);
		System.out.println("Quantity        : "+actR2Quantity			+"  Value Expected  "+expR2Quantity);
		System.out.println("Rate            : "+actR2Rate				+"  Value Expected  "+expR2Rate);
		System.out.println("Gross           : "+actR2Gross				+"  Value Expected  "+expR2Gross);
		System.out.println("Bins            : "+actR2Batch				+"  Value Expected  "+expR2Batch);
		
		System.out.println("*******************************ROW3**********************************");
		
		System.out.println("Warehouse       : "+actR3Warehouse			+"  Value Expected  "+expR3Warehouse);
		System.out.println("Item            : "+actR3Item				+"  Value Expected  "+expR3Item);
		System.out.println("Units           : "+actR3Units				+"  Value Expected  "+expR3Units);
		System.out.println("Quantity        : "+actR3Quantity			+"  Value Expected  "+expR3Quantity);
		System.out.println("Rate            : "+actR3Rate				+"  Value Expected  "+expR3Rate);
		System.out.println("Gross           : "+actR3Gross				+"  Value Expected  "+expR3Gross);
		System.out.println("Bins            : "+actR3Bins				+"  Value Expected  "+expR3Bins);
		System.out.println("Bins2           : "+actR3Bins2				+"  Value Expected  "+expR3Bins2);
		
		System.out.println("*******************************ROW4**********************************");
		
		System.out.println("Warehouse       : "+actR4Warehouse			+"  Value Expected  "+expR4Warehouse);
		System.out.println("Item            : "+actR4Item				+"  Value Expected  "+expR4Item);
		System.out.println("Units           : "+actR4Units				+"  Value Expected  "+expR4Units);
		System.out.println("Quantity        : "+actR4Quantity			+"  Value Expected  "+expR4Quantity);
		System.out.println("Rate            : "+actR4Rate				+"  Value Expected  "+expR4Rate);
		System.out.println("Gross           : "+actR4Gross				+"  Value Expected  "+expR4Gross);
		System.out.println("RMA             : "+actR4RMA				+"  Value Expected  "+expR4RMA);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
			   && actDocNo.equalsIgnoreCase(expDocNo) && actFromWarehouse.equalsIgnoreCase(expFromWarehouse)
			   
			   
			   && actR1Warehouse.equalsIgnoreCase(expR1Warehouse) && actR1Item.equalsIgnoreCase(expR1Item) && actR1Units.equalsIgnoreCase(expR1Units) 
			   && actR1Quantity.equalsIgnoreCase(expR1Quantity) && actR1Rate.equalsIgnoreCase(expR1Rate) && actR1Gross.equalsIgnoreCase(expR1Gross) 
			   
			   && actR2Warehouse.equalsIgnoreCase(expR2Warehouse) && actR2Item.equalsIgnoreCase(expR2Item) && actR2Units.equalsIgnoreCase(expR2Units) 
			   && actR2Quantity.equalsIgnoreCase(expR2Quantity) && actR2Rate.equalsIgnoreCase(expR2Rate) && actR2Gross.equalsIgnoreCase(expR2Gross) && actR2Batch.equalsIgnoreCase(expR2Batch) 
			   
			   && actR3Warehouse.equalsIgnoreCase(expR3Warehouse) && actR3Item.equalsIgnoreCase(expR3Item) && actR3Units.equalsIgnoreCase(expR3Units) 
			   && actR3Quantity.equalsIgnoreCase(expR3Quantity) && actR3Rate.equalsIgnoreCase(expR3Rate) && actR3Gross.equalsIgnoreCase(expR3Gross) 
			   && actR3Bins.equalsIgnoreCase(expR3Bins) && actR3Bins2.equalsIgnoreCase(expR3Bins2)
			   
			   && actR4Warehouse.equalsIgnoreCase(expR4Warehouse) && actR4Item.equalsIgnoreCase(expR4Item) && actR4Units.equalsIgnoreCase(expR4Units) 
			   && actR4Quantity.equalsIgnoreCase(expR4Quantity) && actR4Rate.equalsIgnoreCase(expR4Rate) && actR4Gross.equalsIgnoreCase(expR4Gross) 
			   && actR4RMA.equalsIgnoreCase(expR4RMA)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 1292, 8, resPass);


			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 1292, 8, resFail);

			return false;
		}
	}

	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[25]")
	private static WebElement stockReservationReport;
	
	public boolean checkStockReservationRegisterReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(1000);

		getAction().sendKeys(Keys.END).build().perform();
		
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,1000)");
		Thread.sleep(2000);
		
		getAction().moveToElement(stockTransferReport).build().perform();
		Thread.sleep(2000);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockReservationReport));
		stockReservationReport.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1332, 5)) 
					|| actName.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1333, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 1333, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1333, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=2;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1334, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1334, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=2;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =excelReader.getCellData(xlSheetName, 1335, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1335, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=2;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List = excelReader.getCellData(xlSheetName, 1336, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1336, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =excelReader.getCellData(xlSheetName, 1337, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1337, 7, actRow5List);
		
		
		
		int reportsRow6ListCount = reportsRow6List.size();
		ArrayList<String> reportsRow6ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow6ListCount;i++)
		{
			String data = reportsRow6List.get(i).getText();
			reportsRow6ListArray.add(data);
		}
		String actRow6List = reportsRow6ListArray.toString();
		String expRow6List =excelReader.getCellData(xlSheetName, 1338, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1338, 7, actRow6List);
		
		
		
		int reportsRow7ListCount = reportsRow7List.size();
		ArrayList<String> reportsRow7ListArray = new ArrayList<String>();
		for(int i=2;i<reportsRow7ListCount;i++)
		{
			String data = reportsRow7List.get(i).getText();
			reportsRow7ListArray.add(data);
		}
		String actRow7List = reportsRow7ListArray.toString();
		String expRow7List =excelReader.getCellData(xlSheetName, 1339, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1339, 7, actRow7List);
		
		
		
		int reportsRow8ListCount = reportsRow8List.size();
		ArrayList<String> reportsRow8ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow8ListCount;i++)
		{
			String data = reportsRow8List.get(i).getText();
			reportsRow8ListArray.add(data);
		}
		String actRow8List = reportsRow8ListArray.toString();
		String expRow8List =excelReader.getCellData(xlSheetName, 1340, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1340, 7, actRow8List);
		
		
		
		int reportsRow9ListCount = reportsRow9List.size();
		ArrayList<String> reportsRow9ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow9ListCount;i++)
		{
			String data = reportsRow9List.get(i).getText();
			reportsRow9ListArray.add(data);
		}
		String actRow9List = reportsRow9ListArray.toString();
		String expRow9List =excelReader.getCellData(xlSheetName, 1341, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1341, 7, actRow9List);
		
		
		System.out.println("*********************************checkStockReservationRegisterReport*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		System.out.println(actRow6List);
		System.out.println(expRow6List);
		
		System.out.println(actRow7List);
		System.out.println(expRow7List);
		
		System.out.println(actRow8List);
		System.out.println(expRow8List);
		
		System.out.println(actRow9List);
		System.out.println(expRow9List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List) && actRow6List.equalsIgnoreCase(expRow6List) 
				&& actRow7List.equalsIgnoreCase(expRow7List) && actRow8List.equalsIgnoreCase(expRow8List) && actRow9List.equalsIgnoreCase(expRow9List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1332, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1332, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_665_0_DefaultFilter_0']")
	private static WebElement stockReservationRegister_DefaultFilterTxt;
	
	
	
	@FindBy (xpath="//*[@id='FilterFields_665_0']/ul[4]/li[1]/div/label/span")
	private static WebElement sr_report_FilterItemNameChkbox;
	
	public boolean checkStockReservationRegisterReportWithFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterItemExpansion));
		report_FilterItemExpansion.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sr_report_FilterItemNameChkbox));
		
		if (report_FilterItemNameChkbox.isSelected()==false) 
		{
			sr_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockReservationRegister_DefaultFilterTxt));
		stockReservationRegister_DefaultFilterTxt.click();
		stockReservationRegister_DefaultFilterTxt.sendKeys( excelReader.getCellData(xlSheetName, 1346, 5));
		
		Thread.sleep(2000);
		
		stockReservationRegister_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(3000);
	
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 1346, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1346, 7, actRow1List);

		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=2;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1347, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1346, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 1348, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1348, 7, actRow3List);
		
		System.out.println("****************************checkStockReservationRegisterReportWithFilterOption******************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sr_report_FilterItemNameChkbox));
		
		if (sr_report_FilterItemNameChkbox.isSelected()==true) 
		{
			sr_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
			
		Thread.sleep(3000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		
		
		String actRowCount=String.valueOf(reportsRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 1349, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1349, 7, actRowCount);

		
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRowCount.equals(expRowCount))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1346, 8, resPass);

			return true;
		} 
		else 
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1346, 8, resPass);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public boolean checkStockReservationRegisterReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());

		int Count=reportTableColumn1RowsList.size();
		   
		for (int i = 0; i < Count; i++) 
	    {
             String data=reportTableColumn2RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase( excelReader.getCellData(xlSheetName, 1352, 5)))
             {
            	 reportTableColumn2RowsList.get(i).click();
        		 Thread.sleep(1000);
        		 break;
             }
		     }
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		Thread.sleep(4000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 1352, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1352, 7, actOpenWindowsCount);
	 	getDriver().switchTo().window(openTabs.get(1));
		
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 1353, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1353, 7, actDocNo);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomer=customerAccountTxt.getAttribute("value");
		String expCustomer=excelReader.getCellData(xlSheetName, 1354, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1354, 7, actCustomer);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	String actWarehouse = select1stRow_1stColumn.getText();
	 	String expWarehouse =excelReader.getCellData(xlSheetName, 1355, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1355, 7, actWarehouse);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	 	String actItem = select1stRow_2ndColumn.getText();
	 	String expItem =excelReader.getCellData(xlSheetName, 1356, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1356, 7, actItem);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
	 	String actTaxCode = select1stRow_3rdColumn.getText();
	 	String expTaxCode =excelReader.getCellData(xlSheetName, 1357, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1357, 7, actTaxCode);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
	 	String actSalesAccount = select1stRow_4thColumn.getText();
	 	String expSalesAccount =excelReader.getCellData(xlSheetName, 1358, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1358, 7, actSalesAccount);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_5thColumn));
	 	String actUnits = select1stRow_5thColumn.getText();
	 	String expUnits = excelReader.getCellData(xlSheetName, 1359, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1359, 7, actUnits);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_11thColumn));
	 	String actQuantity = select1stRow_11thColumn.getText();
	 	String expQuantity =excelReader.getCellData(xlSheetName, 1360, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1360, 7, actQuantity);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_12thColumn));
	 	String actLSalesOrder = select1stRow_12thColumn.getText();
	 	String expLSalesOrder = excelReader.getCellData(xlSheetName, 1361, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1361, 7, actLSalesOrder);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_14thColumn));
	 	String actRate = select1stRow_14thColumn.getText();
	 	String expRate =excelReader.getCellData(xlSheetName, 1362, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1362, 7, actRate);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_15thColumn));
	 	String actGross = select1stRow_15thColumn.getText();
	 	String expGross =excelReader.getCellData(xlSheetName, 1363, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1363, 7, actGross);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_21stColumn));
	 	String actRMA = select1stRow_21stColumn.getText();
	 	String expRMA =excelReader.getCellData(xlSheetName, 1364, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1364, 7, actRMA);
	 	
		System.out.println("**********************checkStockReservationRegisterReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
		System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
		System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
		System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
		System.out.println("Tax Code         : "+actTaxCode       +"  Value Expected  "+expTaxCode);
		System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
		System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
		System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
		System.out.println("L Sales Order    : "+actLSalesOrder   +"  Value Expected  "+expLSalesOrder);
		System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
		System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
		System.out.println("RMA              : "+actRMA           +"  Value Expected  "+expRMA);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		if (actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
			   && actItem.equalsIgnoreCase(expItem) && actTaxCode.equalsIgnoreCase(expTaxCode) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
			   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesOrder.equalsIgnoreCase(expLSalesOrder)
			   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross) && actRMA.equalsIgnoreCase(expRMA))
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 1352, 8, resPass);
			return true;
		} 
		else 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 1352, 8, resFail);
			return false;
		}
	}
	

	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[26]")
	private static WebElement transactionTypeWiseStockReport;

	public boolean checkTransactionTypeWiseStockReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(1000);

		getAction().sendKeys(Keys.END).build().perform();
		((JavascriptExecutor)getDriver()).executeScript("window.scrollBy(0,1000)");
		Thread.sleep(2000);
		
		getAction().moveToElement(transactionTypeWiseStockReport).build().perform();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(transactionTypeWiseStockReport));
		transactionTypeWiseStockReport.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1369, 5)) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1370, 5)) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1371, 5))
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1372, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =  excelReader.getCellData(xlSheetName, 1372, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1372, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1373, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1373, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =excelReader.getCellData(xlSheetName, 1374, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1374, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List =excelReader.getCellData(xlSheetName, 1375, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1375, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =excelReader.getCellData(xlSheetName, 1376, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1376, 7, actRow5List);
		
		
		System.out.println("*********************************checkTransactionTypeWiseStockReport*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1369, 8, resPass);

			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1369, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_670_0_DefaultFilter_0']")
	private static WebElement tranTypeWiseStock_DefaultFilterTxt;
	
	@FindBy(xpath="//*[@id='FilterFields_670_0']/ul[4]/li[1]/div/label/span")
	private static WebElement tranTypeWise_report_FilterItemNameChkbox;
	
	
	public boolean checkTransactionTypeWiseStockReportWithFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterItemExpansion));
		report_FilterItemExpansion.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranTypeWise_report_FilterItemNameChkbox));
		
		if (report_FilterItemNameChkbox.isSelected()==false) 
		{
			tranTypeWise_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranTypeWiseStock_DefaultFilterTxt));
		tranTypeWiseStock_DefaultFilterTxt.click();
		tranTypeWiseStock_DefaultFilterTxt.sendKeys(excelReader.getCellData(xlSheetName, 1380, 5));
		
		Thread.sleep(2000);
		
		tranTypeWiseStock_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =excelReader.getCellData(xlSheetName, 1380, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1380, 7, actRow1List);

		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =excelReader.getCellData(xlSheetName, 1381, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1381, 7, actRow2List);
		
		System.out.println("****************************checkTransactionTypeWiseStockReportWithFilterOption******************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		Thread.sleep(200);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(tranTypeWise_report_FilterItemNameChkbox));
		
		if (tranTypeWise_report_FilterItemNameChkbox.isSelected()==true) 
		{
			tranTypeWise_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
			
		Thread.sleep(4000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		String actRowCount=String.valueOf(reportsRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 1382, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1382, 7, actRowCount);

		
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRowCount.equals(expRowCount))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1380, 8, resPass);
			return true;
		} 
		else 
		{
			excelReader.setCellData(xlfile, xlSheetName, 1380, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	
	
	public boolean checkTransactionTypeWiseStockReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
	    excelReader=new ExcelReader(POJOUtility.getExcelPath());

		int reportItemGridListCount=reportTableColumn1RowsList.size();
	    for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1388, 5)))
             {
        		 reportTableColumn1RowsList.get(i).click();
        		 Thread.sleep(1000);
        		 break;
             }
		     }
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 1388, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1388, 7, actMonthList);

		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1389, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 1389, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1389, 7, actDay);
		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1390, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 1390, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1390, 7, actVoucherList);
		
		System.out.println(actVoucherList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1391, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 1391, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1391, 7, actOpenWindowsCount);

	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo=excelReader.getCellData(xlSheetName, 1392, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1392, 7, actDocNo);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomer=customerAccountTxt.getAttribute("value");
		String expCustomer=excelReader.getCellData(xlSheetName, 1393, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1393, 7, actCustomer);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	String actWarehouse = select1stRow_1stColumn.getText();
	 	String expWarehouse =excelReader.getCellData(xlSheetName, 1394, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1394, 7, actWarehouse);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	 	String actItem = select1stRow_2ndColumn.getText();
	 	String expItem =excelReader.getCellData(xlSheetName, 1395, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1395, 7, actItem);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
	 	String actSalesAccount = select1stRow_3rdColumn.getText();
	 	String expSalesAccount =excelReader.getCellData(xlSheetName, 1396, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1396, 7, actSalesAccount);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
	 	String actUnits = select1stRow_4thColumn.getText();
	 	String expUnits =excelReader.getCellData(xlSheetName, 1397, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1397, 7, actUnits);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
	 	String actQuantity = select1stRow_6thColumn.getText();
	 	String expQuantity =excelReader.getCellData(xlSheetName, 1398, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1398, 7, actQuantity);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
	 	String actLSalesInvoice = select1stRow_7thColumn.getText();
	 	String expLSalesInvoice =excelReader.getCellData(xlSheetName, 1399, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1399, 7, actLSalesInvoice);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	 	String actRate = select1stRow_8thColumn.getText();
	 	String expRate = excelReader.getCellData(xlSheetName, 1400, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1400, 7, actRate);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
	 	String actGross = select1stRow_9thColumn.getText();
	 	String expGross =excelReader.getCellData(xlSheetName, 1401, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1401, 7, actGross);
	 	
		System.out.println("**********************checkTransactionTypeWiseStockReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
		System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
		System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
		System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
		System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
		System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
		System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
		System.out.println("L Sales Order    : "+actLSalesInvoice +"  Value Expected  "+expLSalesInvoice);
		System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
		System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) &&actOpenWindowsCount.equals(expOpenWindowsCount)
			   && actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
			   && actItem.equalsIgnoreCase(expItem) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
			   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesInvoice.equalsIgnoreCase(expLSalesInvoice)
			   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 1388, 8, resPass);
			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			excelReader.setCellData(xlfile, xlSheetName, 1388, 8, resFail);
			return false;
		}
	}
	
	
	
	@FindBy(xpath="//*[@id='navigation_menu']/li[3]/ul/li[3]/ul/li[27]")
	private static WebElement actualConsumptionReport;
	

	public boolean checkActualConsumptionReport() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		 LogoutandLoginwithSU();
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		Thread.sleep(1000);
		getAction().sendKeys(Keys.END).build().perform();
		
		Thread.sleep(1000);
		getAction().moveToElement(actualConsumptionReport).build().perform();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actualConsumptionReport));
		actualConsumptionReport.click();
		
		Thread.sleep(2000);
		
		int rowcount=stockLedgerHometableRowCount.size();
		
		System.out.println(rowcount);
		
		for (int i = 0; i < rowcount; i++) 
		{
			String actName = stockLedgerHometableItemNamesList.get(i).getText();
			
			System.out.println(actName);
			
			if(actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1404, 5)) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1405, 5)) 
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1406, 5))
					|| actName.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1407, 5)))
			{
				stockLedgerHometableItemChkboxList.get(i).click();
			}
		}
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =excelReader.getCellData(xlSheetName, 1407, 6);
		  excelReader.setCellData(xlfile, xlSheetName, 1407, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List = excelReader.getCellData(xlSheetName, 1408, 6);
		  excelReader.setCellData(xlfile, xlSheetName, 1408, 7, actRow2List);
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List = excelReader.getCellData(xlSheetName, 1409, 6);
		  excelReader.setCellData(xlfile, xlSheetName, 1409, 7, actRow3List);
		
		
		
		int reportsRow4ListCount = reportsRow4List.size();
		ArrayList<String> reportsRow4ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow4ListCount;i++)
		{
			String data = reportsRow4List.get(i).getText();
			reportsRow4ListArray.add(data);
		}
		String actRow4List = reportsRow4ListArray.toString();
		String expRow4List =excelReader.getCellData(xlSheetName, 1410, 6);
		  excelReader.setCellData(xlfile, xlSheetName, 1410, 7, actRow4List);
		
		
		
		int reportsRow5ListCount = reportsRow5List.size();
		ArrayList<String> reportsRow5ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow5ListCount;i++)
		{
			String data = reportsRow5List.get(i).getText();
			reportsRow5ListArray.add(data);
		}
		String actRow5List = reportsRow5ListArray.toString();
		String expRow5List =excelReader.getCellData(xlSheetName, 1411, 6);
		  excelReader.setCellData(xlfile, xlSheetName, 1411, 7, actRow5List);
		
		
		System.out.println("*********************************checkActualConsumptionReport*****************************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		System.out.println(actRow4List);
		System.out.println(expRow4List);
		
		System.out.println(actRow5List);
		System.out.println(expRow5List);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List) 
				&& actRow4List.equalsIgnoreCase(expRow4List) && actRow5List.equalsIgnoreCase(expRow5List))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1404, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1404, 8, resFail);
			return false;
		}
	}
	
	
	
	
	
	
	@FindBy (xpath="//input[@id='FOption_680_0_DefaultFilter_0']")
	private static WebElement actualConsumption_DefaultFilterTxt;
	
	@FindBy(xpath="//*[@id='FilterFields_680_0']/ul[4]/li[1]/div/label/span")
	private static WebElement actualCons_report_FilterItemNameChkbox;
	
	public boolean checkActualConsumptionReportWithFilterOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterItemExpansion));
		report_FilterItemExpansion.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actualCons_report_FilterItemNameChkbox));
		
		if (report_FilterItemNameChkbox.isSelected()==false) 
		{
			actualCons_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actualConsumption_DefaultFilterTxt));
		actualConsumption_DefaultFilterTxt.click();
		actualConsumption_DefaultFilterTxt.sendKeys(excelReader.getCellData(xlSheetName, 1417, 5));
		
		Thread.sleep(2000);
		
		actualConsumption_DefaultFilterTxt.sendKeys(Keys.TAB);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
		
		Thread.sleep(2000);
	
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List =excelReader.getCellData(xlSheetName, 1417, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1417, 7, actRow1List);

		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =excelReader.getCellData(xlSheetName, 1418, 6);
        excelReader.setCellData(xlfile, xlSheetName, 1418, 7, actRow2List);
		
		System.out.println("****************************checkActualConsumptionReportWithFilterOption******************************");
		
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterExpandBtn));
		report_FilterExpandBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterBtn));
		report_FilterBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterCustomizeBtn));
		report_FilterCustomizeBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(actualCons_report_FilterItemNameChkbox));
		
		if (report_FilterItemNameChkbox.isSelected()==true) 
		{
			actualCons_report_FilterItemNameChkbox.click();
		}

		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reports_Filter_Customize_OkBtn));
		reports_Filter_Customize_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_Filter_RefreshConditionBtn));
		report_Filter_RefreshConditionBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_FilterOkBtn));
		report_FilterOkBtn.click();
			
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
		
		String actRowCount=String.valueOf(reportsRowsCount.size());
		String expRowCount = excelReader.getCellData(xlSheetName, 1419, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1419, 7, actRowCount);
		
		System.out.println("Reports Row Count : "+actRowCount+"  Value Expected  "+expRowCount);
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRowCount==expRowCount)
		{
			return true;
		} 
		else 
		{
			return false;
		}
	}
	
    
                  
    
    
    
  
	public boolean checkActualConsumptionReportBackTrackOption() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath());
		int reportItemGridListCount=reportTableColumn1RowsList.size();
	    for (int i = 0; i < reportItemGridListCount; i++) 
	    {
             String data=reportTableColumn1RowsList.get(i).getText();
             
             if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1424, 5)))
             {
        		 reportTableColumn1RowsList.get(i).click();
        		 Thread.sleep(1000);
        		 break;
             }
		     }
	    
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_BackTrackBtn));
		sl_BackTrackBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsMonthsTable));
		
		int backTrackMonthlyColumnsCount = backTrackItemDetailsMonthsTableList.size();
		ArrayList<String> actMonthlyList = new ArrayList<String>();
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			actMonthlyList.add(data);
		}
		
		Calendar cal=Calendar.getInstance();
		DateFormat df = new SimpleDateFormat("MMMM yyyy");
		String Row2Month = df.format(cal.getTime());
		
		String actMonthList=actMonthlyList.toString();
		String expMonthlyList= excelReader.getCellData(xlSheetName, 1424, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1424, 7, actMonthList);
		
		System.out.println(actMonthList);
		System.out.println(expMonthlyList);
		
		for(int i=0; i<backTrackMonthlyColumnsCount; i++)
		{
			String data = backTrackItemDetailsMonthsTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1425, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsMonthsTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsDaysTable));
		
		int backTrackDayColumnsCount = backTrackItemDetailsDaysTableList.size();
		ArrayList<String> actDayList = new ArrayList<String>();
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			actDayList.add(data);
		}
		
		DateFormat dt = new SimpleDateFormat("d");
		String Row1Date = dt.format(cal.getTime());
		
		String actDay=actDayList.toString();
		String expDayList= excelReader.getCellData(xlSheetName, 1425, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1425, 7, actDay);
		
		System.out.println(actDay);
		System.out.println(expDayList);
		
		for(int i=0; i<backTrackDayColumnsCount; i++)
		{
			String data = backTrackItemDetailsDaysTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1426, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsDaysTableList.get(i)).build().perform();
				break;
			}
		}
		
		Thread.sleep(2000);

		getFluentWebDriverWait().until(ExpectedConditions.visibilityOf(backTrackItemDetailsVouchersTable));
		
		int backTrackVouchersColumnsCount = backTrackItemDetailsVouchersTableList.size();
		ArrayList<String> actVouchersList = new ArrayList<String>();
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			actVouchersList.add(data);
		}
		
		String actVoucherList=actVouchersList.toString();
		String expVouchersList= excelReader.getCellData(xlSheetName, 1426, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1426, 7, actVoucherList);

		System.out.println(actVoucherList);
		System.out.println(expVouchersList);
		
		for(int i=0; i<backTrackVouchersColumnsCount; i++)
		{
			String data = backTrackItemDetailsVouchersTableList.get(i).getText();
			
			if (data.equalsIgnoreCase(excelReader.getCellData(xlSheetName, 1427, 5))) 
			{
				getAction().doubleClick(backTrackItemDetailsVouchersTableList.get(i)).build().perform();
			}
		}
		
		Thread.sleep(2000);
		
		ArrayList<String> openTabs = new ArrayList<String>(getDriver().getWindowHandles());
		
		String actOpenWindowsCount=String.valueOf(getDriver().getWindowHandles().size());
		String expOpenWindowsCount = excelReader.getCellData(xlSheetName, 1427, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1427, 7, actOpenWindowsCount);


	 	getDriver().switchTo().window(openTabs.get(1));
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(documentNumberTxt));
		String actDocNo=documentNumberTxt.getAttribute("value");
		String expDocNo= excelReader.getCellData(xlSheetName, 1428, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1428, 7, actDocNo);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(customerAccountTxt));
		String actCustomer=customerAccountTxt.getAttribute("value");
		String expCustomer=excelReader.getCellData(xlSheetName, 1429, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1429, 7, actCustomer);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_1stColumn));
	 	String actWarehouse = select1stRow_1stColumn.getText();
	 	String expWarehouse =excelReader.getCellData(xlSheetName, 1430, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1430, 7, actWarehouse);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_2ndColumn));
	 	String actItem = select1stRow_2ndColumn.getText();
	 	String expItem =excelReader.getCellData(xlSheetName, 1431, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1431, 7, actItem);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_3rdColumn));
	 	String actSalesAccount = select1stRow_3rdColumn.getText();
	 	String expSalesAccount =excelReader.getCellData(xlSheetName, 1432, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1432, 7, actSalesAccount);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_4thColumn));
	 	String actUnits = select1stRow_4thColumn.getText();
	 	String expUnits =excelReader.getCellData(xlSheetName, 1433, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1433, 7, actUnits);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_6thColumn));
	 	String actQuantity = select1stRow_6thColumn.getText();
	 	String expQuantity =excelReader.getCellData(xlSheetName, 1434, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1434, 7, actQuantity);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_7thColumn));
	 	String actLSalesInvoice = select1stRow_7thColumn.getText();
	 	String expLSalesInvoice = excelReader.getCellData(xlSheetName, 1435, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1435, 7, actLSalesInvoice);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_8thColumn));
	 	String actRate = select1stRow_8thColumn.getText();
	 	String expRate =excelReader.getCellData(xlSheetName, 1436, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1436, 7, actRate);
	 	
	 	getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(select1stRow_9thColumn));
	 	String actGross = select1stRow_9thColumn.getText();
	 	String expGross =excelReader.getCellData(xlSheetName, 1437, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1437, 7, actGross);
	 	
		System.out.println("**********************checkActualConsumptionReportBackTrackOption*********************");
		 
		System.out.println("Open Tabs        : "+actOpenWindowsCount         +"  Value Expected  "+expOpenWindowsCount);
		
		System.out.println("Document No      : "+actDocNo         +"  Value Expected  "+expDocNo);
		System.out.println("Customer Account : "+actCustomer      +"  Value Expected  "+expCustomer);
		System.out.println("Warehouse        : "+actWarehouse     +"  Value Expected  "+expWarehouse);
		System.out.println("Item             : "+actItem          +"  Value Expected  "+expItem);
		System.out.println("Sales Account    : "+actSalesAccount  +"  Value Expected  "+expSalesAccount);
		System.out.println("Units            : "+actUnits         +"  Value Expected  "+expUnits);
		System.out.println("Quantity         : "+actQuantity      +"  Value Expected  "+expQuantity);
		System.out.println("L Sales Order    : "+actLSalesInvoice +"  Value Expected  "+expLSalesInvoice);
		System.out.println("Rate             : "+actRate          +"  Value Expected  "+expRate);
		System.out.println("Gross            : "+actGross         +"  Value Expected  "+expGross);
		
		getDriver().switchTo().window(openTabs.get(1)).close();
		
		Thread.sleep(1000);
		
		getDriver().switchTo().window(openTabs.get(0));
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackItemDetailsBackBtn));
		BackTrackItemDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackMonthDetailsBackBtn));
		BackTrackMonthDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(BackTrackDateDetailsBackBtn));
		BackTrackDateDetailsBackBtn.click();
		
		Thread.sleep(1000);
		
		if (actMonthList.equals(expMonthlyList) && actDay.equals(expDayList) && actVoucherList.equals(expVouchersList) && actOpenWindowsCount.equals(expOpenWindowsCount)
			   && actDocNo.equalsIgnoreCase(expDocNo) && actCustomer.equalsIgnoreCase(expCustomer) && actWarehouse.equalsIgnoreCase(expWarehouse)
			   && actItem.equalsIgnoreCase(expItem) && actSalesAccount.equalsIgnoreCase(expSalesAccount)
			   && actUnits.equalsIgnoreCase(expUnits) && actQuantity.equalsIgnoreCase(expQuantity) && actLSalesInvoice.equalsIgnoreCase(expLSalesInvoice)
			   && actRate.equalsIgnoreCase(expRate) && actGross.equalsIgnoreCase(expGross)) 
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 1424, 8, resPass);

			return true;
		}
		else
		{
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(report_CloseBtn));
			report_CloseBtn.click();
			getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
			sl_CloseBtn.click();
			 excelReader.setCellData(xlfile, xlSheetName, 1424, 8, resFail);
			return false;
		}
	}
    
    
	@FindBy(xpath="//*[@id='Dashboard_ddlList']/li/a")
    private static WebElement labelDashboard ;
    
	public boolean checkLoginWithUserForItemRestriction() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		 excelReader=new ExcelReader(POJOUtility.getExcelPath()); 
		 LogoutandLoginwithSU();
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameImageDisplay));
		 userNameImageDisplay.click();
			Thread.sleep(2000);
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		 logoutOption.click();
		 
		 Thread.sleep(3000);
		 
		 String unamelt=excelReader.getCellData(xlSheetName, 1445, 5);
			
		 String pawslt=excelReader.getCellData(xlSheetName, 1446, 5);
      
		 LoginPage.enterUserName(unamelt);
     
		 LoginPage.enterPassword(pawslt);
  
		 LoginPage.clickOnSignInBtn();  
        
		 Thread.sleep(5000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
		 String actDashboardLabel                = labelDashboard.getText();
		 String expDashboardLabel                = excelReader.getCellData(xlSheetName, 1446, 6);	
		 excelReader.setCellData(xlfile, xlSheetName, 1446, 7, actDashboardLabel);
		 
		if(actDashboardLabel.equalsIgnoreCase(expDashboardLabel))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1445, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1445, 8, resFail);
			return false;
		}
	}
	
	
	public boolean checkStockReportbyTagReportsWithUserForItemRestriction() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath());
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockReportByTag));
		stockReportByTag.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllItemsChkBox));
		sl_SelectAllItemsChkBox.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_1stRow1stCol));
	
		int reportsRow1ListCount = reportsRow1List.size();
		ArrayList<String> reportsRow1ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow1ListCount;i++)
		{
			String data = reportsRow1List.get(i).getText();
			reportsRow1ListArray.add(data);
		}
		String actRow1List = reportsRow1ListArray.toString();
		String expRow1List = excelReader.getCellData(xlSheetName, 1451, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1451, 7, actRow1List);
		
		

		int reportsRow2ListCount = reportsRow2List.size();
		ArrayList<String> reportsRow2ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow2ListCount;i++)
		{
			String data = reportsRow2List.get(i).getText();
			reportsRow2ListArray.add(data);
		}
		String actRow2List = reportsRow2ListArray.toString();
		String expRow2List =excelReader.getCellData(xlSheetName, 1452, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1452, 7, actRow2List);
		
		
		
		
		int reportsRow3ListCount = reportsRow3List.size();
		ArrayList<String> reportsRow3ListArray = new ArrayList<String>();
		for(int i=1;i<reportsRow3ListCount;i++)
		{
			String data = reportsRow3List.get(i).getText();
			reportsRow3ListArray.add(data);
		}
		String actRow3List = reportsRow3ListArray.toString();
		String expRow3List =excelReader.getCellData(xlSheetName, 1453, 6);
		excelReader.setCellData(xlfile, xlSheetName, 1453, 7, actRow3List);
		
		
		System.out.println("************************checkStockReportbyTagReportsWithUserForItemRestriction********************************");
	
		System.out.println(actRow1List);
		System.out.println(expRow1List);
		
		System.out.println(actRow2List);
		System.out.println(expRow2List);
		
		System.out.println(actRow3List);
		System.out.println(expRow3List);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(reportWithLevel_CloseBtn));
		reportWithLevel_CloseBtn.click();

		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_CloseBtn));
		sl_CloseBtn.click();
		
		if(actRow1List.equalsIgnoreCase(expRow1List) && actRow2List.equalsIgnoreCase(expRow2List) && actRow3List.equalsIgnoreCase(expRow3List))
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1451, 8, resPass);
			return true;
		}
		else
		{
			 excelReader.setCellData(xlfile, xlSheetName, 1451, 8, resFail);
			return false;
		}
	}
	
	
	public boolean checkLoginWithSU() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		excelReader=new ExcelReader(POJOUtility.getExcelPath()); 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameImageDisplay));
		 userNameImageDisplay.click();
		 Thread.sleep(1000);
			
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		 logoutOption.click();
		 
		 Thread.sleep(3000);
		 
		 String unamelt=excelReader.getCellData(xlSheetName, 1458, 5);
			
		 String pawslt=excelReader.getCellData(xlSheetName, 1459, 5);
      
		 LoginPage.enterUserName(unamelt);
     
		 LoginPage.enterPassword(pawslt);
  
		 LoginPage.clickOnSignInBtn();  
        
		 Thread.sleep(5000);
		 
		 getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(labelDashboard));
		// boolean actDashboardLabel                = labelDashboard.isDisplayed();
		 //boolean expDashboardLabel                = true;	
		 
		 String actDashboardLabel=String.valueOf(labelDashboard.isDisplayed());
			String expDashboardLabel = excelReader.getCellData(xlSheetName, 1459, 6);
			excelReader.setCellData(xlfile, xlSheetName, 1459, 7, actDashboardLabel);


		 
		if(actDashboardLabel.equals(expDashboardLabel))
		{
			excelReader.setCellData(xlfile, xlSheetName, 1458, 8, resPass);
			return true;
		}
		else
		{
			excelReader.setCellData(xlfile, xlSheetName, 1458, 8, resFail);
			return false;
		}
	}

	
////
	public boolean checkStockLedgerReportwithFilterOptionasRefineforVouchers() throws InterruptedException, EncryptedDocumentException, InvalidFormatException, IOException
	{
		getDriver().navigate().refresh();
		Thread.sleep(2000);
		 
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryMenu));
		inventoryMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(inventoryReportsMenu));
		inventoryReportsMenu.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(stockLedger));
		stockLedger.click();
		
		Thread.sleep(2000);
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_SelectAllChkBox));
		sl_SelectAllChkBox.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_DateOptionDropdown));
		Select s=new Select(sl_DateOptionDropdown);
		s.selectByIndex(1);
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_OkBtn));
		sl_OkBtn.click();
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_FilterExpandBtn));
		sl_FilterExpandBtn.click();
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(sl_RefineBtn));
		sl_RefineBtn.click();
		Thread.sleep(1000);
		
		
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ld_RefineLookinDropdown));
		ld_RefineLookinDropdown.click();
		
		Select s2=new Select(ld_RefineLookinDropdown);
		s2.selectByValue("2");
		
		Thread.sleep(1000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(ld_RefineOperatorDropdown));
		ld_RefineOperatorDropdown.click();
		
		Select s3=new Select(ld_RefineOperatorDropdown);
		s3.selectByValue("0");
		Thread.sleep(2000);
				
		
		
		int reportsByWarehouseRow1ListCount = ld_RefineData.size();
		System.out.println(reportsByWarehouseRow1ListCount);
		
		ArrayList<String> reportsByWarehouseRow1ListArray = new ArrayList<String>();
		for(int i=0;i<reportsByWarehouseRow1ListCount;i++)
		{
			
			String data = ld_RefineData.get(i).getText();
			reportsByWarehouseRow1ListArray.add(data);
			
			
			
			
		}
		String actList = reportsByWarehouseRow1ListArray.toString();
		String expList = "[Opening Balance, NDT57:SU/IND/TEXT1, NDT57:SU/IND/TEXT3, ExeStk:2, StkTrf:2, NDT50:6, NDT50:7, SalRet:1, Reserved stock, NDT57:SU/IND/TEXT2, NDT50:1]";
		
		System.out.println(actList);
		
		System.out.println(expList);
		
		if(actList.equalsIgnoreCase(expList))
		{
		return true;
		}
		else
		{
			return false;
		}
	}

	
	////
	
	
	


	
	
	
	

	

	
	

	
	

	
	
	

	

	
	public void LogoutandLoginwithSU() throws InterruptedException
	{
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(userNameDisplayLogo));
		userNameDisplayLogo.click();
		Thread.sleep(2000);
		
		getFluentWebDriverWait().until(ExpectedConditions.elementToBeClickable(logoutOption));
		logoutOption.click();
		
		Thread.sleep(3000);
		
		 getDriver().navigate().refresh();
		 Thread.sleep(1999);
		 
		// getDriver().navigate().refresh();
		 //Thread.sleep(1999);

		LoginPage lp=new LoginPage(getDriver()); 

		String unamelt="su";
				      
		String pawslt="su";
				      
		LoginPage.enterUserName(unamelt);

		Thread.sleep(1000);
				
		LoginPage.enterPassword(pawslt);
	
		 Select oSelect = new Select(companyDropDownList);
		 //oSelect.selectByVisibleText(Compname); 
		 oSelect.selectByValue("36");
		
		 LoginPage.clickOnSignInBtn();
		 
		 Thread.sleep(2000);
		 LoginPage.reLogin("su", "su", "Automation Company");
		 
		
		 
		
		
	}
	
	


	
	
	
	
	

	
	
	



	
	
	
	

	
	
				
			

	public Focus10InventoryReportsPage(WebDriver driver)
	{
		PageFactory.initElements(driver, this);
	}


}