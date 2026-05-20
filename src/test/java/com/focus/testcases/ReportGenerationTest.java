package com.focus.testcases;

import org.automationtesting.excelreport.Xl;

import org.testng.annotations.Test;

import com.focus.base.BaseEngine;

public class ReportGenerationTest extends BaseEngine{
	@Test(priority=98)
	public void reportGeneration() throws Exception
	{
		Xl.generateReport(getCurrentDateFormatWithTime()+"FocusAutomationExcelReport.xlsx");
	 
	
	
}

}
