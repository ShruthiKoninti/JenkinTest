package com.focus.Pages;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.util.Set;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.focus.supporters.ExcelReader;


public class ExcelUtility {
	
	
	
	private static XSSFSheet sheet;
	
	
	 	private  static int startRow=1;
	    private static int totalRow;
	    private static int startColumn=0;
	    private static int totalColumn;
	    
	  
	    
	
	public static String getCellvalue(int rowNumber, int colNumber){
        Row row = sheet.getRow(rowNumber);
        System.out.println(rowNumber);
        System.out.println(row);
        Cell cell = row.getCell(colNumber);
        String cellValue = "";
     

        if(cell.getCellType()==Cell.CELL_TYPE_STRING)
            cellValue = cell.getStringCellValue();
        if(cell.getCellType()==Cell.CELL_TYPE_NUMERIC) {
            NumberFormat nf = new DecimalFormat("#.####");
            cellValue = String.valueOf(nf.format(cell.getNumericCellValue()));
        }
        return cellValue;

    
    }
    


	public static void main(String[] args) throws IOException {
		
			 FileInputStream inputStream = new FileInputStream("C:\\Users\\Rakesh\\Downloads\\focus.xlsx");
	        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
	        XSSFSheet sheet = workbook.getSheetAt(0);

	        
	     int startRowIndex = 1;
		 int startColumnIndex = 0;
	        totalRow = sheet.getPhysicalNumberOfRows();
            totalColumn = sheet.getRow(0).getPhysicalNumberOfCells();
            System.out.println( startRow);
            System.out.println( startRowIndex);
            System.out.println(totalRow - startRowIndex);
            System.out.println( totalColumn);
	//String cellvalue=getCellvalue(1,2);
	
	 Object[][] loginData = new Object[totalRow - startRowIndex][totalColumn];
	// System.out.println(loginData);
     for(int i = startRowIndex; i <= totalRow; i++){
    	 
         for (int j = startColumnIndex; j < totalColumn; j++){
             loginData[i][j] = getCellvalue(i, j);
             System.out.println(loginData[i][j]);
         }
         System.out.println(loginData);
     }
	
	
	
	
	}	 

}
