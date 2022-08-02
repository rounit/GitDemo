package com.rs.Mavenjava;



import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	public static void main(String[] args) throws IOException 
	{
		FileInputStream fis = new FileInputStream("C://Users//Rounit Sharma//Documents//TestData.xlsx");
		XSSFWorkbook xs = new XSSFWorkbook(fis);
		int sheets=  xs.getNumberOfSheets();
		for(int i=0;i<sheets;i++)
		{
			if(xs.getSheetName(i).equalsIgnoreCase("Sheet 1"))
			{
				XSSFSheet sh = xs.getSheetAt(i);
		       Iterator<Row> rows = sh.rowIterator();
		       Row firstrow = rows.next();
		       Iterator<Cell> ce =  firstrow.cellIterator();
		       int k=0;
		       int col = 0;
		       while(ce.hasNext())
		       {
		    	   Cell value = ce.next();
		    	   if(value.getStringCellValue().equalsIgnoreCase("TestData"))
		    	   {
		    		   col=k;
		    	   }
		    	   k++;
		       }
		       
		        System.out.println(col);
			}
		}

		
		
	}

}
