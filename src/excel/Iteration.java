package excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import javax.swing.text.html.HTMLDocument.Iterator;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFRow.CellIterator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Iteration {

	public static void main(String[] args) throws IOException 
	{
		String excelfile ="/home/active35/Desktop/Tadocs.xlsx";
		
		FileInputStream inputsteam = new FileInputStream(excelfile);
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputsteam); 
		
	
		XSSFSheet sheet = workbook.getSheet("techactive");
		
	 //    int row=sheet.getLastRowNum();
	     
	  //   int cell =sheet.getRow(1).getLastCellNum();
	       //XSSFSheet sheet = workbook.getSheetAt(1);
		
		
		// How to read the data from excel sheet using iterator
		//	How we can work with iterator
		// This is the most popular approch people will use 
		// once you get the sheet we have to read all the rows and columns 
		
		
		
		
		
	
   }
		
		
		
		
		
		
		
		
		
		
		
		
		
	}

}
