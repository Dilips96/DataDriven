package excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.*;

public class Radingdata {

	public static void main(String[] args) throws IOException 
	{
		
		String excelfile = "/home/active35/Desktop/SampleData.xlsx";  // we created file path 
		
		FileInputStream inputstream = new FileInputStream(excelfile);  // we connected stream to that file input stream is represting the file 
		
		XSSFWorkbook workbook = new XSSFWorkbook(inputstream);  // We get the Work book from the file.
		
		// Once you get the work book from that work book we have to get the sheet.
		
		//XSSFSheet sheet = workbook.getSheet("SalesOrders");  // we have to specify the name of the sheet 
		
		
		// this will return the sheet object 
		
		
		
	 	XSSFSheet sheet = workbook.getSheetAt(0);  // using index we can write  // so index starts from 0 
		
		
	 	// once you get the sheet sheet contains rows and columns.
	 	// to do that we have to findout how many row columns we have in excel sheet. WE HAVE TO READ THEM
	 	
	 	// TO CHECK HOW MANY COLUMNS ARE THERE IN A SHEEET 
	 	
	 	
	 	int rows = sheet.getLastRowNum();
	 	
	 	// we have to find out the number of cells in the particular row and inside that row how many shells you have 
	 	
	   int column = sheet.getRow(1).getLastCellNum();    // .getLastCellNum will get how cells we have 
		
		
	//	According we have to write 2 different statements which will read the data from excel 
	   
	   for(int r=0;r<=rows;r++)
	   {
		  XSSFRow row= sheet.getRow(r);
		   
		   
		   
		   for(int c=0; c<column;c++)
		   {
			   
			   XSSFCell cell=row.getCell(c); // we have acced with workbook, sheet, row, cells 
			   
			   // From that we have to extract the shell of the data how to extrat the shell object  from the data 
			   
			   
			   // in sheet we have to find which type of data we have depends up on types of data we have to read the
			   // shell   FOR THAT WE NEED TO FIND OUT WHAT IS THE TYPE OF THE SHELL 
			   
			   // HOW TO KNOW THE SHELL TYPE 
			   
			  switch(cell.getCellType())
			  {
			  // This  will return the type of the cell  and based on the type we will read the data
			  
			  case STRING: System.out.print(cell.getStringCellValue()); break;
			  case BOOLEAN: System.out.print(cell.getBooleanCellValue());break;
			  case NUMERIC: System.out.print(cell.getNumericCellValue()); break;
			  
			  
			  }
			
			   System.out.print(" | ");
			   
			   
			   // This is the one way we can read the data from excel by using for loops 
			   
			   
		   }
		   
		   System.out.println();
	   }
	 	
		
		
		
		
		

	}

}
