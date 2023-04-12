package excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WritingExcelDemo1 {

	public static void main(String[] args) throws IOException 
	{
		
		XSSFWorkbook workbook = new XSSFWorkbook(); // We Have to create a sheet
		
		XSSFSheet sheet = workbook.createSheet("Employee info");
		
		
		Object empdata[][] = {{"EmpId","Name","Designation","location"}, // Header
				              { 1029,"Abhay","TestEngineer","madhyapradesh"},
				              { 1056,"Dilip","TestEngineer","Odisha"},
				              { 1529,"Monika","TestEngineer","GBR"},
				              { 1929,"Shruti","TestEngineer","Bijapur"},
				              { 1329,"Nisarga","TestEngineer","mysore"},
				              { 1829,"Raghu","Developer","Gunutr"},
				              { 1939,"Onkar","Developer","Bihar"},
				              { 1969,"Suman","Developer","Chapra"},
				              { 1969,"Akash","Developer","Godda"},
				              { 1969,"Shrikanth","BA","Bolangir"},
				              { 1969,"Gokul","BA","Tamilnadu"},
			
				              
				};
				
		
		
		// using for loop condition we have check how many rows and columns are there in array 
		
		
		
		int rows = empdata.length;
		int columns = empdata[0].length;
				
		
		System.out.println(rows);  // This Returns 12
		System.out.println(columns); // This Returns 4
		
		 for(int r=0; r<rows;r++) 
		 {
			XSSFRow row= sheet.createRow(r);  // to create a row in Excel we use sheet.createrow(enter the condition);
			
			 for(int c=0; c<columns; c++) 
			 {
				
				    XSSFCell column =row.createCell(c);  // To create a column row.createCell(c);
				    
				    /* once you create a shell we have to capture the data from 2dimensional array and read the data 
				      and update the data in the shell         
				      
				                          */
				    
				    Object value = empdata[r][c];  
				    if(value instanceof String)
				    column.setCellValue((String)value);
				    
				    if(value instanceof Integer)
					    column.setCellValue((Integer)value);
				    
				    if(value instanceof Boolean)
					    column.setCellValue((Boolean)value);
				
			 };
		 };
				
				
				
				
			String filepath = ".\\datafiles\\employeeinfo.xlsx";
			FileOutputStream outputstream = new FileOutputStream(filepath);
			workbook.write(outputstream);

			outputstream.close();
			
			
			System.out.println("Employee data are succesfully printed");
	}

}
