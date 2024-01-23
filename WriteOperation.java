package Task13;
/*
 * 
 * Question no1-4
 * 
 */

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteOperation {               //to write a data in excel sheet

	public static void main(String[] args) throws FileNotFoundException, IOException {
		
		XSSFWorkbook book = new XSSFWorkbook();                   //create a book
		XSSFSheet sheet = book.createSheet("Sheet1");             //creata a sheet
		
		Object[][] data = {                                         //create a data in array
				
				{"Name","Age","Email"},  
				{"John Doe",30,"john@test.com"},   
				{"Jane Doe",28,"john@test.com"},
				{"Bob Smith",35,"jacky@example.com"},
				{"Swapnil",37,"Swapnil@example.com"}
		};
		
		int rowCount =0;                          //for row initialize at row count
		
		
		for(Object[] row : data) {                         //create a row
			
			XSSFRow createrow = sheet.createRow(rowCount++);	
			
			int columnCount=0;                   //create a cell
			
			for(Object column: row) {                   
				
				XSSFCell cell = createrow.createCell(columnCount++);
				                                                     
				if(column instanceof String) {                  //for typecasting to particular data  and add value to the cell
					cell.setCellValue((String) column);
				} else if(column instanceof Integer) {
					cell.setCellValue((Integer) column);
				} 
				
				
			}
		}
		
		try(                                                 //give a file path to write the file
			FileOutputStream output= new FileOutputStream("C:\\Users\\MOORTHI\\Desktop\\Mohana\\Write.xlsx");){
			book.write(output);           //using write method to write the book
		}


	}

}
