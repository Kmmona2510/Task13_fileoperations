package Task13;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readoperation {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook book =new XSSFWorkbook("C:\\Users\\MOORTHI\\Desktop\\Mohana\\Write.xlsx");  //Read the  Excelbook
	     XSSFSheet sheet = book.getSheet("Sheet1");    //Read the Sheet
	     
	     int rowcount = sheet.getLastRowNum();                 //for know thw row count
	     int columncount =sheet.getRow(0).getLastCellNum();     //for know the last cell
	     
	     Object [][]  data =new Object[rowcount][columncount];      //create a array to store read data
	     
	     
	     for(int i=0;i<rowcount;i++) {              //get the row
	    	 XSSFRow row= sheet.getRow(i);
	    	 
	    	 for(int j=0;j<columncount;j++) {       //get the cell
	    		XSSFCell cell = row.getCell(j);
	    		
	    		data[i][j] = cell.getStringCellValue();         //get the cell value and put into array
	    		
	    		System.out.println(cell.getStringCellValue());   //print the value
	    		 
	    	 }
	     }
	     
	     book.close();                         //close the book
		
		

	}

}
