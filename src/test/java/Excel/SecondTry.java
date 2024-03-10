package Excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SecondTry {

	public void readExcel() throws IOException {

		FileInputStream stream = new FileInputStream("C:\\Users\\asana\\OneDrive\\Desktop\\New folder\\Book1.xlsx");

		Workbook book = new XSSFWorkbook(stream);
		Sheet sheet = book.getSheetAt(0);

		Iterator<Row> iterator = sheet.iterator();
		
		while(iterator.hasNext()) {
			
			
			Row next = iterator.next();
			
			Iterator<Cell> iterator2 = next.iterator();
			while(iterator2.hasNext()) {
				
				Cell next2 = iterator2.next();
				
				System.out.println(next2);	
				
			}
			
			
			
			
		}
		
		

		

	}

	public static void main(String[] args) throws IOException {
		
		SecondTry h = new SecondTry();
		h.readExcel();
				
		
		
		
	}
	
	
	
}
