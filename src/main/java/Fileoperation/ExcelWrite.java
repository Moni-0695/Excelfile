package Fileoperation;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		XSSFWorkbook book = new XSSFWorkbook();
		
		XSSFSheet sheet = book.createSheet("Sheet1");
		
		Object[][] data = {
	            {"Name", "Age", "Email"},
	            {"John Doe", 30, "john@test.com"},
	            {"Jane Doe", 28, "jane@test.com"},
	            {"Bob Smith", 35, "bob@example.com"},
	            {"Swapnil", 37, "swapnil@example.com"}
	        };

	        int rowCount = 0;

	        for (Object[] rowData : data) {
	            XSSFRow row = sheet.createRow(rowCount++);
	            int columnCount = 0;

	            for (Object field : rowData) {
	                XSSFCell cell = row.createCell(columnCount++);

	                if (field instanceof String) {
	                    cell.setCellValue((String) field);
	                } else if (field instanceof Integer) {
	                    cell.setCellValue((Integer) field);
	                }
	            }
	        }

	       
	        try (FileOutputStream output = new FileOutputStream("D:\\Java Program\\ExcelReadAndWrite\\src\\main\\java\\TestData.xlsx")) {
	            book.write(output);
	            System.out.println("Excel file written successfully.");
	        } 
	        catch (IOException e) {
	            e.printStackTrace();
	        }

	        book.close();
	}

}
