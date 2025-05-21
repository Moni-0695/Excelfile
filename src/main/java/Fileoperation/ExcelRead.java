package Fileoperation;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream file = new FileInputStream("D:\\Java Program\\ExcelReadAndWrite\\src\\main\\java\\TestData.xlsx");
        XSSFWorkbook book = new XSSFWorkbook(file);

        XSSFSheet sheet = book.getSheet("Sheet1");

        int rowCount = sheet.getLastRowNum();
        int columnCount = sheet.getRow(0).getLastCellNum();

        String[][] data = new String[rowCount][columnCount];

        // Start from row 1 to skip headers
        for (int i = 1; i <= rowCount; i++) {
            XSSFRow row = sheet.getRow(i);

            for (int j = 0; j < columnCount; j++) {
                XSSFCell cell = row.getCell(j);

                // Detect cell type
                String value = "";

                if (cell != null) {
                    if (cell.getCellType() == CellType.STRING) {
                        value = cell.getStringCellValue();
                    } else if (cell.getCellType() == CellType.NUMERIC) {
                        value = String.valueOf((int) cell.getNumericCellValue()); // cast to int if needed
                    } else {
                        value = "";
                    }
                }

                System.out.print(value + " | ");
                data[i - 1][j] = value;
            }

            System.out.println();
        }

        // Print from array
        System.out.println("\nPrinting from array:");
        for (String[] row : data) {
            for (String x : row) {
                System.out.print(x + " | ");
            }
            System.out.println();
        }

        book.close();
    }
}
