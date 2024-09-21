package task8;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelReader {

    public static void main(String[] args) throws Exception {
        String excelFilePath = "C:\\Users\\admir\\Downloads\\Copy of ex.xlsx";

       
        System.out.println("Reading Excel file...");
        readExcelFile(excelFilePath);

      
        System.out.println("Writing new data to Excel file...");
        writeExcelFile(excelFilePath);

        
        System.out.println("Reading Excel file after writing new data...");
        readExcelFile(excelFilePath);
    }

 
    public static void readExcelFile(String excelFilePath) throws IOException {
        Workbook workbook = new XSSFWorkbook(new FileInputStream(excelFilePath));
        Sheet sheet = workbook.getSheetAt(0);  

      
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    System.out.printf("%-20s", cell.getStringCellValue());
                } else if (cell.getCellType() == CellType.NUMERIC) {
                	System.out.printf("%-20d", (int) cell.getNumericCellValue());

                }
            }
            System.out.println(); 
        }

        workbook.close();
        }

  
    public static void writeExcelFile(String excelFilePath) throws IOException {
        FileInputStream inputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);  

        
        int rowCount = sheet.getLastRowNum();
        Row newRow = sheet.createRow(rowCount + 1);

      
        Cell cell1 = newRow.createCell(0);
        cell1.setCellValue("Santhosh");

        Cell cell2 = newRow.createCell(1);
        cell2.setCellValue(21); 

        Cell cell3 = newRow.createCell(2);
        cell3.setCellValue("santhosh@example.com");  

   
        inputStream.close();  
        FileOutputStream outputStream = new FileOutputStream(excelFilePath);
        workbook.write(outputStream);
        workbook.close(); 
        outputStream.close();  
    }
}

