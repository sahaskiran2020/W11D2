package excelprjTestbed;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;

public class ExcelReader {
    public static final String SAMPLE_XLSX_FILE_PATH = "170046_B_AppDev_W10_EffortLogger.xlsm";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        // Creating a Workbook from an Excel file (.xls or .xlsx)
        Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

        // Retrieving the number of sheets in the Workbook
        System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

        /*
           =============================================================
           Iterating over all the sheets in the workbook (Multiple ways)
           =============================================================
        */

        // 1. You can obtain a sheetIterator and iterate over it
        Iterator<Sheet> sheetIterator = workbook.sheetIterator();
        System.out.println("Retrieving Sheets using Iterator");
        while (sheetIterator.hasNext()) {
            Sheet sheet = sheetIterator.next();
            System.out.println("=> " + sheet.getSheetName());
        }

        // 2. Or you can use a for-each loop
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }

        // 3. Or you can use a Java 8 forEach with lambda
        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
        workbook.forEach(sheet -> {
            System.out.println("=> " + sheet.getSheetName());
        });

        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
        Sheet sheet = workbook.getSheetAt(0);

        int counter=sheet.getPhysicalNumberOfRows();
String D= Integer.toString(counter);
System.out.println(D);
int noOfColoumns=sheet.getRow(0).getPhysicalNumberOfCells();
int Y = noOfColoumns;

        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        Sheet sheet1 = workbook.getSheetAt(1);
        int counter1=sheet1.getPhysicalNumberOfRows();
             String U=Integer.toString(counter1);
             System.out.println(U);
           int noOfColoumns1=sheet1.getRow(2).getPhysicalNumberOfCells();
           int H = noOfColoumns1;

           Sheet sheet2 = workbook.getSheetAt(2);
           int counter2=sheet2.getPhysicalNumberOfRows();
                String F=Integer.toString(counter2);
                System.out.println(F);
              int noOfColoumns2=sheet2.getRow(0).getPhysicalNumberOfCells();
              int I = noOfColoumns2;

              Sheet sheet3 = workbook.getSheetAt(3);
              int counter3=sheet3.getPhysicalNumberOfRows();
                   String T=Integer.toString(counter3);
                   System.out.println(T);
                 int noOfColoumns3=sheet3.getRow(0).getPhysicalNumberOfCells();
                 int J = noOfColoumns3;

                 
        // 1. You can obtain a rowIterator and columnIterator and iterate over them
        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
        Iterator<Row> rowIterator = sheet.rowIterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            // Now let's iterate over the columns of the current row
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

        // 2. Or you can use a for-each loop to iterate over the rows and columns
        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
        for (Row row: sheet) {
            for(Cell cell: row) {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            }
            System.out.println();
        }

        // 3. Or you can use Java 8 forEach loop with lambda
        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
        sheet.forEach(row -> {
            row.forEach(cell -> {
                String cellValue = dataFormatter.formatCellValue(cell);
                System.out.print(cellValue + "\t");
            });
            System.out.println();
        });
      
         int colNum = sheet.getRow(0).getLastCellNum();
         int rowNum = sheet.getLastRowNum()+1;
         System.out.println("No.of colums are :"+colNum);
         System.out.println("No.of Rows are :"+rowNum);
         
         System.out.println();
         System.out.println();
         
         System.out.println("Sheet 1 ");
         System.out.println("Number of Rows = " + D);
         System.out.println("Number of Coloumns = " + Y);
         System.out.println("Sheet 5 ");
         System.out.println("Number of Rows = " + U);
         System.out.println("Number of Coloumns = " + H);
         System.out.println("Sheet 3 ");
         System.out.println("Number of Rows = " + F);
         System.out.println("Number of Coloumns = " + I);
         System.out.println("Sheet 2 ");
         System.out.println("Number of Rows = " + T);
         System.out.println("Number of Coloumns = " + J);


         
      
        // Closing the workbook
        workbook.close();
    }
}