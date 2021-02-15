package com.sabu.excelapachepoi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ReadFromExcel {
    public void readFromExcelFile(String filepath,
                                  String fileName,
                                  String sheetName) throws IOException {
        //Create an object of File class to open xlsx file
        File file = new File(filepath + "/" + fileName);
        //Create an object of FileInputStream class to read excel file
        FileInputStream inputStream = new FileInputStream(file);

        Workbook testWorkbook = null;

        //Find the file extension by splitting file name in substring  and getting only extension name
        String fileExtensionName = fileName.substring(fileName.indexOf("."));

        //Check condition if the file is xlsx file
        if (fileExtensionName.equals(".xlsx")) {
            //If it is xlsx file then create object of XSSFWorkbook class
            testWorkbook = new XSSFWorkbook(inputStream);
        }
        //Check condition if the file is xls file
        else if (fileExtensionName.equals(".xls")) {
            //If it is xls file then create object of HSSFWorkbook class
            testWorkbook = new HSSFWorkbook(inputStream);
        }

        //Read sheet inside the workbook by its name
        Sheet mySheet = testWorkbook.getSheetAt(0);

        //Find number of rows in excel file
        int rowCount = mySheet.getLastRowNum() - mySheet.getFirstRowNum();

//        //Create a loop over all the rows of excel file to read it
//        for (int i = 0; i < rowCount + 1; i++) {
//            Row row = mySheet.getRow(i);
//            //Create a loop to print cell values in a row
//            for (int j = 0; j < row.getLastCellNum(); j++) {
//                //Print Excel data in console
//                Cell cell = row.getCell(j);
//                if (cell != null) {
//
//                    switch (cell.getCellType()) {
//                        case NUMERIC:
//                            System.out.print(cell.getNumericCellValue() + "t");
//                            break;
//                        case STRING:
//                            System.out.print(cell.getStringCellValue() + "t");
//                            break;
//                        default:
//                            throw new IllegalStateException("Unexpected value: " + cell.getCellType());
//                }
//            }
//        }
//            System.out.println();
//        }

        Iterator<Row> iterator = mySheet.iterator();
        List<Department> departments = new ArrayList<>();
        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Department department = new Department();
            Iterator<Cell> cellIterator = nextRow.iterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "t");
                        break;
                    case STRING:
                        System.out.print(cell.getStringCellValue() + "t");
                        break;
                    default:
                        throw new IllegalStateException("Unexpected value: " + cell.getCellType());

                }
            }
        }

        testWorkbook.close();
        //Close input stream
        inputStream.close();
    }
}
