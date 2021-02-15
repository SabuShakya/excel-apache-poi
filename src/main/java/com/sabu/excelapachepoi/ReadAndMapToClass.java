package com.sabu.excelapachepoi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Component
public class ReadAndMapToClass {
    public List<Department> readFromExcelFile(String filepath,
                                              String fileName) throws IOException {
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
        List<Department> departments = new ArrayList<>();
        // start at 1 as 0 is heading
        for (int i = 1; i < rowCount + 1; i++) {
            Row row = mySheet.getRow(i);
            Department department = new Department();
            //Create a loop to print cell values in a row
            for (int j = 0; j < row.getLastCellNum(); j++) {
                //Print Excel data in console
                Cell cell = row.getCell(j);
                if (cell != null) {
                    setValues(cell,department,j);
                }
            }
            departments.add(department);
        }

        testWorkbook.close();
        //Close input stream
        inputStream.close();
        return departments;
    }

    private static void setValues(Cell cell,Department department,Integer cellNo){
        switch (cell.getCellType()) {
            case NUMERIC:
                department.setOpdNo((int) cell.getNumericCellValue());
                break;
            case STRING:
                System.out.println(cell.getRichStringCellValue());
                switch (cellNo){
                    case 0:
                        department.setName(cell.getStringCellValue());
                        break;
                    case 2:
                        department.setStartTime(String.valueOf(cell.getRichStringCellValue()));
                        break;
                    case 3:
                        department.setEndTime(String.valueOf(cell.getRichStringCellValue()));
                        break;
                    case 4:
                        department.setDays(cell.getStringCellValue());
                        break;
                }
                break;
            default:
                break;
        }
    }
}
