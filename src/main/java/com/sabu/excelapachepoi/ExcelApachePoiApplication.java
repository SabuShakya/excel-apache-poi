package com.sabu.excelapachepoi;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;
import java.util.List;

@SpringBootApplication
public class ExcelApachePoiApplication {

    static String filepath = System.getProperty("user.dir") + "/src/main/resources";
    static String fileName = "SampleExcel.xlsx";

    public static void main(String[] args) {
        SpringApplication.run(ExcelApachePoiApplication.class, args);
//        readFromExcelFile();
//        writeToExistingExcel();
//        downloadExcel();
    }

    private static void downloadExcel() {
        DownloadExcelWithData excelWithData = new DownloadExcelWithData();
        List<Book> listBook = excelWithData.getListBook();
        excelWithData.downloadExcel("xlsx", listBook, filepath + "/test.xlsx");
    }

    private static void writeToExistingExcel() {
        WriteToExcel writeToExcel = new WriteToExcel();
        String[] valueToWrite = {
                "Testing",
                "5",
                "10:00:00 AM",
                "02:00:00 PM",
                "Sunday, Tuesday, Thursday"
        };

        try {
            writeToExcel.writeExcel(filepath, fileName, "Departments", valueToWrite);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void readFromExcelFile() {
        ReadFromExcel readFromExcel = new ReadFromExcel();

        try {
            readFromExcel.readFromExcelFile(filepath, fileName, "");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
