package com.sabu.excelapachepoi;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class DownloadExcelWithData {

    public void downloadExcel(String excelType,
                                     List<Book> bookList,
                                     String excelFilePath) {
        Workbook testWorkbook3 = null;
        if (excelType.equals("xlsx")) {
            testWorkbook3 = new XSSFWorkbook();
        } else if (excelType.equals("xls")) {
            testWorkbook3 = new HSSFWorkbook();
        }

        Sheet sheet = testWorkbook3.createSheet("mySheet");

        int rowCount = 0;

        for (Book book : bookList) {
            Row row = sheet.createRow(++rowCount);
            writeBook(book, row);
        }

        try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
            testWorkbook3.write(outputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void writeBook(Book aBook, Row row) {
        Cell cell = row.createCell(1);
        cell.setCellValue(aBook.getTitle());

        cell = row.createCell(2);
        cell.setCellValue(aBook.getAuthor());

        cell = row.createCell(3);
        cell.setCellValue(aBook.getPrice());
    }

    public List<Book> getListBook() {
        Book book1 = new Book("Head First Java", "Kathy Serria", 79);
        Book book2 = new Book("Effective Java", "Joshua Bloch", 36);
        Book book3 = new Book("Clean Code", "Robert Martin", 42);
        Book book4 = new Book("Thinking in Java", "Bruce Eckel", 35);

        List<Book> listBook = Arrays.asList(book1, book2, book3, book4);

        return listBook;
    }
}
