# Apache POI #
Apache POI is the pure Java API for reading and writing Excel files in both formats XLS (Excel 2003 and earlier) 
and XLSX (Excel 2007 and later). 

## Dependency ##

```xml
    <dependencies>
        <!--   For Excel 2003  format   -->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi</artifactId>
            <version>4.1.1</version>
        </dependency>
        <!-- For Excel 2007 format     -->
        <dependency>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
            <version>4.1.1</version>
        </dependency>
    </dependencies>
```

## The Apache POI API Basics ##

Two main prefixes:

*HSSF*: denotes the API is for working with Excel 2003 and earlier.

*XSSF*: denotes the API is for working with Excel 2007 and later.

4 interfaces:

*Workbook*: high level representation of an Excel workbook. Concrete implementations are: HSSFWorkbook and XSSFWorkbook.

*Sheet*: high level representation of an Excel worksheet. Typical implementing classes are HSSFSheet and XSSFSheet.

*Row*: high level representation of a row in a spreadsheet. HSSFRow and XSSFRow are two concrete classes.

*Cell*: high level representation of a cell in a row. HSSFCell and XSSFCell are the typical implementing classes.

References:

- https://www.codejava.net/coding/how-to-write-excel-files-in-java-using-apache-poi
- https://www.codejava.net/coding/how-to-read-excel-files-in-java-using-apache-poi
- https://www.guru99.com/all-about-excel-in-selenium-poi-jxl.html