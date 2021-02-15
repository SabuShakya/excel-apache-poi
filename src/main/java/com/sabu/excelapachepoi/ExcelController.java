package com.sabu.excelapachepoi;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;
import java.util.List;

@RestController
public class ExcelController {
    private final ReadAndMapToClass readAndMapToClass;

    static String filepath = System.getProperty("user.dir") + "/src/main/resources";
    static String fileName = "SampleExcel.xlsx";

    public ExcelController(ReadAndMapToClass readAndMapToClass) {
        this.readAndMapToClass = readAndMapToClass;
    }

    @GetMapping("/getValuesFromExcel")
    private ResponseEntity<List<Department>> getValuesFromExcel() {
        try {
            return ResponseEntity.ok(readAndMapToClass.readFromExcelFile(filepath, fileName));
        } catch (IOException e) {
            e.printStackTrace();
            throw new NullPointerException("Error occurred!");
        }
    }
}
