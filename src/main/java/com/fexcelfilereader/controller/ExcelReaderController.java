package com.fexcelfilereader.controller;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ExcelReaderController {
    
	@GetMapping("/read-excel")
	public List<List<String>> readExcelFile() throws IOException {
	    String filePath = "C:/Users/Ravindhiran/workspace-spring-tool-suite-4-4.20.1.RELEASE/ExcelFIleReader/sample.xlsx";
	    // Read the Excel file
	    return readExcelData(filePath);
	}
	public List<List<String>> readExcelData(String filePath) throws IOException {
	    List<List<String>> data = new ArrayList<>();

	    FileInputStream inputStream = new FileInputStream(filePath);
	    XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

	    XSSFSheet sheet = workbook.getSheetAt(0); // Get the first sheet

	    for (Row row : sheet) {
	        List<String> rowData = new ArrayList<>();
	        for (Cell cell : row) {
	            rowData.add(cell.toString());
	        }
	        data.add(rowData);
	    }

	    return data;
	}


}
