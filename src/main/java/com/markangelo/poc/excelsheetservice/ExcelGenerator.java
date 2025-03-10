package com.markangelo.poc.excelsheetservice;

import dto.Person;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelGenerator {

    public static void generateExcel(List<Person> persons, String filename) throws IOException {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fileOut = new FileOutputStream(filename)) {

            Sheet sheet = workbook.createSheet("Persons");

            // Create header row
            Row headerRow = sheet.createRow(0);
            final String[] headers = {"First Name", "Last Name", "Age", "Email", "Address"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Populate data rows
            int rowNum = 1;
            for (Person person : persons) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(person.firstName());
                row.createCell(1).setCellValue(person.lastName());
                row.createCell(2).setCellValue(person.age());
                row.createCell(3).setCellValue(person.email());
                row.createCell(4).setCellValue(person.address());
            }

            workbook.write(fileOut);
        }
    }

    public static void main(String[] args) throws IOException {
        List<Person> persons = List.of(
                new Person("John", "Doe", 30, "john.doe@example.com", "123 Main St"),
                new Person("Jane", "Smith", 25, "jane.smith@example.com", "456 Oak Ave"),
                new Person("Alice", "Johnson", 35, "alice.johnson@example.com", "789 Pine Ln")
        );

        generateExcel(persons, "generated-report.xlsx");
        System.out.println("Excel file generated successfully!");
    }

}