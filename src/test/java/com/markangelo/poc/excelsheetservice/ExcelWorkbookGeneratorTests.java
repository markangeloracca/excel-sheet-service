package com.markangelo.poc.excelsheetservice;

import dto.Person;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

public class ExcelWorkbookGeneratorTests {

        @Test
        void testGenerateExcel_ValidData(@TempDir Path tempDir) throws IOException {
            String filename = tempDir.resolve("test.xlsx").toString();
            List<Person> persons = List.of(
                    new Person("John", "Doe", 30, "john.doe@example.com", "123 Main St"),
                    new Person("Jane", "Smith", 25, "jane.smith@example.com", "456 Oak Ave")
            );

            Workbook workbook = ExcelWorkbookGenerator.generateExcel(persons);
            assertNotNull(workbook);
            Sheet sheet = workbook.getSheet("Persons");
            assertNotNull(sheet);

            // Check header row
            Row headerRow = sheet.getRow(0);
            assertNotNull(headerRow);
            assertEquals("First Name", headerRow.getCell(0).getStringCellValue());
            assertEquals("Last Name", headerRow.getCell(1).getStringCellValue());
            assertEquals("Age", headerRow.getCell(2).getStringCellValue());
            assertEquals("Email", headerRow.getCell(3).getStringCellValue());
            assertEquals("Address", headerRow.getCell(4).getStringCellValue());

            // Check data rows
            Row row1 = sheet.getRow(1);
            assertNotNull(row1);
            assertEquals("John", row1.getCell(0).getStringCellValue());
            assertEquals("Doe", row1.getCell(1).getStringCellValue());
            assertEquals(30, (int) row1.getCell(2).getNumericCellValue());
            assertEquals("john.doe@example.com", row1.getCell(3).getStringCellValue());
            assertEquals("123 Main St", row1.getCell(4).getStringCellValue());

            Row row2 = sheet.getRow(2);
            assertNotNull(row2);
            assertEquals("Jane", row2.getCell(0).getStringCellValue());
            assertEquals("Smith", row2.getCell(1).getStringCellValue());
            assertEquals(25, (int) row2.getCell(2).getNumericCellValue());
            assertEquals("jane.smith@example.com", row2.getCell(3).getStringCellValue());
            assertEquals("456 Oak Ave", row2.getCell(4).getStringCellValue());
        }

        @Test
        void testGenerateExcel_EmptyList() throws IOException {
            List<Person> persons = List.of();

            Workbook workbook = ExcelWorkbookGenerator.generateExcel(persons);
            assertNotNull(workbook);
            Sheet sheet = workbook.getSheet("Persons");
            assertNotNull(sheet);

            // Check header row exists
            Row headerRow = sheet.getRow(0);
            assertNotNull(headerRow);

            // Check no data rows exist
            assertNull(sheet.getRow(1));
        }

        @Test
        void testGenerateExcel_SinglePerson() throws IOException {
            List<Person> persons = List.of(
                    new Person("Single", "Person", 40, "single@example.com", "Single Address")
            );

            Workbook workbook = ExcelWorkbookGenerator.generateExcel(persons);
            assertNotNull(workbook);
            Sheet sheet = workbook.getSheet("Persons");
            assertNotNull(sheet);

            Row dataRow = sheet.getRow(1);
            assertNotNull(dataRow);
            assertEquals("Single", dataRow.getCell(0).getStringCellValue());
            assertEquals("Person", dataRow.getCell(1).getStringCellValue());
        }

}