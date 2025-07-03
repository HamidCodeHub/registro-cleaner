package com.registrocleaner.registro_cleaner.controller;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@RestController
@RequestMapping("/api")
public class ExcelController {


    @PostMapping(value = "/clean", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> cleanExcel(@RequestParam("file") MultipartFile file) throws IOException {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Registro 2");

        // Write header row
        String[] headers = {
                "N. doc.", "Data reg.", "Partita IVA", "Riferimento", "Data doc.",
                "Via", "Desc.tipo doc.", "Ragione Sociale", "CAP", "Città",
                "Codice IVA", "Imponibile", "IVA acquisti", "Importo lordo", "Descrizione IVA"
        };
        Row headerRow = outputSheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i).setCellValue(headers[i]);
        }

        boolean inInvoiceSection = false;
        int outputRowNum = 1;

        for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            String firstCellValue = getCellValueAsString(row.getCell(0)).trim();

            // Detect section start
            if (firstCellValue.contains("IVA acquisti: partite singole")) {
                inInvoiceSection = true;
                continue;
            }

            // Exit condition (section ends)
            if (inInvoiceSection && firstCellValue.startsWith("ACQUISTI NON RILEVANTI")) {
                break;
            }

            // Read rows that likely represent document entries (7-digit code and dates)
            if (inInvoiceSection && isDocumentRow(row)) {
                Row outputRow = outputSheet.createRow(outputRowNum++);

                copyRowCells(row, outputRow, new int[]{
                        1, 3, 5, 6, 9, 21, 11, 14, 20, 22, 25, 26, 27, 28
                });

                // Try to extract "Descrizione IVA" from the next row (if any)
                if (i + 1 < sheet.getPhysicalNumberOfRows()) {
                    Row nextRow = sheet.getRow(i + 1);
                    String ivaDescr = getCellValueAsString(nextRow.getCell(10)); // Column where description appears
                    outputRow.createCell(14).setCellValue(ivaDescr);
                }
            }
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        outputWorkbook.write(outputStream);
        outputWorkbook.close();
        workbook.close();

        byte[] cleanedExcel = outputStream.toByteArray();
        HttpHeaders headersHttp = new HttpHeaders();
        headersHttp.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headersHttp.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=registro2.xlsx");

        return ResponseEntity.ok()
                .headers(headersHttp)
                .body(cleanedExcel);
    }


    @PostMapping(value = "/clean2", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> cleanExcel2(@RequestParam("file") MultipartFile file) throws IOException {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet inputSheet = workbook.getSheetAt(0);

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Cleaned");

        int outputRowNum = 0;

        for (Row inputRow : inputSheet) {
            if (inputRow == null) continue;

            Row outputRow = outputSheet.createRow(outputRowNum++);

            // Copy the first column (keys) as-is in column A
            Cell firstCell = inputRow.getCell(0);
            if (firstCell != null) {
                copyCellValue(firstCell, outputRow.createCell(0));
            }

            // Start copying the remaining non-empty cells starting from column B (index 1)
            int outputCol = 1;
            for (int i = 1; i < inputRow.getLastCellNum(); i++) {
                Cell inputCell = inputRow.getCell(i);
                if (inputCell != null && inputCell.getCellType() != CellType.BLANK) {
                    copyCellValue(inputCell, outputRow.createCell(outputCol++));
                }
            }
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        outputWorkbook.write(outputStream);
        outputWorkbook.close();
        workbook.close();

        byte[] cleanedExcel = outputStream.toByteArray();
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=registro2.xlsx");

        return ResponseEntity.ok()
                .headers(headers)
                .body(cleanedExcel);
    }

    @PostMapping(value = "/clean3", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> cleanExcel3(@RequestParam("file") MultipartFile file) throws IOException {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet inputSheet = workbook.getSheetAt(0);

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Cleaned");

        int outputRowNum = 0;
        boolean insideInvoiceTable = false;

        for (int i = 0; i < inputSheet.getPhysicalNumberOfRows(); i++) {
            Row inputRow = inputSheet.getRow(i);
            Row outputRow = outputSheet.createRow(outputRowNum++);

            if (inputRow == null || isRowEmpty(inputRow)) {
                // preserve empty rows
                continue;
            }

            // Read row as text list
            List<String> rowValues = new ArrayList<>();
            for (Cell cell : inputRow) {
                rowValues.add(getCellValueAsString(cell).trim());
            }

            // Detect table header
            if (rowValues.size() >= 5 && rowValues.get(0).contains("N. doc.") && rowValues.get(1).contains("Data reg.")) {
                insideInvoiceTable = true;
            }

            if (insideInvoiceTable) {
                // Copy row starting at column A (index 0)
                int outCol = 0;
                for (int j = 0; j < inputRow.getLastCellNum(); j++) {
                    Cell inputCell = inputRow.getCell(j);
                    if (inputCell != null && inputCell.getCellType() != CellType.BLANK) {
                        copyCellValue(inputCell, outputRow.createCell(outCol++));
                    }
                }

                // If next row is blank or not part of invoice data, stop invoice table block
                if (i + 1 < inputSheet.getPhysicalNumberOfRows()) {
                    Row nextRow = inputSheet.getRow(i + 1);
                    if (nextRow == null || isRowEmpty(nextRow)) {
                        insideInvoiceTable = false;
                    }
                }
            } else {
                // Shift all values starting at column B (index 1), leave column A as-is
                Cell firstCell = inputRow.getCell(0);
                if (firstCell != null) {
                    copyCellValue(firstCell, outputRow.createCell(0));
                }

                int outCol = 1;
                for (int j = 1; j < inputRow.getLastCellNum(); j++) {
                    Cell inputCell = inputRow.getCell(j);
                    if (inputCell != null && inputCell.getCellType() != CellType.BLANK) {
                        copyCellValue(inputCell, outputRow.createCell(outCol++));
                    }
                }
            }
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        outputWorkbook.write(outputStream);
        outputWorkbook.close();
        workbook.close();

        byte[] cleanedExcel = outputStream.toByteArray();
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=registro2.xlsx");

        return ResponseEntity.ok()
                .headers(headers)
                .body(cleanedExcel);
    }

    private boolean isRowEmpty(Row row) {
        for (int i = 0; i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK && !getCellValueAsString(cell).isBlank()) {
                return false;
            }
        }
        return true;
    }

    private boolean isDocumentRow(Row row) {
        Cell cell = row.getCell(1); // N. doc. is typically column 1
        return cell != null && cell.getCellType() == CellType.NUMERIC && String.valueOf((long) cell.getNumericCellValue()).length() == 7;
    }

    private void copyRowCells(Row sourceRow, Row targetRow, int[] indices) {
        for (int i = 0; i < indices.length; i++) {
            Cell srcCell = sourceRow.getCell(indices[i]);
            Cell tgtCell = targetRow.createCell(i);
            if (srcCell != null) {
                copyCellValue(srcCell, tgtCell);
            }
        }
    }

    private String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }


    private void copyCellValue(Cell from, Cell to) {
        switch (from.getCellType()) {
            case STRING:
                to.setCellValue(from.getStringCellValue());
                break;
            case NUMERIC:
                to.setCellValue(from.getNumericCellValue());
                break;
            case BOOLEAN:
                to.setCellValue(from.getBooleanCellValue());
                break;
            case FORMULA:
                to.setCellFormula(from.getCellFormula());
                break;
            default:
                to.setCellValue("");
        }
    }
}
