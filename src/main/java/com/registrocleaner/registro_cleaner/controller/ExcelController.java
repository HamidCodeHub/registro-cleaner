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

@RestController
@RequestMapping("/api")
public class ExcelController {




    @PostMapping(value = "/clean", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<byte[]> cleanExcel(@RequestParam("file") MultipartFile file) throws IOException {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());
        Sheet sheet = workbook.getSheetAt(0);

        Workbook outputWorkbook = new XSSFWorkbook();
        Sheet outputSheet = outputWorkbook.createSheet("Registro 2");

        int outputRowNum = 0;
        for (Row row : sheet) {
            Cell firstCell = row.getCell(0);
            if (firstCell != null /* && isDocumentRow(firstCell)*/) {
                Row newRow = outputSheet.createRow(outputRowNum++);
                for (int i = 0; i < 14; i++) {
                    Cell inputCell = row.getCell(i);
                    Cell outputCell = newRow.createCell(i);
                    if (inputCell != null) {
                        copyCellValue(inputCell, outputCell);
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

    private boolean isDocumentRow(Cell cell) {
        return cell.getCellType() == CellType.NUMERIC && String.valueOf((long) cell.getNumericCellValue()).length() == 7;
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
