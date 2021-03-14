package com.lightswitch.ramdom.smeta.controllers;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.lightswitch.ramdom.smeta.mappings.Cells;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.*;

@RequestMapping(path = "/api/v1/calculate")
@RestController
public class SmetaController {

    XSSFWorkbook workbook;
    Cells mappings;

    public SmetaController() {
        FileInputStream file = null;

        String path = System.getProperty("user.dir") + "/src/static/program.xlsx";

        try {
            file = new FileInputStream(path);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        assert file != null;

        XSSFWorkbook workbook = null;
        try {
            workbook = new XSSFWorkbook(file);
        } catch (IOException e) {
            e.printStackTrace();
        }
        assert workbook != null;

        this.workbook = workbook;

        ObjectMapper objectMapper = new ObjectMapper();
        String jsonPath = System.getProperty("user.dir") + "/src/mappings.json";
        Cells cells = null;

        try {
            cells = objectMapper.readValue(new File(jsonPath), Cells.class);
        } catch (JsonGenerationException e) {
            System.out.println("json generation exception");
            e.printStackTrace();
        } catch (JsonMappingException e) {
            System.out.println("json mappings exception");
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("io exception");
            e.printStackTrace();
        }

        this.mappings = cells;
    }

    @GetMapping
    public Double calculate() {
        return this.getCalculatedPrice();
    }

    private Double getCalculatedPrice() {
        XSSFWorkbook workbook = this.workbook;
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        Cell valueCell = this.getCell(this.mappings.cell);
        valueCell.setCellValue(15);

        Cell cell = this.getCell(this.mappings.result);
        Double cellValue = evaluator.evaluate(cell).getNumberValue();

        return cellValue;
    }


    // Getters

    private XSSFWorkbook getWorkbook() {
        return this.workbook;
    }

    private XSSFSheet getSheet() {
        return this.workbook.getSheetAt(1);
    }

    private Cell getCell(String cellName) {
        XSSFSheet sheet = this.getSheet();

        CellReference cr = new CellReference(cellName);
        Row row = sheet.getRow(cr.getRow());

        return row.getCell(cr.getCol());
    }

    public Double getDoubleCellValue(String cellName) {
        Cell cell = this.getCell(cellName);

        return cell.getNumericCellValue();
    }

    public String getStringCellValue(String cellName) {
        Cell cell = this.getCell(cellName);

        return cell.getStringCellValue();
    }


    public XSSFWorkbook cloneWorkbook() {

        XSSFWorkbook workbook = null;

        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream(4096);
            this.workbook.write(baos);
            ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
            workbook = new XSSFWorkbook(bais);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return workbook;
    }
}
