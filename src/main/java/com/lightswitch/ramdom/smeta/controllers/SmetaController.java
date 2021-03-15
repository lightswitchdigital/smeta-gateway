package com.lightswitch.ramdom.smeta.controllers;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.lightswitch.ramdom.smeta.WorkbooksPool;
import com.lightswitch.ramdom.smeta.mappings.Cells;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.*;
import java.io.*;
import java.util.Iterator;

@RequestMapping(path = "/api/v1/calculate")
@RestController
public class SmetaController {

    public Cells mappings;
    public WorkbooksPool pool;

    private final String mappings_path;

    public SmetaController() {

        this.mappings_path = "/src/mappings.json";

        this.pool = new WorkbooksPool();

        this.setMappings();

    }

    private void setMappings() {

        ObjectMapper objectMapper = new ObjectMapper();
        String jsonPath = System.getProperty("user.dir") + this.mappings_path;
        Cells mappings = null;

        try {
            mappings = objectMapper.readValue(new File(jsonPath), Cells.class);
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

        this.mappings = mappings;
    }

    @GetMapping
    @ResponseBody
    public Double calculate(@RequestParam Double count) {
        return this.getCalculatedPrice(count);
    }

    private Double getCalculatedPrice(Double count) {

        XSSFWorkbook workbook = this.getWorkbook();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        Cell valueCell = this.getCell(workbook, this.mappings.getCellID("cell"));
        valueCell.setCellValue(count);

        Cell cell = this.getCell(workbook, this.mappings.getCellID("result"));
        Double cellValue = evaluator.evaluate(cell).getNumberValue();
        cell.getCellType();

        return cellValue;
    }

    // Getters

    private XSSFWorkbook getWorkbook() {
        return this.pool.getWorkbook();
    }

    private XSSFSheet getSheet(XSSFWorkbook workbook) {
        return workbook.getSheetAt(1);
    }

    private Cell getCell(XSSFWorkbook workbook, String cellName) {

        XSSFSheet sheet = this.getSheet(workbook);

        CellReference cr = new CellReference(cellName);
        Row row = sheet.getRow(cr.getRow());

        return row.getCell(cr.getCol());
    }
}
