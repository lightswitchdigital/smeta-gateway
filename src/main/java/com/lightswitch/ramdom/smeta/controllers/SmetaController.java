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
import java.util.Map;

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
    public Double calculate(@RequestParam Map<String,String> params) {
        return this.getCalculatedPrice(params);
//        return this.getTestCellValue();
    }

    private Double getTestCellValue() {
        XSSFWorkbook workbook = this.getWorkbook();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        Cell cell =  this.getCell(workbook, "C569");
        Double value = evaluator.evaluate(cell).getNumberValue();

        return value;
    }

    private Double getCalculatedPrice(Map<String,String> params) {

        XSSFWorkbook workbook = this.getWorkbook();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

//        Resetting cells to default values
        this.resetCellsValues(workbook);

//        Setting client provided values
        for (Map.Entry<String,String> entry : params.entrySet()) {

            String name = entry.getKey();

            try {
                double value = Double.parseDouble(entry.getValue());

                Cell valueCell = this.getCell(workbook, this.mappings.getCellID(name));
                valueCell.setCellValue(value);

            }catch (NumberFormatException e) {
                String value = entry.getValue();

                Cell valueCell = this.getCell(workbook, this.mappings.getCellID(name));
                valueCell.setCellValue(value);
            }
        }

//        Getting final result
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

    private void setCellValue(XSSFWorkbook workbook, String cellName, Double value) {

        Cell cell = this.getCell(workbook, cellName);

        cell.setCellValue(value);
    }

    private void setCellValue(XSSFWorkbook workbook, String cellName, String value) {

        Cell cell = this.getCell(workbook, cellName);

        cell.setCellValue(value);

    }

    private void resetCellsValues(XSSFWorkbook workbook) {
        for (Map.Entry<String, com.lightswitch.ramdom.smeta.mappings.Cell> entry : this.mappings.cells.entrySet()) {
            String id = entry.getValue().id;
            String def = entry.getValue().def;

            this.setCellValue(workbook, id, def);
        }
    }
}
