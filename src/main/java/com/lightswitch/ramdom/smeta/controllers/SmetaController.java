package com.lightswitch.ramdom.smeta.controllers;

import com.lightswitch.ramdom.smeta.PDFExporter;
import com.lightswitch.ramdom.smeta.SmetaMappings;
import com.lightswitch.ramdom.smeta.WorkbooksPool;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Map;
import java.util.Objects;

@CrossOrigin(origins = "http://185.225.35.159")
@RestController
public class SmetaController {

    Logger logger = LoggerFactory.getLogger(SmetaController.class);

    @Autowired
    public SmetaMappings mappings;

    @Autowired
    public WorkbooksPool pool;
    @Autowired
    public PDFExporter exporter;

    public SmetaController() {
        this.greet();
    }

    public void greet() {
        System.out.println("    ____  ___    __  _______  ____  __  ___\n" +
                "   / __ \\/   |  /  |/  / __ \\/ __ \\/  |/  /\n" +
                "  / /_/ / /| | / /|_/ / / / / / / / /|_/ / \n" +
                " / _, _/ ___ |/ /  / / /_/ / /_/ / /  / /  \n" +
                "/_/ |_/_/  |_/_/  /_/_____/\\____/_/  /_/   \n" +
                "                                           ");
        System.out.println("|---- HTTP service for smeta calculation");
        System.out.println("|---- Use carefully");
        System.out.println("|---- Made and produced by LightSwitch");
        System.out.println("|---- https://lightswitch.digital");
        System.out.println(" ");
    }

    @GetMapping("/api/v1/calculate")
    @ResponseBody
    public Double calculate(@RequestParam Map<String, String> params) {
        for (var entry : params.entrySet()) {
            System.out.println(entry.getKey() + "/" + entry.getValue());
        }
        Double price = this.getCalculatedPrice(params);

        System.out.println(price);

        return price;
//        return this.getTestCellValue();
    }

    @GetMapping("/api/v1/get-docs")
    @ResponseBody
    public void getDocs(@RequestParam Map<String, String> params) {

        String path = params.get("path");
        if (Objects.equals(path, "")) {
            path = "/smeta/defaults";
        }

        XSSFWorkbook wb = this.getWorkbook();
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        evaluator.clearAllCachedResultValues();

        //////////////
        // Smeta Zak

        XSSFSheet sheetZak = wb.getSheetAt(13);
        System.out.println(sheetZak.getSheetName());

        ArrayList<ArrayList<String>> smetaZak = this.evaluateAndGetSmetaCells(evaluator, sheetZak, 10, 2343, 1);

        try {
            this.exporter.smetaZak(path, evaluator, sheetZak, smetaZak);
        } catch (IOException e) {
            this.logger.error("could not create pdf file");
        }

        //////////////////
        // Smeta Internal

        XSSFSheet sheetInternal = wb.getSheetAt(12);

        ArrayList<ArrayList<String>> smetaInternal = this.evaluateAndGetSmetaCells(evaluator, sheetInternal, 12, 2346, 2);
        try {
            this.exporter.smetaInternal(path, smetaInternal);
        } catch (IOException e) {
            this.logger.error("could not create pdf file");
        }

    }

    private ArrayList<ArrayList<String>> evaluateAndGetSmetaCells(FormulaEvaluator evaluator, XSSFSheet sheet, int startRow, int endRow, int namesCol) {

        int counter = 0;
        ArrayList<ArrayList<String>> result = new ArrayList<>();
        for (Row row : sheet) {

            counter++;

            if (counter < startRow) continue;
            if (counter > endRow) break;

            ArrayList<String> values = new ArrayList<>();

            int col = 1;
            for (Cell cell : row) {

                switch (cell.getCellType()) {

                    case FORMULA:
                        if (col == namesCol) {
                            String evaluated = evaluator.evaluate(cell).getStringValue();
                            values.add(evaluated);
                        } else {
                            double evaluated = evaluator.evaluate(cell).getNumberValue();
                            values.add(Double.toString(evaluated));
                        }
                        break;

                    case STRING:
                        values.add(cell.getStringCellValue());
                        break;
                }

                col++;

            }

            result.add(values);
        }

        return result;
    }

    @GetMapping("/validate-cells")
    public void validateCells() {

        this.logger.info("Launching cells validation. This might take some time");

        XSSFWorkbook wb = this.getWorkbook();
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        ArrayList<String> maliciousCells = new ArrayList<>();

        long startTime = System.nanoTime();

        for (Map.Entry<String, com.lightswitch.ramdom.smeta.mappings.smeta.Cell> entry :
                this.mappings.mappings.cells.entrySet()) {

            // We want to test every cell individually, so we set default
            // values for other cells

            this.setDefaultCellValues(wb);

            com.lightswitch.ramdom.smeta.mappings.smeta.Cell cell = entry.getValue();

            this.setCellValue(wb, cell.id, cell.def);
            this.logger.info("Setting cell " + cell.id + " to value " + cell.def);

            Cell resultCell = this.getCell(wb, this.mappings.getCellID("result"));

            evaluator.clearAllCachedResultValues();
            double result = evaluator.evaluate(resultCell).getNumberValue();

            if (result == 0.0) {
                this.logger.error("Encountered malicious cell: " + cell.id);
                maliciousCells.add(cell.id);
            }
        }

        long endTime = System.nanoTime();
        long totalTime = endTime - startTime;

        if (maliciousCells.size() > 0) {
            this.logger.error("Found malicious cells: ");
            maliciousCells
                    .stream()
                    .reduce((id1, id2) -> id1 + ", " + id2)
                    .ifPresent(this.logger::error);
        } else {
            this.logger.info("All cells are clear");
        }

        this.logger.info("Test took " + totalTime / 1_000_000_000 + " seconds");
    }

    private Double getCalculatedPrice(Map<String, String> params) {

        XSSFWorkbook workbook = this.getWorkbook();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

//        Resetting cells to default values

        this.setDefaultCellValues(workbook);

//        Setting client provided values
        for (Map.Entry<String, String> entry : params.entrySet()) {

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

        evaluator.clearAllCachedResultValues();
        return evaluator.evaluate(cell).getNumberValue();
    }

    private void setDefaultCellValues(XSSFWorkbook workbook) {
        for (Map.Entry<String, com.lightswitch.ramdom.smeta.mappings.smeta.Cell> entry : this.mappings.mappings.cells.entrySet()) {
            String id = entry.getValue().id;
            String def = entry.getValue().def;

            this.setCellValue(workbook, id, def);
        }
    }

    // Getters

    private XSSFWorkbook getWorkbook() {
        return this.pool.getWorkbook();
    }

    private XSSFSheet getSheet(XSSFWorkbook workbook) {
        return workbook.getSheetAt(2);
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
}
