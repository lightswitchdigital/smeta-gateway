package com.lightswitch.ramdom.smeta.controllers;

import com.lightswitch.ramdom.smeta.PDFExporter;
import com.lightswitch.ramdom.smeta.PricelistMappings;
import com.lightswitch.ramdom.smeta.SmetaMappings;
import com.lightswitch.ramdom.smeta.WorkbooksPool;
import com.lightswitch.ramdom.smeta.requests.GetDocsRequest;
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
import java.text.DecimalFormat;
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
    public PricelistMappings pricelistMappings;

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
    public String calculate(@RequestParam Map<String, String> params) {
        for (var entry : params.entrySet()) {
            System.out.println(entry.getKey() + "/" + entry.getValue());
        }
        Double price = this.getCalculatedPrice(params);

        DecimalFormat df = new DecimalFormat("0.00");
        System.out.println(price);

        return df.format(price);
//        return this.getTestCellValue();
    }

    @PostMapping("/api/v1/get-docs/{dir}")
    @ResponseBody
    public void getDocs(@PathVariable String dir, @RequestBody GetDocsRequest request) {

        if (Objects.equals(dir, "")) {
            dir = "undefined";
        }

        XSSFWorkbook wb = this.getWorkbook();
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        evaluator.clearAllCachedResultValues();

        // Выставляем блять дефолтные значения нахуй
        this.setDefaultCellValues(wb);
        this.setDefaultPricelistValues(wb);

        // Проставляем нахуй данные блять
        for (Map.Entry<String, String> entry : request.data.entrySet()) {
            String name = entry.getKey();

            try {
                double value = Double.parseDouble(entry.getValue());

                Cell valueCell = this.getCell(wb, this.mappings.getCellID(name));
                valueCell.setCellValue(value);

            } catch (NumberFormatException e) {
                String value = entry.getValue();

                Cell valueCell = this.getCell(wb, this.mappings.getCellID(name));
                valueCell.setCellValue(value);
            }
        }

        for (Map.Entry<String, String> entry: request.pricelist.entrySet()) {
            String name = entry.getKey();

            try {
                double value = Double.parseDouble(entry.getValue());

                Cell valueCell = this.getPricelistCell(wb, this.pricelistMappings.getCellID(name));
                valueCell.setCellValue(value);

            } catch (NumberFormatException e) {
                String value = entry.getValue();

                Cell valueCell = this.getPricelistCell(wb, this.pricelistMappings.getCellID(name));
                valueCell.setCellValue(value);
            }
        }

        //////////////
        // Smeta Zak

        XSSFSheet sheetZak = wb.getSheetAt(13);
        System.out.println(sheetZak.getSheetName());

        ArrayList<ArrayList<String>> smetaZak = this.evaluateAndGetSmetaCells(evaluator, sheetZak, 10, 2343);

        try {
            this.exporter.smetaZak(dir, evaluator, sheetZak, smetaZak);
        } catch (IOException e) {
            this.logger.error("could not create pdf file");
        }

        //////////////////
        // Smeta Internal

        XSSFSheet sheetInternal = wb.getSheetAt(12);

        ArrayList<ArrayList<String>> smetaInternal = this.evaluateAndGetSmetaCells(evaluator, sheetInternal, 12, 2496);
        try {
            this.exporter.smetaInternal(dir, evaluator, sheetInternal, smetaInternal);
        } catch (IOException e) {
            this.logger.error("could not create pdf file");
        }

        //////////////////
        // Smeta zak rassh

        XSSFSheet sheetZakRassh = wb.getSheetAt(14);
        System.out.println(sheetZakRassh.getSheetName());

        ArrayList<ArrayList<String>> smetaZakRassh = this.evaluateAndGetSmetaCells(evaluator, sheetInternal, 12, 2488);
        try {
            this.exporter.smetaZakRassh(dir, evaluator, sheetZakRassh, smetaZakRassh);
        } catch (IOException e) {
            this.logger.error("could not create pdf file");
        }

    }

    private ArrayList<ArrayList<String>> evaluateAndGetSmetaCells(FormulaEvaluator evaluator, XSSFSheet sheet, int startRow, int endRow) {

        DecimalFormat df = new DecimalFormat("0.00");

        int counter = 0;
        ArrayList<ArrayList<String>> result = new ArrayList<>();
        for (Row row : sheet) {

            counter++;

            if (counter < startRow) continue;
            if (counter > endRow) break;

            ArrayList<String> values = new ArrayList<>();

            for (Cell cell : row) {

                switch (cell.getCellType()) {

                    case FORMULA:
                        String evaluated = evaluator.evaluate(cell).getStringValue();

                        if (evaluated == null) {
                            double num = evaluator.evaluate(cell).getNumberValue();

                            values.add(df.format(num));
                        } else {

                            values.add(evaluated);
                        }

                        break;

                    case STRING:
                        values.add(cell.getStringCellValue());
                        break;
                }

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

    public void setDefaultPricelistValues(XSSFWorkbook workbook) {
        for (Map.Entry<String, com.lightswitch.ramdom.smeta.mappings.pricelist.Cell> entry : this.pricelistMappings.mappings.cells.entrySet()) {
            String id = entry.getValue().id;
            String def = entry.getValue().def;

            this.setPricelistCellValue(workbook, id, def);
        }
    }

    // Getters

    private XSSFWorkbook getWorkbook() {
        return this.pool.getWorkbook();
    }

    private XSSFSheet getSheet(XSSFWorkbook workbook) {
        return workbook.getSheetAt(2);
    }

//    private XSSFSheet getPricelistSheet(XSSFWorkbook workbook) {
//        return workbook.getSheetAt();
//    }

    private Cell getCell(XSSFWorkbook workbook, String cellName) {
        XSSFSheet sheet = this.getSheet(workbook);

        CellReference cr = new CellReference(cellName);
        Row row = sheet.getRow(cr.getRow());

        return row.getCell(cr.getCol());
    }

    private Cell getPricelistCell(XSSFWorkbook workbook, String cellname) {
        XSSFSheet sheet = workbook.getSheetAt(1);

        CellReference cr = new CellReference(cellname);
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

    private void setPricelistCellValue(XSSFWorkbook workbook, String cellName, String value) {

    }
}
