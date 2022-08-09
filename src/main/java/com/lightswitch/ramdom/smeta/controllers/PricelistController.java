package com.lightswitch.ramdom.smeta.controllers;

import com.lightswitch.ramdom.smeta.PricelistMappings;
import com.lightswitch.ramdom.smeta.WorkbooksPool;
import com.lightswitch.ramdom.smeta.mappings.pricelist.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import java.util.ArrayList;
import java.util.Map;

@RestController
public class PricelistController {

    @Autowired
    public PricelistMappings mappings;
    @Autowired
    public WorkbooksPool pool;
    Logger logger = LoggerFactory.getLogger(PricelistController.class);

    public PricelistController() {

    }

    @GetMapping("/api/v1/testtest")
    public void testtest() {
        for (Map.Entry<String, Cell> entry :
                this.mappings.mappings.cells.entrySet()) {
            System.out.println(entry.getKey() + " : " + entry.getValue().def);
        }
        return;
    }

    @GetMapping("/api/v1/pricelist")
    @ResponseBody
    public void pricelist(@RequestParam Map<String, String> params) {
        XSSFWorkbook workbook = this.getWorkbook();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

//        Resetting cells to default values

        this.setDefaultCellValues(workbook);

        for (Map.Entry<String, String> entry : params.entrySet()) {

            String name = entry.getKey();

            try {
                double value = Double.parseDouble(entry.getValue());

                org.apache.poi.ss.usermodel.Cell valueCell = this.getCell(workbook, this.mappings.getCellID(name));
                valueCell.setCellValue(value);

            } catch (NumberFormatException e) {
                String value = entry.getValue();

                org.apache.poi.ss.usermodel.Cell valueCell = this.getCell(workbook, this.mappings.getCellID(name));
                valueCell.setCellValue(value);
            }
        }
    }

    @GetMapping("/validate-pricelist")
    public void validatePricelist() {

        this.logger.info("Launching pricelist cells validation. This might take some time");

        XSSFWorkbook wb = this.getWorkbook();
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        ArrayList<String> maliciousCells = new ArrayList<>();

        long startTime = System.nanoTime();

        for (Map.Entry<String, com.lightswitch.ramdom.smeta.mappings.pricelist.Cell> entry :
                this.mappings.mappings.cells.entrySet()) {

            // We want to test every cell individually, so we set default
            // values for other cells

            this.setDefaultCellValues(wb);

            com.lightswitch.ramdom.smeta.mappings.pricelist.Cell cell = entry.getValue();

            this.setCellValue(wb, cell.id, cell.def);
            this.logger.info("Setting cell " + cell.id + " to value " + cell.def);

            org.apache.poi.ss.usermodel.Cell resultCell = this.getCell(wb, this.mappings.getCellID("result"));

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


    private void setDefaultCellValues(XSSFWorkbook workbook) {
        for (Map.Entry<String, com.lightswitch.ramdom.smeta.mappings.pricelist.Cell> entry : this.mappings.mappings.cells.entrySet()) {
            String id = entry.getValue().id;
            String def = entry.getValue().def;

            this.setCellValue(workbook, id, def);
        }
    }

    private XSSFWorkbook getWorkbook() {
        return this.pool.getWorkbook();
    }

    private XSSFSheet getSheet(XSSFWorkbook workbook) {
        return workbook.getSheetAt(1);
    }

    private org.apache.poi.ss.usermodel.Cell getCell(XSSFWorkbook workbook, String cellName) {

        XSSFSheet sheet = this.getSheet(workbook);

        CellReference cr = new CellReference(cellName);
        Row row = sheet.getRow(cr.getRow());

        return row.getCell(cr.getCol());
    }

    private void setCellValue(XSSFWorkbook workbook, String cellName, Double value) {

        org.apache.poi.ss.usermodel.Cell cell = this.getCell(workbook, cellName);

        cell.setCellValue(value);
    }

    private void setCellValue(XSSFWorkbook workbook, String cellName, String value) {

        org.apache.poi.ss.usermodel.Cell cell = this.getCell(workbook, cellName);

        cell.setCellValue(value);

    }
}
