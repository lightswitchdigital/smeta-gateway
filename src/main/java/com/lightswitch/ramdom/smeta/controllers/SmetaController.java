package com.lightswitch.ramdom.smeta.controllers;

import com.lightswitch.ramdom.smeta.Mappings;
import com.lightswitch.ramdom.smeta.WorkbooksPool;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;

import java.util.Map;

@CrossOrigin(origins = "http://ramdom.test")
@RestController
public class SmetaController {

    @Autowired
    public Mappings mappings;
    public WorkbooksPool pool;

    public SmetaController() {
        this.pool = new WorkbooksPool();
        this.pool.loadWorkbooks();
    }

    @GetMapping("/api/v1/calculate")
    @ResponseBody
    public Double calculate(@RequestParam Map<String, String> params) {
        return this.getCalculatedPrice(params);
//        return this.getTestCellValue();
    }

    private Double getTestCellValue() {
        XSSFWorkbook workbook = this.getWorkbook();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        Cell cell =  this.getCell(workbook, "C11");
        return cell.getNumericCellValue();
//        Double value = evaluator.evaluate(cell).getNumberValue();

//        return value;
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

                Cell valueCell = this.getCell(workbook, this.mappings.mappings.getCellID(name));
                valueCell.setCellValue(value);

            }catch (NumberFormatException e) {
                String value = entry.getValue();

                Cell valueCell = this.getCell(workbook, this.mappings.mappings.getCellID(name));
                valueCell.setCellValue(value);
            }
        }

//        Getting final result
        Cell cell = this.getCell(workbook, this.mappings.mappings.getCellID("result"));

        return evaluator.evaluate(cell).getNumberValue();
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

    private void resetCellsValues(XSSFWorkbook workbook) {
        Integer counter = 0;
        for (Map.Entry<String, com.lightswitch.ramdom.smeta.mappings.Cell> entry : this.mappings.mappings.cells.entrySet()) {
            String id = entry.getValue().id;
            String def = entry.getValue().def;

            this.setCellValue(workbook, id, def);
            counter++;
        }
        System.out.println(counter);
    }
}
