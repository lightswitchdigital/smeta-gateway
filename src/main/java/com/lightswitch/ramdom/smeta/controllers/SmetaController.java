package com.lightswitch.ramdom.smeta.controllers;

import com.lightswitch.ramdom.smeta.Mappings;
import com.lightswitch.ramdom.smeta.WorkbooksPool;
import com.lightswitch.ramdom.smeta.protocol.Protocol;
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

import java.util.Map;

@CrossOrigin(origins = "http://ramdom.test")
@RestController
public class SmetaController {

    Logger logger = LoggerFactory.getLogger(SmetaController.class);

    @Autowired
    public Mappings mappings;
    public WorkbooksPool pool;

    public SmetaController() {
        this.greet();

        this.pool = new WorkbooksPool();
        this.pool.loadWorkbooks();
    }

    @GetMapping("/test")
    public void test() {
        Protocol.Command.Builder builder = Protocol.Command.newBuilder()
                .setClient("js");

        String[] values = new String[]{"1", "2", "3", "4"};
        for (String value :
                values) {
            builder.putValues(value, "123123");
        }

        Protocol.Command command = builder.build();

        System.out.println(command);
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

    // TODO: 11.09.2021 Сделать проверку на битые клетки
//    @GetMapping("/validate-cells")
//    public void validateCells() {
//        XSSFWorkbook wb = this.getWorkbook();
//        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
//
//        ArrayList<String> maliciousCells = new ArrayList<>();
//
//        long startTime = System.nanoTime();
//
//        for (Map.Entry<String, com.lightswitch.ramdom.smeta.mappings.Cell> entry :
//                this.mappings.mappings.cells.entrySet()) {
//
//            com.lightswitch.ramdom.smeta.mappings.Cell cell = entry.getValue();
//
//            this.setCellValue(wb, cell.id, cell.def);
//            this.logger.debug("setting cell " + cell.id + " with value " + cell.def);
//
//            Cell resultCell = this.getCell(wb, this.mappings.getCellID("result"));
//
//            evaluator.clearAllCachedResultValues();
//            double result = evaluator.evaluate(resultCell).getNumberValue();
//
//            if (result == 0.0) {
////                this.logger.error("Encountered malicious cell: " + cell.id);
//                maliciousCells.add(cell.id);
//            }
//        }
//
//        long endTime = System.nanoTime();
//        long totalTime = endTime - startTime;
//
//        if (maliciousCells.size() > 0) {
//            this.logger.error("Found malicious cells: ");
//            maliciousCells
//                    .stream()
//                    .reduce((id1, id2) -> id1 + ", " + id2)
//                    .ifPresent(this.logger::error);
//        }else {
//            this.logger.info("All cells are clear");
//        }
//    }

    private Double getCalculatedPrice(Map<String, String> params) {

        XSSFWorkbook workbook = this.getWorkbook();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

//        Resetting cells to default values

        for (Map.Entry<String, com.lightswitch.ramdom.smeta.mappings.Cell> entry : this.mappings.mappings.cells.entrySet()) {
            String id = entry.getValue().id;
            String def = entry.getValue().def;

            this.setCellValue(workbook, id, def);
        }

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
