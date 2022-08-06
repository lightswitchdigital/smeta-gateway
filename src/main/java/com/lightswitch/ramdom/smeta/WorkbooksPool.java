package com.lightswitch.ramdom.smeta;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class WorkbooksPool {

    Logger logger = LoggerFactory.getLogger(WorkbooksPool.class);

    private final List<XSSFWorkbook> pool;
    private final String program_path;

    private final Integer workbooks_count;

    public WorkbooksPool() {

        this.pool = new ArrayList<>();
        this.program_path = "/src/program.xlsx";
        this.workbooks_count = 1;

        this.loadWorkbooks();
    }

    public void loadWorkbooks() {
        long startTime = System.nanoTime();

        logger.info("Loading workbooks into the pool (" + this.workbooks_count + ")");

        XSSFWorkbook workbook = this.getProgram();

        for (int i = 0; i < this.workbooks_count; i++) {

            XSSFWorkbook wb = this.cloneWorkbook(workbook);

            logger.info("Added workbook " + (i + 1));

            this.pool.add(wb);
        }

        long elapsedTime = System.nanoTime() - startTime;
        double seconds = (double) elapsedTime / 1_000_000_000.0;

        logger.info("Workbooks loading took: " + Math.round(seconds) + "s");
    }

    public XSSFWorkbook getWorkbook() {
        int i = this.getRandomNumber(this.pool.size());

        return this.pool.get(i);
    }

    private int getRandomNumber(int max) {
        return (int) ((Math.random() * (max)) + 0);
    }

    private XSSFWorkbook getProgram() {
        FileInputStream file = null;

        String path = System.getProperty("user.dir") + this.program_path;

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

        return workbook;
    }

    public XSSFWorkbook cloneWorkbook(XSSFWorkbook wb) {

        XSSFWorkbook workbook = null;

        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream(4096);
            wb.write(baos);
            ByteArrayInputStream bais = new ByteArrayInputStream(baos.toByteArray());
            workbook = new XSSFWorkbook(bais);
        } catch (IOException e) {
            e.printStackTrace();
        }

        return workbook;
    }
}
