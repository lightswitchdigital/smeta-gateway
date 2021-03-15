package com.lightswitch.ramdom.smeta;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;

import java.io.*;
import java.util.*;

public class WorkbooksPool {

    private List<XSSFWorkbook> pool;

    public WorkbooksPool() {

        this.pool = new ArrayList<>();

        long startTime = System.nanoTime();

        System.out.println("loading programs");

        XSSFWorkbook workbook = this.getProgram();

        for (int i = 0; i < 5; i++) {

            XSSFWorkbook wb = this.cloneWorkbook(workbook);

            System.out.println("adding workbook " + i);

            this.pool.add(wb);
        }

        long elapsedTime = System.nanoTime() - startTime;
        double seconds = (double)elapsedTime / 1_000_000_000.0;

        System.out.println("time passed since loading: " + seconds);
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

        String path = System.getProperty("user.dir") + "/src/static/program.xlsx";
        System.out.println(path);

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

    private void resetCellsValues() {
//        Todo
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
