package com.lightswitch.ramdom.smeta.mappings;

import java.util.Map;

public class Cells {

    public Map<String, Cell> cells;

    public String getCellID(String cellName) {
        return this.getCell(cellName).id;
    }

    private Cell getCell(String cellName) {
        return this.cells.get(cellName);
    }
}


