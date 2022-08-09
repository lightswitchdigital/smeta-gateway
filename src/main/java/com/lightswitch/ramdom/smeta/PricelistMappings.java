package com.lightswitch.ramdom.smeta;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.lightswitch.ramdom.smeta.mappings.pricelist.Cell;
import com.lightswitch.ramdom.smeta.mappings.pricelist.Cells;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.IOException;

public class PricelistMappings {

    public Cells mappings;
    Logger logger = LoggerFactory.getLogger(PricelistMappings.class);
    public PricelistMappings() {
        this.load();
    }

    public void load() {
        ObjectMapper objectMapper = new ObjectMapper();
        String jsonPath = System.getProperty("user.dir") + "/src/pricelist.json";
        Cells mappings = null;

        try {
            mappings = objectMapper.readValue(new File(jsonPath), Cells.class);
        } catch (JsonGenerationException e) {
            logger.error("json generation exception");
            e.printStackTrace();
        } catch (JsonMappingException e) {
            logger.error("json mappings exception");
            e.printStackTrace();
        } catch (IOException e) {
            logger.error("io exception");
            e.printStackTrace();
        }

        this.mappings = mappings;
    }

    public void reload() {
        this.load();
    }

    public String getCellID(String cellName) {
        Cell cell = this.getCell(cellName);
        if (cell != null) {
            return cell.id;
        }

        this.logger.error("couldn't find cell id with name {}", cellName);
        return "AZ999";
    }

    public Cell getCell(String cellName) {
        return this.mappings.cells.get(cellName);
    }
}
