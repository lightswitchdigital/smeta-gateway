package com.lightswitch.ramdom.smeta;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.lightswitch.ramdom.smeta.mappings.Cell;
import com.lightswitch.ramdom.smeta.mappings.Cells;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Repository;

import java.io.File;
import java.io.IOException;

@Repository
public class Mappings {

    Logger logger = LoggerFactory.getLogger(Mappings.class);

    public Cells mappings;

    public void load() {
        ObjectMapper objectMapper = new ObjectMapper();
        String jsonPath = System.getProperty("user.dir") + "/src/mappings.json";
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
        return this.getCell(cellName).id;
    }

    public Cell getCell(String cellName) {
        return this.mappings.cells.get(cellName);
    }
}
