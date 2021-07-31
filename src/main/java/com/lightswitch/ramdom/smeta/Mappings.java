package com.lightswitch.ramdom.smeta;

import com.fasterxml.jackson.core.JsonGenerationException;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.lightswitch.ramdom.smeta.mappings.Cells;
import org.springframework.stereotype.Repository;

import java.io.File;
import java.io.IOException;

@Repository
public class Mappings {

    public Cells mappings;

    public void load() {
        ObjectMapper objectMapper = new ObjectMapper();
        String jsonPath = System.getProperty("user.dir") + "/src/mappings.json";
        Cells mappings = null;

        try {
            mappings = objectMapper.readValue(new File(jsonPath), Cells.class);
        } catch (JsonGenerationException e) {
            System.out.println("json generation exception");
            e.printStackTrace();
        } catch (JsonMappingException e) {
            System.out.println("json mappings exception");
            e.printStackTrace();
        } catch (IOException e) {
            System.out.println("io exception");
            e.printStackTrace();
        }

        this.mappings = mappings;
    }

    public void reload() {
        this.load();
    }
}
