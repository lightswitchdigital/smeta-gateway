package com.lightswitch.ramdom.smeta.controllers;

import com.lightswitch.ramdom.smeta.PricelistMappings;
import com.lightswitch.ramdom.smeta.WorkbooksPool;
import com.lightswitch.ramdom.smeta.mappings.pricelist.Cell;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

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

    @GetMapping("/api/v1/pricelist")
    public void testing() {
        for (Map.Entry<String, Cell> entry :
                mappings.mappings.cells.entrySet()) {
            System.out.println(entry.getKey() + " : " + entry.getValue().def);
        }
    }
}
