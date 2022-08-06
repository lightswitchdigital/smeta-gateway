package com.lightswitch.ramdom.smeta.controllers;

import com.lightswitch.ramdom.smeta.SmetaMappings;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.PostConstruct;

@RestController
public class ReloadController {

    Logger logger = LoggerFactory.getLogger(ReloadController.class);

    @Autowired
    SmetaMappings mappings;

    @PostConstruct
    @GetMapping("/reload")
    public void reload() {
        logger.info("reloading mappings");
        mappings.reload();
    }
}
