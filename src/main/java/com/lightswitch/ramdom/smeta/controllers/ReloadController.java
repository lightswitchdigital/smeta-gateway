package com.lightswitch.ramdom.smeta.controllers;

import com.lightswitch.ramdom.smeta.Mappings;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.annotation.PostConstruct;

@RestController
public class ReloadController {

    @Autowired
    Mappings mappings;

    @PostConstruct
    @GetMapping("/reload")
    public void reload() {
        System.out.println("reloading mappings");
        mappings.reload();
    }
}
