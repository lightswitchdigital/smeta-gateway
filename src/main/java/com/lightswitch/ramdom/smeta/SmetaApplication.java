package com.lightswitch.ramdom.smeta;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication(scanBasePackages = {"com.lightswitch.ramdom.smeta"})
public class SmetaApplication {

	public static void main(String[] args) {
		SpringApplication.run(SmetaApplication.class, args);
	}
}
