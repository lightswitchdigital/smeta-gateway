package com.lightswitch.ramdom.smeta;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ImportResource;

@SpringBootApplication(scanBasePackages = {"com.lightswitch.ramdom.smeta"})
@ImportResource("classpath:beans.xml")
public class SmetaApplication {

	public static void main(String[] args) {
		SpringApplication.run(SmetaApplication.class, args);
	}
}
