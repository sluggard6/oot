package org.sluggard.oot;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@EnableAutoConfiguration
@MapperScan("org.sluggard.oot.dao")
public class OotApplication {

	public static void main(String[] args) {
		SpringApplication.run(OotApplication.class, args);
	}

}
