package io.github.aaejo.dataexporter;

import java.util.Optional;

import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Profile;

@SpringBootApplication
public class DataExporterApplication {

	public static void main(String[] args) {
		SpringApplication.run(DataExporterApplication.class, args);
	}

	@Bean
	@Profile ("console")
    public ApplicationRunner runner(DataExporter dataExporter) {
        return args -> {
			dataExporter.retrieveData(Optional.of("DIA_CUP"));
        };
    }

}
