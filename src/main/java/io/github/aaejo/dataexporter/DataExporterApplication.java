package io.github.aaejo.dataexporter;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.Optional;

import org.springframework.beans.factory.annotation.Value;
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
	public ApplicationRunner runner(DataExporter dataExporter,
			@Value("${aaejo.jds.data-exporter.file}") String file,
			@Value("${aaejo.jds.data-exporter.password}") Optional<String> password) {
        return args -> {
			String fileName = dataExporter.retrieveData(password);
			Files.copy(Paths.get(fileName), Paths.get(file), StandardCopyOption.REPLACE_EXISTING);
        };
    }
}
