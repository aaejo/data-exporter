package io.github.aaejo.dataexporter;

import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;

@SpringBootApplication
public class DataExporterApplication {

	private final DataExporter dataExporter;

	public DataExporterApplication(DataExporter dataExporter) {
		this.dataExporter = dataExporter;
	}

	public static void main(String[] args) {
		SpringApplication.run(DataExporterApplication.class, args);
	}

	@Bean
    public ApplicationRunner runner() {
        return args -> {
			dataExporter.retrieveData();
        };
    }

}
