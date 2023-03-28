package io.github.aaejo.dataexporter;



import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Optional;

import org.apache.commons.compress.utils.IOUtils;
import org.springframework.context.annotation.Profile;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

import lombok.extern.slf4j.Slf4j;

@Slf4j
@RestController
@Profile("default")

public class FileDownloadController {

    private final DataExporter dataExporter;

    public FileDownloadController(DataExporter dataExporter) {
        this.dataExporter = dataExporter;
    }

    @GetMapping(
  value = "/",
  produces = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
    public @ResponseBody byte[] downloadFile(@RequestParam("password") Optional<String> password) throws Exception {
        log.info("Downloading file");
        String fileName = dataExporter.retrieveData(password);
        InputStream in = new FileInputStream(fileName);
        return IOUtils.toByteArray(in);
    }

    
}
