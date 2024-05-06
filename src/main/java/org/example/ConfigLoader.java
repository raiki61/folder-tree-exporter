package org.example;

import org.yaml.snakeyaml.Yaml;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ConfigLoader {

    public static Config loadConfig(String configFile) throws IOException {
        Yaml yaml = new Yaml();
        try (FileInputStream inputStream = new FileInputStream(configFile)) {
            Map<String, Object> configData = yaml.load(inputStream);
            List<String> fileExtensionFilters;
            if (configData.get("fileExtensionFilters") == null) {
                fileExtensionFilters = new ArrayList<>();
            } else {
                fileExtensionFilters = (List<String>) configData.get("fileExtensionFilters");
            }

            List<String> excludeDirectoryPatterns;
            if (configData.get("excludeDirectoryPatterns") == null) {
                excludeDirectoryPatterns = new ArrayList<>();
            } else {
                excludeDirectoryPatterns = (List<String>) configData.get("excludeDirectoryPatterns");
            }

            List<String> excludeFilePatterns;
            if (configData.get("excludeFilePatterns") == null) {
                excludeFilePatterns = new ArrayList<>();
            } else {
                excludeFilePatterns = (List<String>) configData.get("excludeFilePatterns");
            }
            return new Config(fileExtensionFilters, excludeDirectoryPatterns, excludeFilePatterns);
        }
    }

    public static class Config {
        private final List<String> fileExtensionFilters;
        private final List<String> excludeDirectoryPatterns;
        private final List<String> excludeFilePatterns;

        public Config(List<String> fileExtensionFilters, List<String> excludeDirectoryPatterns, List<String> excludeFilePatterns) {
            this.fileExtensionFilters = fileExtensionFilters;
            this.excludeDirectoryPatterns = excludeDirectoryPatterns;
            this.excludeFilePatterns = excludeFilePatterns;
        }

        public List<String> getExcludeDirectoryPatterns() {
            return excludeDirectoryPatterns;
        }

        public List<String> getExcludeFilePatterns() {
            return excludeFilePatterns;
        }

        public List<String> getFileExtensionFilters() {
            return fileExtensionFilters;
        }
    }
}
