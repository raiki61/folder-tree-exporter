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
            List<String> fileExtensionFilters = (List<String>) configData.getOrDefault("fileExtensionFilters", new ArrayList<>());
            List<String> excludeDirectoryPatterns = (List<String>) configData.getOrDefault("excludeDirectoryPatterns", new ArrayList<>());
            List<String> excludeFilePatterns = (List<String>) configData.getOrDefault("excludeFilePatterns", new ArrayList<>());
            return new Config(fileExtensionFilters, excludeDirectoryPatterns, excludeFilePatterns);
        }
    }

    public static class Config {
        private final List<String> excludeDirectoryPatterns;
        private final List<String> excludeFilePatterns;

        public Config(List<String> fileExtensionFilters, List<String> excludeDirectoryPatterns, List<String> excludeFilePatterns) {
            this.excludeDirectoryPatterns = excludeDirectoryPatterns;
            this.excludeFilePatterns = excludeFilePatterns;
        }

        public List<String> getExcludeDirectoryPatterns() {
            return excludeDirectoryPatterns;
        }

        public List<String> getExcludeFilePatterns() {
            return excludeFilePatterns;
        }
    }
}
