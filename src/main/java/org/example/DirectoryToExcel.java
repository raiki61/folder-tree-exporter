package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.net.URI;
import java.nio.file.Files;
import java.util.List;

public class DirectoryToExcel {

    private PrintStream logStream;

    private ConfigLoader.Config config;

    public DirectoryToExcel(String yamlConfigPath, String logFileDir) throws IOException {
        try {
            File logFile = new File(logFileDir + File.separator + "directory_tree_log.txt");
            this.logStream = new PrintStream(Files.newOutputStream(logFile.toPath()));
            System.setOut(this.logStream);

            if (new File(yamlConfigPath).exists()) {
                this.config = ConfigLoader.loadConfig(yamlConfigPath);
            }
        } catch (IOException e) {
            System.err.println("ファイル操作中にエラーが発生しました: " + e.getMessage());
        }
    }

    public static void main(String[] args) throws IOException {
        String currentDirectoryPath = System.getProperty("user.dir");
        String configFile = "folder-tree-exporter-config.yml"; // 外部設定ファイル

        // エクセルファイルにディレクトリのツリー構造を出力する
        new DirectoryToExcel(currentDirectoryPath + File.separator + configFile, currentDirectoryPath).writeDirectoryTreeToExcel(currentDirectoryPath, currentDirectoryPath);
    }

    private static String getRelativePath(String baseDir, String filePath) {
        File base = new File(baseDir);
        File file = new File(filePath);

        try {
            URI baseUri = base.toURI();
            URI fileUri = file.toURI();
            return baseUri.relativize(fileUri).getPath();
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }
    }

    private static boolean shouldExclude(File file, List<String> excludeDirectoryPatterns, List<String> excludeFilePatterns) {
        for (String excludeDirectoryPattern : excludeDirectoryPatterns) {
            if (file.isDirectory() && !excludeDirectoryPattern.isEmpty() && file.getAbsolutePath().matches(excludeDirectoryPattern)) {
                return true;
            }
        }
        for (String excludeFilePattern : excludeFilePatterns) {
            if (file.isFile() && !excludeFilePattern.isEmpty() && file.getName().matches(excludeFilePattern)) {
                return true;
            }
        }
        return false;
    }

    @Override
    protected void finalize() throws Throwable {
        // ログファイルをクローズする
        if (this.logStream != null) {
            this.logStream.close();
        }
        super.finalize();
    }

    public void writeDirectoryTreeToExcel(String directoryPath, String outputDir) {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Directory Tree");

            // ディレクトリのツリー構造を再帰的にエクセルに書き込む
            writeDirectoryToExcel(new File(directoryPath), sheet, 0, 0, directoryPath, workbook);

            // エクセルファイルに書き込む
            try (FileOutputStream outputStream = new FileOutputStream(outputDir + File.separator + "directory_tree.xlsx")) {
                workbook.write(outputStream);
                System.out.println("ディレクトリのツリー構造がエクセルに書き込まれました: " + outputDir + File.separator + "directory_tree.xlsx");
            } catch (IOException e) {
                System.err.println("エクセルファイルの書き込み中にエラーが発生しました: " + e.getMessage());
            }
        } catch (IOException e) {
            System.err.println("エクセルファイルの作成中にエラーが発生しました: " + e.getMessage());
        } finally {
            logStream.close();
        }
    }

    private int writeDirectoryToExcel(File directory, Sheet sheet, int rowNum, int colNum, String baseDir, Workbook workbook) {
        if (!directory.isDirectory()) {
            return rowNum;
        }

        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(colNum);
        cell.setCellValue(directory.getName());

        for (File file : directory.listFiles()) {

            if (config != null && shouldExclude(file, config.getExcludeDirectoryPatterns(), config.getExcludeFilePatterns())) {
                continue;
            }

            if (file.isDirectory()) {

                rowNum = writeDirectoryToExcel(file, sheet, rowNum, colNum + 1, baseDir, workbook);
            } else {

                Row fileRow = sheet.createRow(rowNum++);
                Cell fileCell = fileRow.createCell(colNum + 1);
                fileCell.setCellValue(file.getName());

                // ファイルの相対パスをURIとして解釈可能な形式に変換してからリンクとして追加する
                String relativePath = getRelativePath(baseDir, file.getAbsolutePath());
                String linkFormula = "HYPERLINK(\"" + relativePath + "\",\"" + file.getName() + "\")";
                fileCell.setCellFormula(linkFormula);

                // リンクのテキストを青色にし、下線を追加する
                CellStyle hyperlinkStyle = workbook.createCellStyle();
                Font font = workbook.createFont();
                font.setColor(IndexedColors.BLUE.getIndex());
                font.setUnderline(Font.U_SINGLE);
                hyperlinkStyle.setFont(font);
                fileCell.setCellStyle(hyperlinkStyle);
            }
        }

        return rowNum;
    }
}
