package org.example;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.net.URI;
import java.net.URISyntaxException;

public class DirectoryToExcel {

    public static void main(String[] args) {
        String currentDirectoryPath = System.getProperty("user.dir");
        String directoryPath = currentDirectoryPath;
        String outputDir = currentDirectoryPath;

        // エクセルファイルにディレクトリのツリー構造を出力する
        writeDirectoryTreeToExcel(directoryPath, outputDir);
    }

    public static void writeDirectoryTreeToExcel(String directoryPath, String outputDir) {
        try {
            File logFile = new File(outputDir + File.separator + "directory_tree_log.txt");
            PrintStream logStream = new PrintStream(new FileOutputStream(logFile));
            System.setOut(logStream);

            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Directory Tree");

                // ディレクトリのツリー構造を再帰的にエクセルに書き込む
                writeDirectoryToExcel(new File(directoryPath), sheet, 0, 0);

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
        } catch (IOException e) {
            System.err.println("ログファイルの作成中にエラーが発生しました: " + e.getMessage());
        }
    }

    private static int writeDirectoryToExcel(File directory, Sheet sheet, int rowNum, int colNum) {
        if (!directory.isDirectory()) {
            return rowNum;
        }

        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(colNum);
        cell.setCellValue(directory.getName());

        for (File file : directory.listFiles()) {
            if (file.isDirectory()) {
                rowNum = writeDirectoryToExcel(file, sheet, rowNum, colNum + 1);
            } else {
                Row fileRow = sheet.createRow(rowNum++);
                Cell fileCell = fileRow.createCell(colNum + 1);
                fileCell.setCellValue(file.getName());

                // ファイルの相対パスをURIとして解釈可能な形式に変換してからリンクとして追加する
                URI uri = file.toURI();
                Hyperlink link = sheet.getWorkbook().getCreationHelper().createHyperlink(HyperlinkType.FILE);
                link.setAddress(uri.toString());
                fileCell.setHyperlink(link);
            }
        }

        return rowNum;
    }
}
