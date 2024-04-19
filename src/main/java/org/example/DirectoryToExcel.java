package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.net.URI;

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
        } catch (IOException e) {
            System.err.println("ログファイルの作成中にエラーが発生しました: " + e.getMessage());
        }
    }

    private static int writeDirectoryToExcel(File directory, Sheet sheet, int rowNum, int colNum, String baseDir, Workbook workbook) {
        if (!directory.isDirectory()) {
            return rowNum;
        }

        Row row = sheet.createRow(rowNum++);
        Cell cell = row.createCell(colNum);
        cell.setCellValue(directory.getName());

        for (File file : directory.listFiles()) {
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
}
