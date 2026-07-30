package com.example;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xslf.extractor.XSLFExtractor;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.Loader;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;

public class TextFileCollector {

    public static void main(String[] args) {
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(System.in))) {

            // 1. Запрашиваем путь к папке с файлами
            System.out.println("Введите путь к папке для поиска файлов (.docx, .pdf, .txt, .pptx, .xlsx, .xls):");
            String sourcePath = reader.readLine();

            Path sourceDir = Paths.get(sourcePath);

            if (!Files.exists(sourceDir) || !Files.isDirectory(sourceDir)) {
                System.err.println("Ошибка: Указанный путь не существует или не является папкой.");
                return;
            }

            // 2. Запрашиваем имя для выходного файла (будет сохранен на D:)
            System.out.println("Введите имя выходного файла (по умолчанию сохраняется на диск D), например, result.java:");
            String outputFileName = reader.readLine();

            // Формируем полный путь на диске D
            Path outputFile = Paths.get("D:\\" + outputFileName);

            // Проверяем, существует ли диск D:
            if (!Files.exists(Paths.get("D:\\"))) {
                System.err.println("Ошибка: Диск D: не найден!");
                return;
            }

            // 3. Запускаем процесс сбора и записи данных
            try (BufferedWriter writer = Files.newBufferedWriter(outputFile)) {
                System.out.println("Поиск и обработка файлов...");

                // Добавляем заголовок как в Java файле
                writer.write("// Собранные тексты из файлов (.docx, .pdf, .txt, .pptx, .xlsx, .xls)");
                writer.newLine();
                writer.write("// Дата создания: " + java.time.LocalDate.now());
                writer.newLine();
                writer.write("// Программа: TextFileCollector");
                writer.newLine();
                writer.write("// ========================================");
                writer.newLine();
                writer.newLine();

                collectTextFromFiles(sourceDir, writer);

                System.out.println("Готово! Результат сохранен в файл: " + outputFile.toAbsolutePath());
                System.out.println("Всего обработано файлов: " + fileCounter);
                System.out.println("Из них:");
                System.out.println("  - DOCX: " + docxCounter);
                System.out.println("  - PDF: " + pdfCounter);
                System.out.println("  - TXT: " + txtCounter);
                System.out.println("  - PPTX: " + pptxCounter);
                System.out.println("  - Excel (XLSX/XLS): " + excelCounter);

                // Показываем статистику ошибок
                if (errorCounter > 0) {
                    System.out.println("Файлов с ошибками: " + errorCounter);
                }

            } catch (IOException e) {
                System.err.println("Ошибка при записи в выходной файл: " + e.getMessage());
            }

        } catch (IOException e) {
            System.err.println("Ошибка ввода/вывода: " + e.getMessage());
        }
    }

    // Счетчики для статистики
    private static int fileCounter = 0;
    private static int errorCounter = 0;
    private static int docxCounter = 0;
    private static int pdfCounter = 0;
    private static int txtCounter = 0;
    private static int pptxCounter = 0;
    private static int excelCounter = 0;

    private static void collectTextFromFiles(Path rootDir, BufferedWriter writer) throws IOException {
        Files.walkFileTree(rootDir, new SimpleFileVisitor<>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                if (!attrs.isRegularFile()) {
                    return FileVisitResult.CONTINUE;
                }

                String fileName = file.toString().toLowerCase();

                try {
                    if (fileName.endsWith(".docx")) {
                        processDocxFile(file, rootDir, writer);
                        docxCounter++;
                        fileCounter++;
                    } else if (fileName.endsWith(".pdf")) {
                        processPdfFile(file, rootDir, writer);
                        pdfCounter++;
                        fileCounter++;
                    } else if (fileName.endsWith(".txt")) {
                        processTxtFile(file, rootDir, writer);
                        txtCounter++;
                        fileCounter++;
                    } else if (fileName.endsWith(".pptx")) {
                        processPptxFile(file, rootDir, writer);
                        pptxCounter++;
                        fileCounter++;
                    } else if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
                        processExcelFile(file, rootDir, writer);
                        excelCounter++;
                        fileCounter++;
                    }
                } catch (Exception e) {
                    errorCounter++;
                    System.err.println("Ошибка при обработке файла: " + file);
                    System.err.println("Причина: " + e.getMessage());
                }

                return FileVisitResult.CONTINUE;
            }

            @Override
            public FileVisitResult visitFileFailed(Path file, IOException exc) {
                System.err.println("Не удалось прочитать: " + file);
                return FileVisitResult.CONTINUE;
            }
        });
    }

    private static void processDocxFile(Path file, Path rootDir, BufferedWriter writer) throws IOException {
        Path relativePath = rootDir.relativize(file.getParent());
        String address = relativePath.toString();
        if (address.isEmpty()) {
            address = ".";
        }

        String fileName = file.getFileName().toString();
        String content = extractTextFromDocx(file);

        writeFileHeader(writer, address, fileName, "DOCX");
        writeContent(writer, content);

        System.out.println("✓ Обработан DOCX: " + fileName + " (" + address + ")");
    }

    private static void processPdfFile(Path file, Path rootDir, BufferedWriter writer) throws IOException {
        Path relativePath = rootDir.relativize(file.getParent());
        String address = relativePath.toString();
        if (address.isEmpty()) {
            address = ".";
        }

        String fileName = file.getFileName().toString();
        String content = extractTextFromPdf(file);

        writeFileHeader(writer, address, fileName, "PDF");
        writeContent(writer, content);

        System.out.println("✓ Обработан PDF: " + fileName + " (" + address + ")");
    }

    private static void processTxtFile(Path file, Path rootDir, BufferedWriter writer) throws IOException {
        Path relativePath = rootDir.relativize(file.getParent());
        String address = relativePath.toString();
        if (address.isEmpty()) {
            address = ".";
        }

        String fileName = file.getFileName().toString();
        String content = extractTextFromTxt(file);

        writeFileHeader(writer, address, fileName, "TXT");
        writeContent(writer, content);

        System.out.println("✓ Обработан TXT: " + fileName + " (" + address + ")");
    }

    private static void processPptxFile(Path file, Path rootDir, BufferedWriter writer) throws IOException {
        Path relativePath = rootDir.relativize(file.getParent());
        String address = relativePath.toString();
        if (address.isEmpty()) {
            address = ".";
        }

        String fileName = file.getFileName().toString();
        String content = extractTextFromPptx(file);

        writeFileHeader(writer, address, fileName, "PPTX");
        writeContent(writer, content);

        System.out.println("✓ Обработан PPTX: " + fileName + " (" + address + ")");
    }

    // НОВЫЙ МЕТОД для обработки Excel файлов
    private static void processExcelFile(Path file, Path rootDir, BufferedWriter writer) throws IOException {
        Path relativePath = rootDir.relativize(file.getParent());
        String address = relativePath.toString();
        if (address.isEmpty()) {
            address = ".";
        }

        String fileName = file.getFileName().toString();
        String content = extractTextFromExcel(file);

        writeFileHeader(writer, address, fileName, "EXCEL");
        writeContent(writer, content);

        System.out.println("✓ Обработан Excel: " + fileName + " (" + address + ")");
    }

    private static void writeFileHeader(BufferedWriter writer, String address, String fileName, String fileType) throws IOException {
        writer.write("/*");
        writer.newLine();
        writer.write(" * Тип файла: " + fileType);
        writer.newLine();
        writer.write(" * Адрес: " + address);
        writer.newLine();
        writer.write(" * Название файла: " + fileName);
        writer.newLine();
        writer.write(" */");
        writer.newLine();
        writer.write("// Содержание:");
        writer.newLine();
    }

    private static void writeContent(BufferedWriter writer, String content) throws IOException {
        String[] lines = content.split("\\r?\\n");
        for (String line : lines) {
            writer.write("// " + line);
            writer.newLine();
        }

        writer.newLine();
        writer.write("// ----------------------------------------");
        writer.newLine();
        writer.newLine();
    }

    private static String extractTextFromDocx(Path filePath) {
        try (InputStream fis = Files.newInputStream(filePath);
             XWPFDocument document = new XWPFDocument(fis);
             XWPFWordExtractor extractor = new XWPFWordExtractor(document)) {

            String text = extractor.getText();
            if (text == null || text.trim().isEmpty()) {
                return "[Файл не содержит текста]";
            }
            return text.trim();

        } catch (Exception e) {
            return "[ОШИБКА: Не удалось извлечь текст из DOCX. " +
                    "Ошибка: " + e.getMessage() + "]";
        }
    }

    private static String extractTextFromPdf(Path filePath) {
        try (PDDocument document = Loader.loadPDF(filePath.toFile())) {

            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(document);

            if (text == null || text.trim().isEmpty()) {
                return "[PDF файл не содержит текста (возможно, это отсканированный документ)]";
            }
            return text.trim();

        } catch (Exception e) {
            return "[ОШИБКА: Не удалось извлечь текст из PDF. " +
                    "Ошибка: " + e.getMessage() + "]";
        }
    }

    private static String extractTextFromTxt(Path filePath) {
        try {
            byte[] bytes = Files.readAllBytes(filePath);
            String text = new String(bytes, "UTF-8");

            if (text == null || text.trim().isEmpty() || text.trim().length() < 2) {
                text = new String(bytes, "Windows-1251");
            }

            if (text == null || text.trim().isEmpty()) {
                return "[TXT файл пуст]";
            }
            return text.trim();

        } catch (Exception e) {
            return "[ОШИБКА: Не удалось прочитать TXT файл. " +
                    "Ошибка: " + e.getMessage() + "]";
        }
    }

    private static String extractTextFromPptx(Path filePath) {
        try (InputStream fis = Files.newInputStream(filePath);
             XMLSlideShow ppt = new XMLSlideShow(fis);
             XSLFExtractor extractor = new XSLFExtractor(ppt)) {

            String text = extractor.getText();
            if (text == null || text.trim().isEmpty()) {
                return "[Презентация не содержит текста]";
            }

            // Добавляем информацию о слайдах для лучшей читаемости
            StringBuilder formattedText = new StringBuilder();
            formattedText.append("=== Презентация ===\n");
            formattedText.append("Количество слайдов: ").append(ppt.getSlides().size()).append("\n\n");
            formattedText.append(text.trim());

            return formattedText.toString();

        } catch (Exception e) {
            return "[ОШИБКА: Не удалось извлечь текст из PPTX. " +
                    "Ошибка: " + e.getMessage() + "]";
        }
    }

    // НОВЫЙ МЕТОД для извлечения текста из Excel файлов
    private static String extractTextFromExcel(Path filePath) {
        StringBuilder result = new StringBuilder();

        try (InputStream fis = Files.newInputStream(filePath);
             Workbook workbook = filePath.toString().toLowerCase().endsWith(".xlsx")
                     ? new XSSFWorkbook(fis)
                     : new HSSFWorkbook(fis)) {

            int numberOfSheets = workbook.getNumberOfSheets();
            result.append("=== Excel файл ===\n");
            result.append("Количество листов: ").append(numberOfSheets).append("\n\n");

            // Проходим по всем листам
            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(i);
                result.append("--- Лист ").append(i + 1).append(": ");
                result.append(sheet.getSheetName()).append(" ---\n");

                // Проходим по всем строкам
                boolean hasData = false;
                for (Row row : sheet) {
                    StringBuilder rowText = new StringBuilder();
                    boolean rowHasData = false;

                    // Проходим по всем ячейкам в строке
                    for (Cell cell : row) {
                        String cellValue = getCellValue(cell);
                        if (cellValue != null && !cellValue.trim().isEmpty()) {
                            if (rowHasData) {
                                rowText.append(" | ");
                            }
                            rowText.append(cellValue);
                            rowHasData = true;
                            hasData = true;
                        }
                    }

                    if (rowHasData) {
                        result.append("  Ряд ").append(row.getRowNum() + 1).append(": ");
                        result.append(rowText.toString()).append("\n");
                    }
                }

                if (!hasData) {
                    result.append("  [Лист пуст]\n");
                }
                result.append("\n");
            }

            if (result.toString().contains("=== Excel файл ===") &&
                    !result.toString().contains("[Лист пуст]") &&
                    !result.toString().contains("Ряд")) {
                return "[Excel файл не содержит данных]";
            }

            return result.toString();

        } catch (Exception e) {
            return "[ОШИБКА: Не удалось извлечь текст из Excel. " +
                    "Ошибка: " + e.getMessage() + "]";
        }
    }

    // Вспомогательный метод для получения значения ячейки
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    double value = cell.getNumericCellValue();
                    if (value == (long) value) {
                        return String.valueOf((long) value);
                    } else {
                        return String.valueOf(value);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (IllegalStateException e) {
                    return "[ФОРМУЛА: " + cell.getCellFormula() + "]";
                }
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}