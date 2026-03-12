package com.example;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.nio.file.*;
import java.nio.file.attribute.BasicFileAttributes;

public class TextFileCollector {

    public static void main(String[] args) {
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(System.in))) {

            // 1. Запрашиваем путь к папке с Word файлами
            System.out.println("Введите путь к папке для поиска .docx файлов:");
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
                System.out.println("Поиск и обработка Word файлов...");

                // Добавляем заголовок как в Java файле
                writer.write("// Собранные тексты из Word файлов");
                writer.newLine();
                writer.write("// Дата создания: " + java.time.LocalDate.now());
                writer.newLine();
                writer.write("// Программа: TextFileCollector");
                writer.newLine();
                writer.write("// ========================================");
                writer.newLine();
                writer.newLine();

                collectTextFromDocxFiles(sourceDir, writer);

                System.out.println("Готово! Результат сохранен в файл: " + outputFile.toAbsolutePath());
                System.out.println("Всего обработано файлов: " + fileCounter);

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

    private static void collectTextFromDocxFiles(Path rootDir, BufferedWriter writer) throws IOException {
        Files.walkFileTree(rootDir, new SimpleFileVisitor<>() {
            @Override
            public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) {
                // Проверяем, что это файл и имеет расширение .docx
                if (attrs.isRegularFile() && file.toString().toLowerCase().endsWith(".docx")) {
                    try {
                        processDocxFile(file, rootDir, writer);
                        fileCounter++;
                    } catch (Exception e) {
                        errorCounter++;
                        System.err.println("Ошибка при обработке файла: " + file);
                        System.err.println("Причина: " + e.getMessage());
                    }
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

    private static void processDocxFile(Path docxFile, Path rootDir, BufferedWriter writer) throws IOException {
        // Вычисляем относительный путь от корневой папки
        Path relativePath = rootDir.relativize(docxFile.getParent());
        String address = relativePath.toString();
        if (address.isEmpty()) {
            address = "."; // Если файл в корневой папке
        }

        String fileName = docxFile.getFileName().toString();

        // Извлекаем текст из Word файла
        String content = extractTextFromDocx(docxFile);

        // Записываем в выходной файл в формате Java комментария
        writer.write("/*");
        writer.newLine();
        writer.write(" * Адрес: " + address);
        writer.newLine();
        writer.write(" * Название файла: " + fileName);
        writer.newLine();
        writer.write(" */");
        writer.newLine();
        writer.write("// Содержание:");
        writer.newLine();

        // Разбиваем содержимое на строки и добавляем // перед каждой
        String[] lines = content.split("\\r?\\n");
        for (String line : lines) {
            writer.write("// " + line);
            writer.newLine();
        }

        writer.newLine();
        writer.write("// ----------------------------------------");
        writer.newLine();
        writer.newLine();

        // Показываем прогресс в консоли
        System.out.println("✓ Обработан: " + fileName + " (" + address + ")");
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
            return "[ОШИБКА: Не удалось извлечь текст. Файл может быть поврежден или это не .docx. " +
                    "Ошибка: " + e.getMessage() + "]";
        }
    }
}