using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Text;

class Program
{
    static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.InputEncoding = System.Text.Encoding.UTF8;

        while (true)
        {
            Console.WriteLine("Выберите действие:");
            Console.WriteLine("1 - Экспорт из Excel в текстовый файл");
            Console.WriteLine("2 - Чтение текстового файла");
            Console.WriteLine("3 - Выход");

            var choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    ExportExcelToText();
                    break;
                case "2":
                    ReadTextFile();
                    break;
                case "3":
                    return;
                default:
                    Console.WriteLine("Неверный выбор. Попробуйте снова.");
                    break;
            }
        }
    }

    static void ExportExcelToText()
    {
        try
        {
            Console.WriteLine("Введите путь к файлу Excel:");
            string excelPath = Console.ReadLine();

            Console.WriteLine("Введите имя листа (или оставьте пустым для первого листа):");
            string sheetName = Console.ReadLine();

            Console.WriteLine("Введите номер столбца для экспорта (начиная с 1):");
            int columnNumber = int.Parse(Console.ReadLine());

            Console.WriteLine("Введите путь для сохранения текстового файла:");
            string textFilePath = Console.ReadLine();

            // Проверяем существование файла
            if (File.Exists(textFilePath))
            {
                Console.WriteLine($"Файл {textFilePath} уже существует. Перезаписать? (y/n)");
                var response = Console.ReadLine().ToLower();

                if (response != "y" && response != "н") // 'н' для русской раскладки
                {
                    // Генерируем уникальное имя файла
                    string directory = Path.GetDirectoryName(textFilePath);
                    string fileName = Path.GetFileNameWithoutExtension(textFilePath);
                    string extension = Path.GetExtension(textFilePath);

                    int counter = 1;
                    string newFilePath;
                    do
                    {
                        newFilePath = Path.Combine(directory, $"{fileName}_{counter}{extension}");
                        counter++;
                    } while (File.Exists(newFilePath));

                    Console.WriteLine($"Создаем новый файл: {newFilePath}");
                    textFilePath = newFilePath;
                }
            }

            // Установка лицензионного контекста
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var worksheet = string.IsNullOrEmpty(sheetName)
                    ? package.Workbook.Worksheets[0]
                    : package.Workbook.Worksheets[sheetName];

                if (worksheet == null)
                {
                    Console.WriteLine("Лист не найден.");
                    return;
                }

                // Используем FileStream с UTF-8 кодировкой
                using (FileStream fs = new FileStream(textFilePath, FileMode.OpenOrCreate))
                {
                    using (StreamWriter writer = new StreamWriter(fs, Encoding.UTF8))
                    {
                        int rowCount = worksheet.Dimension.Rows;
                        for (int row = 1; row <= rowCount; row++)
                        {
                            string cellValue = worksheet.Cells[row, columnNumber].Text;
                            writer.WriteLine(cellValue);
                        }
                    }
                }
                Console.WriteLine($"Данные успешно экспортированы в {textFilePath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка: {ex.Message}");
        }
    }

    static void ReadTextFile()
    {
        try
        {
            Console.WriteLine("Введите путь к текстовому файлу:");
            string filePath = Console.ReadLine();

            if (!File.Exists(filePath))
            {
                Console.WriteLine("Файл не найден.");
                return;
            }

            Console.WriteLine("Содержимое файла:");
            Console.WriteLine("----------------");

            string[] lines = File.ReadAllLines(filePath);
            foreach (string line in lines)
            {
                Console.WriteLine(line);
            }

            Console.WriteLine("----------------");
            Console.WriteLine($"Всего строк: {lines.Length}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка: {ex.Message}");
        }
    }
}