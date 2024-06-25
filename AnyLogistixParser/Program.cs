using System.Net;
using ClosedXML.Excel;


        string downloadUrl = "https://m-invest.ru/upload/iblock/646/m9gjk826078xpcl08pzdcjw26003dzek.xlsx";
        string downloadedFilePath = @"downloaded_file.xlsx";
        string projectFilePath = @"Scenario from Template 16_12_43 test.xlsx";

        // Скачиваем файл по указанной ссылке
        try
        {
            using (WebClient webClient = new WebClient())
            {
                webClient.DownloadFile(downloadUrl, downloadedFilePath);
                Console.WriteLine("Файл успешно скачан.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка при скачивании файла по ссылке {downloadUrl}: {ex.Message}");
            return;
        }

        // Проверяем, существует ли скачанный файл
        if (!File.Exists(downloadedFilePath))
        {
            Console.WriteLine("Скачанный файл не найден: " + downloadedFilePath);
            return;
        }

        // Проверяем, существует ли файл проекта
        if (!File.Exists(projectFilePath))
        {
            Console.WriteLine("Файл проекта не найден: " + projectFilePath);
            return;
        }

        // Обработка скачанного файла
        try
        {
            using (var workbook = new XLWorkbook(downloadedFilePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    // Пропускаем лист "Сортовый прокат"
                    if (worksheet.Name == "Сортовый прокат")
                    {
                        continue;
                    }

                    Console.WriteLine($"Поиск в листе '{worksheet.Name}':");

                    foreach (var row in worksheet.RowsUsed())
                    {
                        bool containsKeyword = false;

                        foreach (var cell in row.Cells())
                        {
                            if (cell.GetValue<string>().Contains("Арматура"))
                            {
                                containsKeyword = true;
                                break;
                            }
                        }

                        if (containsKeyword)
                        {
                            foreach (var cell in row.Cells())
                            {
                                Console.Write(cell.GetValue<string>() + "\t");
                            }
                            Console.WriteLine();
                        }
                    }

                    Console.WriteLine(); // Добавляем пустую строку для отделения выводов
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ошибка при чтении скаченного файла Excel: " + ex.Message);
        }

        // Обработка файла проекта
        try
        {
            using (var workbook = new XLWorkbook(projectFilePath))
            {
                // Получаем объекты Worksheet для листов "Custom Constraints" и "Linear expressions"
                var customConstraintsSheet = workbook.Worksheet("Custom Constraints");
                var linearExpressionsSheet = workbook.Worksheet("Linear expressions");

                // Проверяем, что листы существуют
                if (customConstraintsSheet == null || linearExpressionsSheet == null)
                {
                    Console.WriteLine("Листы 'Custom Constraints' или 'Linear expressions' не найдены в файле проекта.");
                    return;
                }

                // Выводим данные из листа "Custom Constraints"
                Console.WriteLine($"Содержимое листа 'Custom Constraints':");
                foreach (var row in customConstraintsSheet.RowsUsed())
                {
                    foreach (var cell in row.Cells())
                    {
                        Console.Write(cell.GetValue<string>() + "\t");
                    }
                    Console.WriteLine();
                }

                Console.WriteLine(); // Добавляем пустую строку для отделения выводов

                // Выводим данные из листа "Linear expressions"
                Console.WriteLine($"Содержимое листа 'Linear expressions':");
                foreach (var row in linearExpressionsSheet.RowsUsed())
                {
                    foreach (var cell in row.Cells())
                    {
                        Console.Write(cell.GetValue<string>() + "\t");
                    }
                    Console.WriteLine();
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ошибка при чтении файла проекта Excel: " + ex.Message);
        }

        Console.WriteLine("Нажмите любую клавишу для выхода...");
        Console.ReadKey();
