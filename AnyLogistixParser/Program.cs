using System;
using System.IO;
using System.Linq;
using System.Net;
using ClosedXML.Excel;

class Program
{
    static void Main()
    {
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
            double totalPriceRub = 0.0;
            int priceCount = 0;

            using (var workbook = new XLWorkbook(downloadedFilePath))
            {
                var worksheet = workbook.Worksheet(4); // Предполагаем, что данные находятся в четвертом листе
                foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропускаем заголовок
                {
                    var priceCell = row.Cell(3); // Предполагаем, что цена в третьем столбце
                    if (double.TryParse(priceCell.GetValue<string>().Replace(" ", ""), out double price))
                    {
                        totalPriceRub += price;
                        priceCount++;
                    }
                }
            }

            if (priceCount > 0)
            {
                double weightedAveragePriceRub = totalPriceRub / priceCount;
                Console.WriteLine($"Средневзвешенная цена арматуры: {weightedAveragePriceRub} руб.");

                // Умножаем на коэффициент
                double priceWithCoef = weightedAveragePriceRub * 0.63;

                // Получение текущего курса доллара (замените на ваш код получения курса доллара)
                double usdRub = 75.0; // Пример значения курса доллара

                // Конвертация цены в доллары
                double priceUsd = priceWithCoef / usdRub;
                Console.WriteLine($"Цена в долларах: {priceUsd}");

                // Обработка файла проекта
                using (var workbook = new XLWorkbook(projectFilePath))
                {
                    var customConstraintsSheet = workbook.Worksheet("Custom Constraints");
                    var linearExpressionsSheet = workbook.Worksheet("Linear expressions");

                    // Проверяем, что листы существуют
                    if (customConstraintsSheet == null || linearExpressionsSheet == null)
                    {
                        Console.WriteLine("Листы 'Custom Constraints' или 'Linear expressions' не найдены в файле проекта.");
                        return;
                    }

                    // Находим строку с меткой Slab и подставляем значения
                    foreach (var row in customConstraintsSheet.RowsUsed())
                    {
                        foreach (var cell in row.Cells())
                        {
                            if (cell.GetValue<string>().Contains("Slab"))
                            {
                                var linearExpressionNum = cell.GetValue<string>().Split().Last();
                                foreach (var exprRow in linearExpressionsSheet.RowsUsed())
                                {
                                    if (exprRow.FirstCell().GetValue<string>().Contains(linearExpressionNum))
                                    {
                                        foreach (var exprCell in exprRow.Cells())
                                        {
                                            if (exprCell.GetValue<string>().Contains("[·o52]"))
                                            {
                                                exprCell.Value = priceWithCoef;
                                            }
                                            else if (exprCell.GetValue<string>().Contains("[·o53]"))
                                            {
                                                exprCell.Value = priceUsd;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Сохраняем файл
                    workbook.SaveAs("Updated_" + projectFilePath);
                }

                Console.WriteLine("Данные успешно обновлены в файле проекта.");
            }
            else
            {
                Console.WriteLine("Не удалось найти данные о цене арматуры в скачанном файле.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Ошибка при обработке файлов: " + ex.Message);
        }

        Console.WriteLine("Нажмите любую клавишу для выхода...");
        Console.ReadKey();
    }
}
