using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;

class Program
{
    static IConfigurationRoot configuration;
    static ILogger logger;

    // Метод для получения текущего курса USD к RUB
    static async Task<double> GetUsdToRubExchangeRate()
    {
        double defaultRate = 90.0;
        try
        {
            using (HttpClient client = new HttpClient())
            {
                string url = configuration["ExchangeRateApiUrl"];
                HttpResponseMessage response = await client.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    JObject data = JObject.Parse(jsonResponse);
                    double usdToRub = (double)data["rates"]["RUB"];
                    return usdToRub;
                }
                else
                {
                    logger.LogWarning($"Не удалось получить курс обмена. Код состояния: {response.StatusCode}");
                }
            }
        }
        catch (Exception e)
        {
            logger.LogError($"Ошибка при получении курса обмена: {e.Message}");
        }
        return defaultRate;
    }

    static async Task Main(string[] args)
    {
        // Создаем построитель конфигурации
        var builder = new ConfigurationBuilder()
            .SetBasePath(Directory.GetCurrentDirectory())
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

        // Считываем конфигурацию из файла
        configuration = builder.Build();

        // Настраиваем логирование
        var loggerFactory = LoggerFactory.Create(builder =>
        {
            builder.AddConsole(); // Добавляем консольный логгер
        });

        logger = loggerFactory.CreateLogger<Program>(); // Создаем логгер для текущего класса Program

        // Читаем URL для скачивания и пути к файлам из конфигурации
        string downloadUrl = configuration["DownloadUrl"];
        string downloadedFilePath = @"downloaded_file.xlsx";
        string projectFilePath = configuration["ProjectFilePath"];

        // Попытка скачивания файла
        try
        {
            using (WebClient webClient = new WebClient())
            {
                webClient.DownloadFile(downloadUrl, downloadedFilePath);
                logger.LogInformation("Файл успешно скачан.");
            }
        }
        catch (Exception ex)
        {
            logger.LogError($"Ошибка при скачивании файла по ссылке {downloadUrl}: {ex.Message}");
            return;
        }

        // Проверяем, существует ли скачанный файл
        if (!File.Exists(downloadedFilePath))
        {
            logger.LogError($"Скачанный файл не найден: {downloadedFilePath}");
            return;
        }

        // Обработка данных из скачанного файла
        try
        {
            double totalPriceRub = 0.0;
            int priceCount = 0;

            using (var workbook = new ClosedXML.Excel.XLWorkbook(downloadedFilePath))
            {
                var worksheet = workbook.Worksheet(4); // Предполагаем, что данные находятся на 4-м листе
                foreach (var row in worksheet.RowsUsed().Skip(1)) // Пропускаем заголовок
                {
                    var priceCell = row.Cell(3); // Предполагаем, что цена находится в третьем столбце
                    if (double.TryParse(priceCell.GetValue<string>().Replace(" ", ""), out double price))
                    {
                        totalPriceRub += price;
                        priceCount++;
                    }
                }
            }

            // Если найдены данные о ценах
            if (priceCount > 0)
            {
                double weightedAveragePriceRub = totalPriceRub / priceCount;
                logger.LogInformation($"Средневзвешенная цена арматуры: {weightedAveragePriceRub} руб.");

                // Расчет цены с коэффициентом
                double priceWithCoef = weightedAveragePriceRub * 0.63;

                // Получение текущего курса доллара к рублю
                double usdRub = await GetUsdToRubExchangeRate();
                logger.LogInformation($"Текущий курс доллара к рублю: {usdRub}");

                // Конвертация цены в доллары
                double priceUsd = priceWithCoef / usdRub;
                logger.LogInformation($"Цена в долларах: {priceUsd}");

                // Обработка файла проекта
                using (var projectWorkbook = new ClosedXML.Excel.XLWorkbook(projectFilePath))
                {
                    var customConstraintsSheet = projectWorkbook.Worksheet("Custom Constraints");
                    var linearExpressionsSheet = projectWorkbook.Worksheet("Linear expressions");

                    // Проверяем, что нужные листы существуют
                    if (customConstraintsSheet == null || linearExpressionsSheet == null)
                    {
                        logger.LogError("Листы 'Custom Constraints' или 'Linear expressions' не найдены в файле проекта.");
                        return;
                    }

                    // Находим строку с меткой Slab и обновляем значения
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

                    // Сохраняем измененный файл проекта
                    projectWorkbook.SaveAs("Updated_" + projectFilePath);
                }

                logger.LogInformation("Данные успешно обновлены в файле проекта.");
            }
            else
            {
                logger.LogWarning("Не удалось найти данные о цене арматуры в скачанном файле.");
            }
        }
        catch (Exception ex)
        {
            logger.LogError($"Ошибка при обработке файлов: {ex.Message}");
        }

        logger.LogInformation("Нажмите любую клавишу для выхода...");
        Console.ReadKey();
    }
}
