using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
namespace AnyLogistixParser
{
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

        [Obsolete("Obsolete")]
        static async Task Main(string[] args)
        {
            // Создаем построитель конфигурации
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            // Считываем конфигурацию из файла
            configuration = builder.Build();

            // Настраиваем логирование
            var loggerFactory = LoggerFactory.Create(logging =>
            {
                logging.AddConsole(); // Добавляем консольный логгер
            });

            logger = loggerFactory.CreateLogger<Program>(); // Создаем логгер для текущего класса Program

            // Читаем URL для скачивания и пути к файлам из конфигурации
            string downloadUrl = configuration["DownloadUrl"];
            string downloadedFilePath = @"downloaded_file.xlsx";
            string projectFilePath = configuration["ProjectFilePath"];

            // Читаем настройки для обработки Excel файла
            int sheetNumber = int.Parse(configuration["ExcelSettings:SheetNumber"]);
            int priceColumn = int.Parse(configuration["ExcelSettings:PriceColumn"]);
            int skipRows = int.Parse(configuration["ExcelSettings:SkipRows"]);
            string searchString = configuration["ExcelSettings:SearchString"];

            // Попытка скачивания файла
            try
            {
                using (var webClient = new System.Net.WebClient())
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

            // Проверяем, существует ли файл проекта
            if (!File.Exists(projectFilePath))
            {
                logger.LogError($"Файл проекта не найден: {projectFilePath}");
                return;
            }

            // Новый массив для хранения значений из ячеек C
            var slabValues = new List<int>();

            // Парсинг листа Custom Constraints
            try
            {
                using (var workbook = new XLWorkbook(projectFilePath))
                {
                    var customConstraintsSheet = workbook.Worksheet("Custom Constraints");

                    if (customConstraintsSheet != null)
                    {
                        foreach (var row in customConstraintsSheet.RowsUsed())
                        {
                            var cellA = row.Cell(1).GetValue<string>();
                            if (cellA.Contains("Slab [·o52]") || cellA.Contains("Slab [·o53]") || cellA.Contains("Slab [·o51]"))
                            {
                                var linearExpression = row.Cell(3).GetValue<string>();
                                var match = Regex.Match(linearExpression, @"\d+");
                                if (match.Success && int.TryParse(match.Value, out int value))
                                {
                                    slabValues.Add(value);
                                    logger.LogInformation($"Добавлено значение Linear Expression: {value}");
                                }
                            }
                        }
                    }
                    else
                    {
                        logger.LogWarning("Лист 'Custom Constraints' не найден в файле проекта.");
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Ошибка при обработке листа 'Custom Constraints': {ex.Message}");
                return;
            }

            // Логируем найденные значения
            logger.LogInformation($"Найденные значения из ячеек C: {string.Join(", ", slabValues)}");

            // Обработка данных из скачанного файла
            try
            {
                double totalPriceRub = 0.0;
                int priceCount = 0;
                bool startParsing = false;

                using (var workbook = new XLWorkbook(downloadedFilePath))
                {
                    var worksheet = workbook.Worksheet(sheetNumber); // Номер листа из конфигурации
                    foreach (var row in worksheet.RowsUsed().Skip(skipRows)) // Количество пропускаемых строк из конфигурации
                    {
                        var firstCell = row.Cell(1).GetValue<string>();
                        if (!startParsing && firstCell.Contains(searchString)) // Строка поиска из конфигурации
                        {
                            startParsing = true; // Начинаем парсинг, если нашли строку с поисковой строкой
                        }

                        if (startParsing)
                        {
                            var itemName = row.Cell(1).GetValue<string>();
                            var priceCell = row.Cell(priceColumn); // Номер столбца из конфигурации
                            if (double.TryParse(priceCell.GetValue<string>().Replace(" ", ""), out double price))
                            {
                                Console.WriteLine($"Номенклатура: {itemName}, Цена: {price} руб.");
                                totalPriceRub += price;
                                priceCount++;
                            }
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

                    // Округление до целого числа
                    int roundedPriceWithCoef = (int)Math.Round(priceWithCoef);
                    logger.LogInformation($"Округленная цена с коэффициентом: {roundedPriceWithCoef}");

                    // Получение текущего курса доллара к рублю
                    double usdRub = await GetUsdToRubExchangeRate();
                    logger.LogInformation($"Текущий курс доллара к рублю: {usdRub}");

                    // Конвертация цены в доллары
                    double priceUsd = priceWithCoef / usdRub;
                    
                    // Округление до целого числа
                    int roundedPriceUsd = (int)Math.Round(priceUsd);
                    logger.LogInformation($"Округленная цена в долларах: {roundedPriceUsd}");

                    // Обработка файла проекта
                    using (var projectWorkbook = new XLWorkbook(projectFilePath))
                    {
                        var linearExpressionsSheet = projectWorkbook.Worksheet("Linear expressions");

                        // Проверяем, что нужный лист существует
                        if (linearExpressionsSheet == null)
                        {
                            logger.LogError("Лист 'Linear expressions' не найден в файле проекта.");
                            return;
                        }

                        // Находим строку с меткой "Slab" в столбце D и обновляем значения в столбце C, если в ячейке A есть совпадение
                        foreach (var row in linearExpressionsSheet.RowsUsed())
                        {
                            var cellD = row.Cell("D").GetValue<string>();
                            var cellA = row.Cell("A").GetValue<string>();

                            // Извлечение числа из ячейки A и проверка на совпадение
                            var match = Regex.Match(cellA, @"\d+");
                            if (match.Success && int.TryParse(match.Value, out int cellAValue) && slabValues.Contains(cellAValue))
                            {
                                if (cellD.Contains("[·o51]"))
                                {
                                    var targetCell = row.Cell("C");
                                    logger.LogInformation($"Обновляем [·o51] в столбце C с {targetCell.Value} до {roundedPriceUsd}");
                                    targetCell.Value = roundedPriceUsd;
                                }
                                else if (cellD.Contains("[·o52]"))
                                {
                                    var targetCell = row.Cell("C");
                                    logger.LogInformation($"Обновляем [·o52] в столбце C с {targetCell.Value} до {roundedPriceWithCoef}");
                                    targetCell.Value = roundedPriceWithCoef;
                                }
                                else if (cellD.Contains("[·o53]"))
                                {
                                    var targetCell = row.Cell("C");
                                    logger.LogInformation($"Обновляем [·o53] в столбце C с {targetCell.Value} до {roundedPriceWithCoef}");
                                    targetCell.Value = roundedPriceWithCoef;
                                }
                            }
                        }

                        // Сохраняем изменения в файле проекта
                        projectWorkbook.SaveAs(projectFilePath);
                        logger.LogInformation("Файл проекта успешно обновлен.");
                    }
                }
                else
                {
                    logger.LogWarning("Не найдены данные о ценах для обновления файла проекта.");
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Ошибка при обработке скачанного файла: {ex.Message}");
            }
            finally
            {
                // Удаляем скачанный файл
                if (File.Exists(downloadedFilePath))
                {
                    File.Delete(downloadedFilePath);
                    logger.LogInformation("Скачанный файл успешно удален.");
                }
            }
        }
    }
}
