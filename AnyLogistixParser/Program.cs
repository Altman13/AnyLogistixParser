using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;

namespace AnyLogistixParser
{
    class Program
    {
        private static IConfigurationRoot configuration;
        private static ILogger logger;

        static async Task Main(string[] args)
        {
            Initialize();

            string downloadUrl = configuration["DownloadUrl"];
            string downloadedFilePath = @"downloaded_file.xlsx";
            string projectFilePath = configuration["ProjectFilePath"];

            if (!DownloadFile(downloadUrl, downloadedFilePath))
            {
                return;
            }

            var slabValues = ParseSlabValues(projectFilePath);
            if (slabValues == null)
            {
                return;
            }

            double weightedAveragePriceRub = CalculateWeightedAveragePrice(downloadedFilePath);
            if (weightedAveragePriceRub == 0)
            {
                logger.LogWarning("Не найдены данные о ценах для обновления файла проекта.");
                return;
            }

            int roundedPriceWithCoef = (int)Math.Round(weightedAveragePriceRub * 0.63);
            double usdRub = await GetUsdToRubExchangeRate();

            int roundedPriceUsd = (int)Math.Round(roundedPriceWithCoef / usdRub);
            
            logger.LogInformation($"Средневзвешенная цена в рублях: {weightedAveragePriceRub}");
            logger.LogInformation($"Округленная цена с коэффициентом в рублях: {roundedPriceWithCoef}");
            logger.LogInformation($"Округленная цена в долларах: {roundedPriceUsd}");

            UpdateProjectFile(projectFilePath, slabValues, roundedPriceUsd, roundedPriceWithCoef);
        }

        private static void Initialize()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            configuration = builder.Build();

            var loggerFactory = LoggerFactory.Create(logging =>
            {
                logging.AddConsole();
            });

            logger = loggerFactory.CreateLogger<Program>();
        }

        private static bool DownloadFile(string url, string filePath)
        {
            try
            {
                using (var webClient = new System.Net.WebClient())
                {
                    webClient.DownloadFile(url, filePath);
                    logger.LogInformation("Файл успешно скачан.");
                    return true;
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Ошибка при скачивании файла по ссылке {url}: {ex.Message}");
                return false;
            }
        }

        private static List<int> ParseSlabValues(string projectFilePath)
        {
            var slabValues = new List<int>();

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
                return null;
            }

            logger.LogInformation($"Найденные значения из ячеек C: {string.Join(", ", slabValues)}");
            return slabValues;
        }

        private static double CalculateWeightedAveragePrice(string downloadedFilePath)
        {
            double totalPriceRub = 0.0;
            int priceCount = 0;
            bool startParsing = false;

            try
            {
                using (var workbook = new XLWorkbook(downloadedFilePath))
                {
                    int sheetNumber = int.Parse(configuration["ExcelSettings:SheetNumber"]);
                    int priceColumn = int.Parse(configuration["ExcelSettings:PriceColumn"]);
                    int skipRows = int.Parse(configuration["ExcelSettings:SkipRows"]);
                    string searchString = configuration["ExcelSettings:SearchString"];

                    var worksheet = workbook.Worksheet(sheetNumber);

                    foreach (var row in worksheet.RowsUsed().Skip(skipRows))
                    {
                        var firstCell = row.Cell(1).GetValue<string>();
                        if (!startParsing && firstCell.Contains(searchString))
                        {
                            startParsing = true;
                        }

                        if (startParsing)
                        {
                            var priceCell = row.Cell(priceColumn);
                            if (double.TryParse(priceCell.GetValue<string>().Replace(" ", ""), out double price))
                            {
                                totalPriceRub += price;
                                priceCount++;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Ошибка при обработке скачанного файла: {ex.Message}");
                return 0;
            }

            if (priceCount > 0)
            {
                return totalPriceRub / priceCount;
            }
            else
            {
                return 0;
            }
        }

        private static async Task<double> GetUsdToRubExchangeRate()
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

        private static void UpdateProjectFile(string projectFilePath, List<int> slabValues, int roundedPriceUsd, int roundedPriceWithCoef)
        {
            try
            {
                using (var projectWorkbook = new XLWorkbook(projectFilePath))
                {
                    var linearExpressionsSheet = projectWorkbook.Worksheet("Linear expressions");

                    if (linearExpressionsSheet == null)
                    {
                        logger.LogError("Лист 'Linear expressions' не найден в файле проекта.");
                        return;
                    }

                    foreach (var row in linearExpressionsSheet.RowsUsed())
                    {
                        var cellD = row.Cell("D").GetValue<string>();
                        var cellA = row.Cell("A").GetValue<string>();

                        var match = Regex.Match(cellA, @"\d+");
                        if (match.Success && int.TryParse(match.Value, out int cellAValue) && slabValues.Contains(cellAValue))
                        {
                            var targetCell = row.Cell("C");

                            if (cellD.Contains("[·o51]"))
                            {
                                logger.LogInformation($"Обновляем [·o51] в столбце C с {targetCell.Value} до {roundedPriceUsd}");
                                targetCell.Value = roundedPriceUsd;
                            }
                            else if (cellD.Contains("[·o52]") || cellD.Contains("[·o53]"))
                            {
                                logger.LogInformation($"Обновляем {cellD} в столбце C с {targetCell.Value} до {roundedPriceWithCoef}");
                                targetCell.Value = roundedPriceWithCoef;
                            }
                        }
                    }

                    projectWorkbook.SaveAs(projectFilePath);
                    logger.LogInformation("Файл проекта успешно обновлен.");
                }
            }
            catch (Exception ex)
            {
                logger.LogError($"Ошибка при обновлении файла проекта: {ex.Message}");
            }
        }
    }
}
