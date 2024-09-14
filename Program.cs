using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeOpenXml;
using UnpassNotifier.Classes;

namespace UnpassNotifier;

internal class Program
{
    private const int headerRow = 11;
    private const int subHeaderRow = 12;
    private static string period1 = "14.09.2024";
    private static string period2 = "14.10.2024";

    private static void Main(string[] args)
    {
        Console.WriteLine("Запуск программы");

        var directories = Directory.GetDirectories(Environment.CurrentDirectory);
        var resourcesDirectory = directories.FirstOrDefault(x => x.Contains("Resources")) ??
                                 Directory.CreateDirectory(Environment.CurrentDirectory + @"\Resources").FullName;
        var resultsDirectory = directories.FirstOrDefault(x => x.Contains("Result"))
                               ?? Directory.CreateDirectory(Environment.CurrentDirectory + @"\Result").FullName;

        var templatePath = Directory
            .GetFiles(resourcesDirectory, @"*.docx", SearchOption.AllDirectories)
            .First(x => x.Contains("УВЕДОМЛЕНИЕ"));
        var excelFiles = Directory
            .GetFiles(resourcesDirectory + @"\Excel", "*.xlsx", SearchOption.AllDirectories);

        var tasks = new List<Task>();

        Console.WriteLine("Запуск обработки Excel файлов");
        foreach (var excelFile in excelFiles)
        {
            if (excelFile.Contains("~$")) continue;
            tasks.Add(Task.Run(async () =>
            {
                var notifies = await ExcelParse(excelFile);

                var targetFile = excelFile.Split(@"\").Last() + $@" - {DateTime.Now.ToShortDateString()}";
                var outputDirectory = Directory
                    .CreateDirectory(resourcesDirectory + @$"\{targetFile}").FullName;

                Console.WriteLine($"Создание Word файлов для {targetFile}");
                await WordGenerate(notifies, outputDirectory, templatePath);
            }));
        }


        Task.WaitAll(tasks.ToArray());

        Console.WriteLine("Конец работы.");
    }

    private static async Task WordGenerate(List<NotifyItem> notifyItems, string outputDirectory, string templatePath)
    {
        var innerTasks = new List<Task>();

        foreach (var notifyItem in notifyItems)
        {
            innerTasks.Add(Task.Run(() =>
            {
                var currentPath = outputDirectory + @$"\{notifyItem.FIO}.docx";
                // File.Copy(templatePath, currentPath, true);

                using var document = WordprocessingDocument.Open(templatePath, false);
                var innerText = document.MainDocumentPart.Document.InnerText;
                
                if (innerText.Contains("{{ФИО}}"))
                {
                    innerText = innerText.Replace("{{ФИО}}", notifyItem.FIO);
                }

                if (innerText.Contains("{{Дата}}"))
                {
                    innerText = innerText.Replace("{{Дата}}", DateTime.Now.ToShortDateString());
                }

                if (innerText.Contains("{{период1}}"))
                {
                    innerText = innerText.Replace("{{период1}}", period1);
                }

                if (innerText.Contains("{{период2}}"))
                {
                    innerText = innerText.Replace("{{период2}}", period2);
                }

                // document.MainDocumentPart.Document.InnerText = innerText;
                // document.MainDocumentPart.Document.Body.InnerText = innerText;
                
                
                document.SaveAs(currentPath);
                document.Close();
            }));
        }

        Task.WaitAll(innerTasks.ToArray());
    }

    private static async Task<List<NotifyItem>> ExcelParse(string filePath)
    {
        var notifyItems = new List<NotifyItem>();

        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.First();

        var disciplines = FillDisciplines(worksheet);

        var listOfDisciplines = disciplines.Values.SelectMany(x => x).ToList();

        for (var row = 13; row <= worksheet.Dimension.End.Row; row++)
        {
            var fio = worksheet.Cells[row, 2].Text; // Предположим, что ФИО в первой колонке
            for (var col = 4; worksheet.Cells[row, col].Value != null; col++)
            {
                var disciplineName = listOfDisciplines[col - 4];
                var attestationType = FindAttestationTypeByDiscipline(disciplines, disciplineName);
                var controlResult = worksheet.Cells[row, col].Text;

                // Проверяем оценку (если оценка ниже 3 или неявка)
                if (int.TryParse(controlResult, out var grade) && grade < 3 ||
                    controlResult.ToLower() == "неявка" || controlResult.ToLower() == "незачтено")
                {
                    // Ищем существующий объект NotifyItem для этого студента
                    var notifyItem = notifyItems.Find(item => item.FIO == fio);

                    if (notifyItem == null)
                    {
                        notifyItem = new NotifyItem { FIO = fio };
                        notifyItems.Add(notifyItem);
                    }

                    // Добавляем информацию о незданной дисциплине
                    notifyItem.UnpassedList.Add(new UnpassItem
                    {
                        DisciplineName = disciplineName,
                        TypeControl = attestationType,
                        ControlResult = controlResult
                    });
                }
            }
        }

        return notifyItems;
    }

    private static Dictionary<string, List<string>> FillDisciplines(ExcelWorksheet worksheet)
    {
        var disciplines = new Dictionary<string, List<string>>();


        var lastAttestationType = "";

        for (int col = 4; col <= worksheet.Dimension.End.Column; col++)
        {
            var attestationType = worksheet.Cells[headerRow, col].Text;
            if (!string.IsNullOrEmpty(attestationType))
            {
                lastAttestationType = attestationType;
            }

            var disciplineName = worksheet.Cells[subHeaderRow, col].Text; // Название дисциплины

            if (!string.IsNullOrWhiteSpace(disciplineName))
            {
                if (!disciplines.ContainsKey(lastAttestationType))
                {
                    disciplines[lastAttestationType] = [];
                }

                disciplines[lastAttestationType].Add(disciplineName);
            }
        }

        return disciplines;
    }

    static string FindAttestationTypeByDiscipline(Dictionary<string, List<string>> disciplines, string discipline)
    {
        foreach (var entry in disciplines)
        {
            if (entry.Value.Contains(discipline))
            {
                return entry.Key;
            }
        }

        return null;
    }
}