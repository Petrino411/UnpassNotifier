using OfficeOpenXml;
using UnpassNotifier.Classes;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace UnpassNotifier;

internal class Program
{
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
            .First(x => x.Contains("УВЕДОМЛЕНИЕ")); // TODO: переделать на выбор шаблона
        var excelFiles = Directory
            .GetFiles(resourcesDirectory + @"\Excel", "*.xlsx", SearchOption.AllDirectories);
        var wordFiles = Directory
            .GetFiles(resourcesDirectory + @"\Word", "*.docx", SearchOption.AllDirectories);

        var tasks = new List<Task>();

        Console.WriteLine("Запуск обработки Excel файлов");
        foreach (var excelFile in excelFiles)
        {
            if (excelFile.Contains("~$")) continue;
            tasks.Add(Task.Run(async () =>
            {
                var groupName = excelFile.Split(@"\").Last().Split('.').First();

                var wordFilePath = wordFiles.FirstOrDefault(x => x.Contains(groupName));
                var disciplines = await WordExtensions.DisciplinesAttestationFill(wordFilePath);
                if (disciplines == null)
                {
                    Console.WriteLine($"Программа не смогла прочесть данные из графика: {wordFilePath ?? groupName}. Файла либо нет, либо произошла ошибка.");
                    return;
                }

                var notifies = await ExcelExtensions.ExcelParse(excelFile, disciplines);

                var targetFile = groupName + $@" - {DateTime.Now.ToShortDateString()}";
                var outputDirectory = Directory.CreateDirectory(resultsDirectory + @$"\{targetFile}").FullName;

                Console.WriteLine($"Создание Word уведомлений для {groupName}");
                await WordExtensions.WordGenerate(notifies, outputDirectory, templatePath, disciplines);
                return;
            }));
        }


        Task.WaitAll(tasks.ToArray());

        Console.WriteLine("Конец работы.");
        return;
    }
}