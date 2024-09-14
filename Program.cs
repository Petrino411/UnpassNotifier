using OfficeOpenXml;
using UnpassNotifier.Classes;

namespace UnpassNotifier;

internal static class Program
{
    private static void Main(string[] args)
    {
        var directories = Directory.GetDirectories(Environment.CurrentDirectory);
        string resourcesDirectory;
        try
        {
            resourcesDirectory = directories.First(x => x.Contains("Resources"));
        }
        catch (Exception e)
        {
            Console.WriteLine("Должна быть папка Resources в директории программы");
            throw;
        }
        var excelFiles = Directory.GetFiles(resourcesDirectory, "*.xlsx", SearchOption.AllDirectories);
        
        
        var filePath = excelFiles.First();  

        var notifyItems = new List<NotifyItem>();
        
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets.First();

            var disciplines = new Dictionary<string, List<string>>();

            var headerRow = 11;
            var subHeaderRow = 12;

            var lastAttestationType = "";

            for (int col = 4; col <= worksheet.Dimension.End.Column; col++)
            {
                var attestationType = worksheet.Cells[headerRow, col].Text;
                if (!string.IsNullOrEmpty(attestationType))
                {
                    lastAttestationType = attestationType;
                }
                var disciplineName = worksheet.Cells[subHeaderRow, col].Text;  // Название дисциплины

                if (!string.IsNullOrWhiteSpace(disciplineName))
                {
                    if (!disciplines.ContainsKey(lastAttestationType))
                    {
                        disciplines[lastAttestationType] = [];
                    }
                    disciplines[lastAttestationType].Add(disciplineName);
                }
            }
            var listOfDisciplines = disciplines.Values.SelectMany(x => x).ToList();

            for (var row = 13; row <= worksheet.Dimension.End.Row; row++)
            {
                var fio = worksheet.Cells[row, 2].Text; // Предположим, что ФИО в первой колонке
                for (var col = 4; worksheet.Cells[row,col].Value != null; col++)
                {
                    var disciplineName = listOfDisciplines[col-4];
                    var attestationType = FindAttestationTypeByDiscipline(disciplines, disciplineName);
                    var controlResult = worksheet.Cells[row, col].Text;
                    
                    // Проверяем оценку (если оценка ниже 3 или неявка)
                    if (int.TryParse(controlResult, out int grade) && grade < 3 || controlResult.ToLower() == "неявка" || controlResult.ToLower() == "незачтено")
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
        }
        

        Console.WriteLine("Анализ завершен. Результаты сохранены в файл 'неявки_и_незданные.txt'.");
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