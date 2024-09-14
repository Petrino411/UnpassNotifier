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
        
        
        var filePath = "207ис.xlsx";  // Замените на путь к вашему файлу

        var notifyItems = new List<NotifyItem>();

        // Чтение файла Excel
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0]; 

            
            for (var row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var fio = worksheet.Cells[row, 2].Text; // Предположим, что ФИО в первой колонке
                var disciplineName = worksheet.Cells[row, 2].Text; // Предположим, что дисциплина во второй колонке
                var typeControl = worksheet.Cells[row, 3].Text; // Тип контроля (зачет/экзамен)
                var controlResult = worksheet.Cells[row, 4].Text; // Результат контроля (оценка)

                // Проверяем оценку (если оценка ниже 3 или неявка)
                if (int.TryParse(controlResult, out int grade) && grade < 3 || controlResult.ToLower() == "неявка")
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
                        TypeControl = typeControl,
                        ControlResult = controlResult
                    });
                }
            }
        }

        // Сохранение результата в файл (например, в txt)
        using (var writer = new StreamWriter("неявки_и_незданные.txt"))
        {
            foreach (var notifyItem in notifyItems)
            {
                writer.WriteLine($"ФИО: {notifyItem.FIO}");
                foreach (var unpassItem in notifyItem.UnpassedList)
                {
                    writer.WriteLine($"  Дисциплина: {unpassItem.DisciplineName}, Тип: {unpassItem.TypeControl}, Результат: {unpassItem.ControlResult}");
                }
            }
        }

        Console.WriteLine("Анализ завершен. Результаты сохранены в файл 'неявки_и_незданные.txt'.");
    }
}