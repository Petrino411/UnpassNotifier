using OfficeOpenXml;
using UnpassNotifier.Classes;

namespace UnpassNotifier;

class Program
{
    static void Main(string[] args)
    {
        string filePath = "207ис.xlsx";  // Замените на путь к вашему файлу

        List<NotifyItem> notifyItems = new List<NotifyItem>();

        // Чтение файла Excel
        using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; 

            
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                string fio = worksheet.Cells[row, 2].Text; // Предположим, что ФИО в первой колонке
                string disciplineName = worksheet.Cells[row, 2].Text; // Предположим, что дисциплина во второй колонке
                string typeControl = worksheet.Cells[row, 3].Text; // Тип контроля (зачет/экзамен)
                string controlResult = worksheet.Cells[row, 4].Text; // Результат контроля (оценка)

                // Проверяем оценку (если оценка ниже 3 или неявка)
                if (int.TryParse(controlResult, out int grade) && grade < 3 || controlResult.ToLower() == "неявка")
                {
                    // Ищем существующий объект NotifyItem для этого студента
                    NotifyItem notifyItem = notifyItems.Find(item => item.FIO == fio);

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
        using (StreamWriter writer = new StreamWriter("неявки_и_незданные.txt"))
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