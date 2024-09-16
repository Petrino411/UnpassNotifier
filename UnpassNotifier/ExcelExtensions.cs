using OfficeOpenXml;
using UnpassNotifier.Classes;

namespace UnpassNotifier;

public static class ExcelExtensions
{
    private const int headerRow = 11;
    private const int subHeaderRow = 12;

    public static async Task<List<NotifyItem>> ExcelParse(string filePath, HashSet<Discipline> disciplines)
    {
        Console.WriteLine($"{filePath.Split(@"\").Last()} в процессе");
        var notifyItems = new List<NotifyItem>();

        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets.First();


        for (var row = 13; row <= worksheet.Dimension.End.Row; row++)
        {
            var fio = worksheet.Cells[row, 2].Text; // Предположим, что ФИО в первой колонке
            for (var col = 4; worksheet.Cells[row, col].Value != null; col++)
            {
                var discipline = disciplines.FirstOrDefault(x => x.DisciplineName.Equals(
                        worksheet.Cells[12, col].Text.Trim(),
                        StringComparison.CurrentCultureIgnoreCase
                    )
                );
                if (discipline == null)
                {
                    throw new AggregateException($"Дисциплина |'{worksheet.Cells[12, col].Text}'|. Ошибка в поиске в discipline.");
                }

                var controlResult = worksheet.Cells[row, col].Text;

                // Проверяем оценку (если оценка ниже 3 или неявка)
                if (int.TryParse(controlResult, out var grade) && grade < 3
                    || controlResult.Equals("неявка", StringComparison.CurrentCultureIgnoreCase)
                    || controlResult.Equals("незачтено", StringComparison.CurrentCultureIgnoreCase))
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
                        Discipline = discipline,
                        ControlResult = controlResult == "2" ? "2 (неудовлетворительно)" : controlResult,
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

        for (var col = 4; col <= worksheet.Dimension.End.Column; col++)
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

    private static string FindAttestationTypeByDiscipline(Dictionary<string, List<string>> disciplines,
        string discipline)
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