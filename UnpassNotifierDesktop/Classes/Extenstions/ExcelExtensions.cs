using System.IO;
using OfficeOpenXml;

namespace UnpassNotifierDesktop.Classes.Extenstions;

public static class ExcelExtensions
{
    public static async IAsyncEnumerable<(List<NotifyItem> items, string groupName)> ExcelParse(ExcelPackage package, HashSet<Discipline> disciplines)
    {
        foreach (var worksheet in package.Workbook.Worksheets)
        {
            var notifyItems = new List<NotifyItem>();
            for (var row = 13; !string.IsNullOrEmpty(worksheet.Cells[row, 1].Text); row++)
            {
                var fio = worksheet.Cells[row, 2].Text.Trim();
                for (var col = 6; worksheet.Cells[row, col].Value != null; col++)
                {
                    var discipline = disciplines.FirstOrDefault(x => x.DisciplineName.Contains(
                            worksheet.Cells[10, col].Text.Trim(),
                            StringComparison.CurrentCultureIgnoreCase
                        )
                    );
                    if (discipline == null)
                    {
                        throw new AggregateException(
                            $"Дисциплина |'{worksheet.Cells[10, col].Text}'|. Ошибка в поиске в discipline.");
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

            yield return (notifyItems, worksheet.Name);
        }
    }
}