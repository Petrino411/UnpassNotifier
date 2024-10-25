using System.IO;
using OfficeOpenXml;

namespace UnpassNotifierDesktop.Classes.Extenstions;

public static class ExcelExtensions
{
    public static async IAsyncEnumerable<(List<NotifyItem> items, string groupName)> ExcelParse(ExcelPackage package)
    {
        foreach (var worksheet in package.Workbook.Worksheets)
        {
            var notifyItems = new List<NotifyItem>();
            for (var row = 2; !string.IsNullOrEmpty(worksheet.Cells[row, 1].Text); row++)
            {
                var fio = worksheet.Cells[row, 2].Text.Trim();
                for (var col = 3; worksheet.Cells[row, col].Value != null; col++)
                {
                    var discipline = worksheet.Cells[1, col].Text?.Trim();
                    if (string.IsNullOrEmpty(discipline) == null)
                    {
                        throw new AggregateException(
                            $"Дисциплина |'{worksheet.Cells[10, col].Text}'|. Ошибка в строке...");
                    }

                    var controlResult = worksheet.Cells[row, col].Text;

                    // Проверяем оценку (если оценка ниже 3 или неявка)
                    if (int.TryParse(controlResult, out var grade) && grade < 3
                        || controlResult.Contains("неявка", StringComparison.CurrentCultureIgnoreCase)
                        || controlResult.Contains("незачтено", StringComparison.CurrentCultureIgnoreCase))
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
                            Discipline = new Discipline(discipline),
                            ControlResult = controlResult == "2" ? "2 (неудовлетворительно)" : controlResult,
                        });
                    }
                }
            }

            yield return (notifyItems, worksheet.Name);
        }
    }
}