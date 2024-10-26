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
            try
            {
                if (worksheet.Name.Contains("Лист"))
                {
                    continue;
                }

                var hasTypeControl = worksheet.Cells[1, 3].Text != null &&
                                     string.IsNullOrEmpty(worksheet.Cells[1, 1].Text);
                var startRow = hasTypeControl ? 3 : 2;
                var disciplineRow = hasTypeControl ? 2 : 1;
                var startColumn = worksheet.Cells.Text
                    .Contains("Основа обучения", StringComparison.InvariantCultureIgnoreCase)
                    ? 4
                    : 3;

                if (!hasTypeControl)
                    Console.Error.WriteLine($"{worksheet.Name} не имеет форм контроля, либо не в ожидаемом формате");

                var lastSeenTypeControl = hasTypeControl ? worksheet.Cells[1, 3].Text : null;


                for (var row = startRow;
                     !string.IsNullOrEmpty(worksheet.Cells[row, 1].Text);
                     row++)
                {
                    var fio = worksheet.Cells[row, 2].Text.Trim();

                    for (var col = startColumn;
                         !string.IsNullOrEmpty(worksheet.Cells[disciplineRow, col].Text);
                         col++)
                    {
                        var disciplineText = worksheet.Cells[disciplineRow, col].Text?.Trim();
                        if (hasTypeControl)
                        {
                            var typeControl = worksheet.Cells[disciplineRow - 1, col].Text;
                            if (!string.IsNullOrEmpty(typeControl))
                                lastSeenTypeControl = typeControl;
                        }

                        if (string.IsNullOrEmpty(disciplineText))
                        {
                            throw new AggregateException(
                                $"Дисциплина |'{worksheet.Cells[10, col].Text}'|. Ошибка в строке...");
                        }

                        var discipline = hasTypeControl
                            ? new Discipline(disciplineText, lastSeenTypeControl)
                            : new Discipline(disciplineText);

                        var controlResult = worksheet.Cells[row, col].Text;

                        if (string.IsNullOrEmpty(controlResult))
                        {
                            await Console.Error.WriteLineAsync(
                                $"{worksheet.Name}: {fio}. Пустое значение по столбцу `{discipline}`");
                        }

                        if (controlResult.Contains("справка", StringComparison.InvariantCultureIgnoreCase))
                        {
                            continue;
                        }


                        // Проверяем оценку (если оценка ниже 3 или неявка)
                        if (string.IsNullOrEmpty(controlResult)
                            || int.TryParse(controlResult, out var grade) && grade == 2
                            || controlResult.Contains("неявка", StringComparison.CurrentCultureIgnoreCase)
                            || controlResult.Contains("незачтено", StringComparison.CurrentCultureIgnoreCase)
                            || controlResult.Contains("не зачтено", StringComparison.CurrentCultureIgnoreCase))
                        {
                            // Ищем существующий объект NotifyItem для этого студента
                            var notifyItem = notifyItems.FirstOrDefault(item => item.FIO == fio);

                            if (notifyItem == null)
                            {
                                notifyItem = new NotifyItem { FIO = fio };
                                notifyItems.Add(notifyItem);
                            }

                            // Добавляем информацию о незданной дисциплине
                            notifyItem.UnpassedList.Add(new UnpassItem
                            {
                                Discipline = discipline,
                                ControlResult =
                                    controlResult == "2" ? "2 (неудовлетворительно)" : controlResult,
                            });
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine(e);
            }

            yield return (notifyItems, worksheet.Name);
        }
    }
}