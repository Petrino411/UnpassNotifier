using System.IO;
using System.Windows;
using System.Windows.Controls;
using Spire.Doc;
using UnpassNotifierDesktop.Classes.Models;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Document = Spire.Doc.Document;

namespace UnpassNotifierDesktop.Classes.Extenstions;

public static class WordExtensions
{
    #region Reader

    public static async Task<HashSet<Discipline>?> DisciplinesAttestationFill(string? filePath)
    {
        if (filePath is null)
            return null;
        try
        {
            using var document = DocX.Load(filePath);
            var table = document?.Tables.FirstOrDefault();
            if (table == null)
                return null;

            var rows = table.Rows;
            rows.RemoveAt(0);
            if (rows.Count == 0)
                return null;

            var disciplines = new HashSet<Discipline>();
            for (var row = 1; row <= rows.Count; row++)
            {
                var disciplineName = table.Rows[row].Cells[1].Paragraphs.First().Text;
                var typeControl = table.Rows[row].Cells[2].Paragraphs.First().Text;
                var attestationDate = table.Rows[row].Cells[3].Paragraphs.First().Text;

                if (string.IsNullOrWhiteSpace(typeControl))
                {
                    var tryRow = row - 1;
                    while (string.IsNullOrWhiteSpace(typeControl))
                    {
                        typeControl = table.Rows[tryRow].Cells[2].Paragraphs.First().Text;
                        tryRow -= 1;
                    }
                }

                disciplines.Add(new Discipline(disciplineName.Trim(), attestationDate.Trim(), typeControl.Trim()));
            }

            return disciplines;
        }
        catch (Exception e)
        {
            Console.WriteLine(e);
            return null;
        }
    }

    #endregion

    #region Generator

    public static async Task WordGenerate(List<NotifyItem> notifyItems, string outputDirectory, string templatePath,
        ListView? outputFiles, Queue<Task> tasks)
    {
        Directory.CreateDirectory(outputDirectory + @"\Word");
        Directory.CreateDirectory(outputDirectory + @"\PDF");

        foreach (var notifyItem in notifyItems)
        {
            tasks.Enqueue( Task.Run(async () =>
            {
                var currentWordPath = outputDirectory + @$"\Word\{notifyItem.FIO}.docx";
                var currentPdfPath = outputDirectory + $@"\PDF\{notifyItem.FIO}.pdf";
                // Открытие документа-шаблона
                using var document = DocX.Load(templatePath);

                // Подстановка данных в документ
                document.ReplaceText("{{ФИО}}", notifyItem.FIO);
                document.ReplaceText("{{дата}}", DateTime.Now.ToShortDateString());
                // document.ReplaceText("{{период1}}", period1);
                // document.ReplaceText("{{период2}}", period2);

                var table = document.Tables.FirstOrDefault();
                // Заполнение таблицы
                if (table != null)
                {
                    table.RemoveRow(1);

                    var smallFontFormat = new Formatting
                    {
                        Size = 11
                    };

                    for (var row = 1; row <= notifyItem.UnpassedList.Count; row++)
                    {
                        table.InsertRow();
                        table.Rows[row].Cells[0].Paragraphs[0]
                            .Append($"{row}.", smallFontFormat);
                        table.Rows[row].Cells[1].Paragraphs[0]
                            .Append($"{notifyItem.UnpassedList[row - 1].Discipline.DisciplineName}", smallFontFormat);
                        table.Rows[row].Cells[2].Paragraphs[0]
                            .Append($"{notifyItem.UnpassedList[row - 1].Discipline.TypeControl}", smallFontFormat);
                        table.Rows[row].Cells[3].Paragraphs[0]
                            .Append($"{notifyItem.UnpassedList[row - 1].Discipline.AttestationDate}", smallFontFormat);
                        table.Rows[row].Cells[4].Paragraphs[0]
                            .Append($"{notifyItem.UnpassedList[row - 1].ControlResult}", smallFontFormat);
                    }
                }

                Console.WriteLine($"Завершение создания {currentWordPath}");
                document.SaveAs(currentWordPath);

                // var hasPdf = await ConvertDocxToPdf(currentWordPath, currentPdfPath);
                //
                // outputFiles?.Dispatcher.InvokeAsync(() =>
                // {
                //     outputFiles.Items.Add(hasPdf
                //         ? new FilePathModel(currentWordPath, currentPdfPath)
                //         : new FilePathModel(currentWordPath));
                // });
                
                outputFiles?.Dispatcher.InvokeAsync(() =>
                {
                    outputFiles.Items.Add(new FilePathModel(currentWordPath));
                });
                
            }));
        }
        
    }

    #endregion

    public static async Task<bool> ConvertDocxToPdf(string inputFile, string outputFile)
    {
        try
        {
            Console.WriteLine($"Создание PDF для {inputFile}");
            using var document = new Document();
            document.LoadFromFile(inputFile);
            document.SaveToFile(outputFile, FileFormat.PDF);
            document.Close();
            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine($"Ошибка при конвертации {inputFile}: {e.Message}");
            return false;
        }
    }
}