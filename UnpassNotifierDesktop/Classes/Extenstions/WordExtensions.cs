using System.Collections.ObjectModel;
using System.Drawing;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Spire.Doc;
using UnpassNotifierDesktop.Classes.Models;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Border = Xceed.Document.NET.Border;
using Document = Spire.Doc.Document;
using Font = Xceed.Document.NET.Font;

namespace UnpassNotifierDesktop.Classes.Extenstions;

public static class WordExtensions
{
    #region Reader

    // public static async Task<HashSet<Discipline>?> DisciplinesAttestationFill(string? filePath)
    // {
    //     if (filePath is null)
    //         return null;
    //     try
    //     {
    //         using var document = DocX.Load(filePath);
    //         var table = document?.Tables.FirstOrDefault();
    //         if (table == null)
    //             return null;
    //
    //         var rows = table.Rows;
    //         rows.RemoveAt(0);
    //         if (rows.Count == 0)
    //             return null;
    //
    //         var disciplines = new HashSet<Discipline>();
    //         for (var row = 1; row <= rows.Count; row++)
    //         {
    //             var disciplineName = table.Rows[row].Cells[1].Paragraphs.First().Text;
    //             var typeControl = table.Rows[row].Cells[2].Paragraphs.First().Text;
    //             var attestationDate = table.Rows[row].Cells[3].Paragraphs.First().Text;
    //
    //             if (string.IsNullOrWhiteSpace(typeControl))
    //             {
    //                 var tryRow = row - 1;
    //                 while (string.IsNullOrWhiteSpace(typeControl))
    //                 {
    //                     typeControl = table.Rows[tryRow].Cells[2].Paragraphs.First().Text;
    //                     tryRow -= 1;
    //                 }
    //             }
    //
    //             disciplines.Add(new Discipline(disciplineName.Trim(), attestationDate.Trim(), typeControl.Trim()));
    //         }
    //
    //         return disciplines;
    //     }
    //     catch (Exception e)
    //     {
    //         Console.WriteLine(e);
    //         return null;
    //     }
    // }

    #endregion

    public static bool TemplateCheck(string? templatePath)
    {
        var document = DocX.Load(templatePath);
        var isSuccess = document.Text.Contains("{{ФИО}}")
                        && document.Text.Contains("{{дата}}")
                        && document.Tables.FirstOrDefault(
                            x => x.ColumnCount == 5 && x.RowCount > 0
                        ) != null;
        document.Dispose();
        return isSuccess;
    }

    #region Generator

    public static void WordGenerate(List<NotifyItem> notifyItems, string outputDirectory, string templatePath,
        ObservableCollection<WordFilePathModel> outputFiles, ListView outputFilesView, Queue<Task> tasks)
    {
        var curTime = DateTime.Now;
        Directory.CreateDirectory(outputDirectory + @"\Word");
        Directory.CreateDirectory(outputDirectory + @"\PDF");
        tasks.Enqueue(Task.Run(() =>
        {
            foreach (var notifyItem in notifyItems)
            {
                var currentWordPath = outputDirectory + @$"\Word\{notifyItem.FIO}.docx";
                using var document = DocX.Load(templatePath);

                // Подстановка данных в документ
                document.ReplaceText("{{ФИО}}", notifyItem.FIO);
                document.ReplaceText("{{дата}}", curTime.ToShortDateString());

                var table = document.Tables.First(x => x.ColumnCount == 5);
                if (table.RowCount > 1)
                {
                    for (var row = table.RowCount - 1; row != 0; row--)
                        table.RemoveRow(row);
                }

                var smallFontFormat = new Formatting
                {
                    Size = 11,
                    FontFamily = new Font("Times New Roman"),
                };

                foreach (var item in notifyItem.UnpassedList)
                {
                    var curRow = table.InsertRow();
                    curRow.Cells[0].Paragraphs[0]
                        .Append($"{notifyItem.UnpassedList.IndexOf(item) + 1}.", smallFontFormat);
                    curRow.Cells[1].Paragraphs[0]
                        .Append($"{item.Discipline}", smallFontFormat);
                    curRow.Cells[2].Paragraphs[0]
                        .Append($"{item.Discipline.TypeControl}", smallFontFormat);
                    curRow.Cells[3].Paragraphs[0]
                        .Append($"{item.Discipline.AttestationDate}", smallFontFormat);
                    curRow.Cells[4].Paragraphs[0]
                        .Append($"{item.ControlResult}", smallFontFormat);
                }


                table.SetBorder(TableBorderType.InsideH,
                    new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));
                table.SetBorder(TableBorderType.Bottom,
                    new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));
                table.SetBorder(TableBorderType.Left,
                    new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));
                table.SetBorder(TableBorderType.Right,
                    new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));
                table.SetBorder(TableBorderType.Top,
                    new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));
                table.SetBorder(TableBorderType.InsideV,
                    new Border(BorderStyle.Tcbs_single, BorderSize.one, 1, Color.Black));

                document.SaveAs(currentWordPath);

                outputFilesView.Dispatcher.InvokeAsync(() =>
                {
                    outputFiles.Add(new WordFilePathModel(currentWordPath));
                    outputFilesView.Items.Refresh();
                });
            }
        }));
    }

    public static bool ConvertDocxToPdf(string inputFile, string outputFile)
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

    #endregion
}