using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.Win32;
using OfficeOpenXml;
using UnpassNotifierDesktop.Classes;
using UnpassNotifierDesktop.Classes.Extenstions;
using UnpassNotifierDesktop.Classes.Models;


namespace UnpassNotifierDesktop;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private string ResultsDirectory { get; }
    private string ResourcesDirectory { get; }
    private bool IsRunning { get; set; }
    private FilePathModel? TemplateFile { get; set; }
    private FilePathModel? ScheduleFile { get; set; }
    private FilePathModel? StatementFile { get; set; }
    private ObservableCollection<WordFilePathModel> OutputFiles { get; set; } = [];

    public MainWindow()
    {
        InitializeComponent();

        var directories = Directory.GetDirectories(Environment.CurrentDirectory);
        ResourcesDirectory = directories.FirstOrDefault(x => x.Contains("Resources")) ??
                             Directory.CreateDirectory(Environment.CurrentDirectory + @"\Resources").FullName;
        ResultsDirectory = directories.FirstOrDefault(x => x.Contains("Result"))
                           ?? Directory.CreateDirectory(Environment.CurrentDirectory + @"\Result").FullName;

        OutputFilesView.ItemsSource = OutputFiles;
        StatementFileLabel.MouseDoubleClick += OpenOnMouseDoubleClick;
        ScheduleFileLabel.MouseDoubleClick += OpenOnMouseDoubleClick;
        TemplateFileLabel.MouseDoubleClick += OpenOnMouseDoubleClick;
        OutputFilesView.MouseDoubleClick += OpenOnMouseDoubleClick;
    }

    #region InteractionEvents

    private static void OpenOnMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        try
        {
            if (sender is not ListView listView) return;

            var obj = listView.SelectedItem;
            if (obj is not WordFilePathModel filePathModel)
            {
                Console.WriteLine("Кажется, это не файл....");
                return;
            }

            if (!string.IsNullOrEmpty(filePathModel.PdfPath))
            {
                Process.Start(new ProcessStartInfo(filePathModel.PdfPath)
                {
                    UseShellExecute = true,
                });
                return;
            }

            Process.Start(new ProcessStartInfo(filePathModel.WordPath)
            {
                UseShellExecute = true,
            });
            return;
        }
        catch (Exception exception)
        {
            Console.WriteLine("Не удалось открыть файл");
        }
    }

    // protected override void OnClosing(CancelEventArgs e)
    // {
    //     if (Tasks.Count(x => x.IsCompleted) < Tasks.Count)
    //     {
    //         MessageBox.Show("Кажется, висит несколько задач. Подождите завершения их выполнения.");
    //         e.Cancel = true;
    //         return;
    //     }
    //
    //     base.OnClosing(e);
    // }

    #endregion

    private async Task<bool> WorkBody(FilePathModel scheduleFile, ExcelPackage excelPackage, ProgressBar progressBar)
    {
        try
        {
            var disciplines = await WordExtensions.DisciplinesAttestationFill(scheduleFile.FilePath);
            if (disciplines == null) return false;
            try
            {
                var tasks = new Queue<Task>();

                await foreach (var (items, groupName) in ExcelExtensions.ExcelParse(excelPackage, disciplines))
                {
                    var targetFile = $@"{groupName} - {DateTime.Now.ToShortDateString()}";
                    var outputDirectory = Directory.CreateDirectory(ResultsDirectory + @$"\{targetFile}").FullName;

                    Console.WriteLine($"Создание Word уведомлений для {groupName}");
                    WordExtensions.WordGenerate(items, outputDirectory,
                        TemplateFile!.FilePath, OutputFiles,
                        OutputFilesView, tasks);
                    progressBar.Dispatcher.InvokeAsync(() => { progressBar.Value++; });
                }

                await Task.WhenAll(tasks);
            }
            catch (Exception e)
            {
                Console.WriteLine($"Произошла ошибка, {e.Message}");
                return false;
            }

            return true;
        }
        catch (Exception e)
        {
            Console.WriteLine($"Произошла ошибка при извлечении периодов из {scheduleFile.FilePath}: {e.Message}");
            return false;
        }
    }

    #region Buttons

    private void SelectScheduleBtn(object sender, RoutedEventArgs e)
    {
        var fileDialog = new OpenFileDialog
        {
            Multiselect = true,
            Filter = "Графики аттестации|*.docx|All Files|*.*",
            Title = "Выберите файлы"
        };

        if (fileDialog.ShowDialog() != true) return;

        ScheduleFile = new FilePathModel(fileDialog.FileName);
        ScheduleFileLabel.Content = ScheduleFile;
    }

    private void SelectAttestationBtn(object sender, RoutedEventArgs e)
    {
        var fileDialog = new OpenFileDialog
        {
            Multiselect = true,
            Filter = "Сводные ведомости|*.xlsx|All Files|*.*",
            Title = "Выберите файлы"
        };

        if (fileDialog.ShowDialog() != true) return;

        StatementFile = new FilePathModel(fileDialog.FileName);
        StatementFileLabel.Content = StatementFile;
    }

    private void SelectTemplateBtn(object sender, RoutedEventArgs e)
    {
        var fileDialog = new OpenFileDialog
        {
#if DEBUG
            InitialDirectory = ResourcesDirectory,
#endif
            Filter = "Шаблон уведомления|*.docx|All Files|*.*",
            Title = "Выберите файлы"
        };

        if (fileDialog.ShowDialog() != true) return;


        if (!WordExtensions.TemplateCheck(fileDialog.FileName))
        {
            MessageBox.Show("Файл должен содержать {{ФИО}}, {{дата}} и таблицу с 2 строками");
            return;
        }

        TemplateFile = new FilePathModel(fileDialog.FileName);
        TemplateFileLabel.Content = TemplateFile;
    }

    private async void RunBtn(object sender, RoutedEventArgs e)
    {
        // Проверка на выполнение задачи и корректность путей
        if (IsRunning
            || TemplateFile == null
            || StatementFile == null
            || ScheduleFile == null)
        {
            MessageBox.Show("Какой-то из файлов не выбран или другая задача уже запущена");
            return;
        }

        IsRunning = true;
        OutputFiles.Clear();

        using var package = new ExcelPackage(StatementFile.GetFileInfo());

        var tasks = new Queue<Task>();

        ProgressBarParse.Value = 0;
        ProgressBarParse.Maximum = package.Workbook.Worksheets.Count;
        ParseStatusLabel.Content = "";
        ProgressBarParse.Visibility = Visibility.Visible;

        Console.WriteLine("Начата обработка");
        tasks.Enqueue(Task.Run(async () => { await WorkBody(ScheduleFile, package, ProgressBarParse); }));

        await Task.WhenAll(tasks);
        
        ParseStatusLabel.Content = "Создание файлов завершено!";
        ProgressBarParse.Visibility = Visibility.Collapsed;
        ProgressBarParse.Value = 0;
        
        Console.WriteLine("Обработка завершена");
        IsRunning = false;

    }

    #endregion

    // private void OutputFiles_OnMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    // {
    //     OutputFilesView.ContextMenu!.IsOpen = true;
    // }

    private async void ConvertToPdf(object sender, RoutedEventArgs e)
    {
        if (IsRunning)
        {
            MessageBox.Show("Какая-то задача уже запущена. Пожалуйста, подождите.");
            return;
        }

        IsRunning = true;
        var selectedItems = OutputFilesView.SelectedItems.Cast<WordFilePathModel>().ToList();

        // Обнуление прогресс-бара перед началом
        PdfProgressBar.Value = 0;
        PdfProgressBar.Maximum = selectedItems.Count;
        PdfStatusLabel.Content = "";
        PdfProgressBar.Visibility = Visibility.Visible;

        await Task.Run(() =>
        {
            // Счетчик обработанных файлов
            foreach (var selectedItem in selectedItems)
            {
                selectedItem.PdfPath = selectedItem.WordPath.Replace(@"\Word\", @"\PDF\") + ".pdf";

                var isSuccess = WordExtensions.ConvertDocxToPdf(selectedItem.WordPath, selectedItem.PdfPath);

                if (isSuccess)
                {
                    // Обновляем прогресс-бар
                    // Это нужно делать в UI-потоке
                    Dispatcher.Invoke(() =>
                    {
                        PdfProgressBar.Value++;
                        OutputFilesView.Items.Refresh();
                    });
                }
            }

            Console.WriteLine("Элементы преобразованы в pdf.");
        });

        // Обновляем статус после завершения
        PdfStatusLabel.Content = "Преобразование завершено!";
        IsRunning = false;
        await Task.Run(() =>
        {
            PdfProgressBar.Dispatcher.InvokeAsync(() =>
            {
                PdfProgressBar.Visibility = Visibility.Collapsed;
                PdfProgressBar.Value = 0;
            });
        });
    }

    #region ClearPaths

    private void TemplateClear_OnClick(object sender, RoutedEventArgs e)
    {
        if (IsRunning)
        {
            MessageBox.Show("В данный момент изменение невозможно");
            return;
        }

        TemplateFile = null;
        TemplateFileLabel.Content = "";
    }

    private void ScheduleClear_OnClick(object sender, RoutedEventArgs e)
    {
        if (IsRunning)
        {
            MessageBox.Show("В данный момент изменение невозможно");
            return;
        }

        ScheduleFile = null;
        ScheduleFileLabel.Content = "";
    }

    private void StatementClear_OnClick(object sender, RoutedEventArgs e)
    {
        if (IsRunning)
        {
            MessageBox.Show("В данный момент изменение невозможно");
            return;
        }

        StatementFile = null;
        StatementFileLabel.Content = "";
    }

    #endregion

    private void OpenPdf_OnClick(object sender, RoutedEventArgs e)
    {
        try
        {
            var obj = OutputFilesView.SelectedItem;
            if (obj is not WordFilePathModel filePathModel)
            {
                Console.WriteLine("Кажется, это не файл....");
                return;
            }

            if (!string.IsNullOrEmpty(filePathModel.PdfPath))
            {
                Process.Start(new ProcessStartInfo(filePathModel.PdfPath)
                {
                    UseShellExecute = true,
                });
            }
        }
        catch (Exception exception)
        {
            Console.WriteLine("Не удалось открыть файл");
        }
    }
}