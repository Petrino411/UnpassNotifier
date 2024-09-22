using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using UnpassNotifierDesktop.Classes.Extenstions;
using UnpassNotifierDesktop.Classes.Models;
using Path = System.IO.Path;

namespace UnpassNotifierDesktop;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private ObservableCollection<FilePathModel> sheduleFiles { get; } = [];
    private ObservableCollection<FilePathModel> excelFiles { get; } = [];
    private string resultsDirectory { get; }
    private string resourcesDirectory { get; }
    private static Queue<Task> tasks { get; } = new();
    private bool IsRunning = false;
    private string templatePath { get; set; }

    public MainWindow()
    {
        InitializeComponent();


        var directories = Directory.GetDirectories(Environment.CurrentDirectory);
        resourcesDirectory = directories.FirstOrDefault(x => x.Contains("Resources")) ??
                             Directory.CreateDirectory(Environment.CurrentDirectory + @"\Resources").FullName;
        resultsDirectory = directories.FirstOrDefault(x => x.Contains("Result"))
                           ?? Directory.CreateDirectory(Environment.CurrentDirectory + @"\Result").FullName;

        ExcelFilesListView.KeyDown += RemoveOnKeyDown(excelFiles);
        WordFilesListView.KeyDown += RemoveOnKeyDown(sheduleFiles);
        TemplateListView.KeyDown += RemoveOnKeyDown();
        ExcelFilesListView.MouseDoubleClick += OpenOnMouseDoubleClick;
        WordFilesListView.MouseDoubleClick += OpenOnMouseDoubleClick;
        TemplateListView.MouseDoubleClick += OpenOnMouseDoubleClick;
        OutputFiles.MouseDoubleClick += OpenOnMouseDoubleClick;

        ExcelFilesListView.ItemsSource = excelFiles;
        WordFilesListView.ItemsSource = sheduleFiles;
        
        FillSheduleFiles(Directory
            .GetFiles(resourcesDirectory + @"\Word", "*.docx", SearchOption.AllDirectories)
            .Select(x => new FilePathModel(x))
            .ToList()
        );
        FillExcelFiles(Directory
            .GetFiles(resourcesDirectory + @"\Excel", "*.xlsx", SearchOption.AllDirectories)
            .Select(x => new FilePathModel(x))
            .ToList()
        );
        FillTemplate(Directory.GetFiles(resourcesDirectory, "*.docx", SearchOption.AllDirectories)
            .FirstOrDefault(x => x.Contains("УВЕДОМЛЕНИЕ")));

    }

    private void OpenOnMouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        try
        {
            if (sender is not ListView listView) return;

            var obj = listView.SelectedItem;
            if (obj is not FilePathModel filePathModel)
            {
                Console.WriteLine("Кажется, это не файл....");
                return;
            }

            if (!string.IsNullOrEmpty(filePathModel.PdfPath))
            {
                Process.Start(new ProcessStartInfo(filePathModel.PdfPath)
                {
                    UseShellExecute = true
                });
            }
            
            Process.Start(new ProcessStartInfo(filePathModel.WordPath)
            {
                UseShellExecute = true
            });
        }
        catch (Exception exception)
        {
            Console.WriteLine("Не удалось открыть файл");
        }
    }

    private KeyEventHandler RemoveOnKeyDown(ObservableCollection<FilePathModel> collection)
    {
        return (sender, args) =>
        {
            if (sender is not ListView listView) return;
            if (args.Key is not (Key.Delete or Key.Back)) return;

            var array = new FilePathModel[listView.Items.Count];
            listView.SelectedItems.CopyTo(array, 0);
            foreach (var item in array)
            {
                collection.Remove(item);
            }
        };
    }

    private KeyEventHandler RemoveOnKeyDown()
    {
        return (sender, args) =>
        {
            if (sender is not ListView listView) return;
            if (args.Key is not (Key.Delete or Key.Back)) return;

            var array = new FilePathModel[listView.Items.Count];
            listView.SelectedItems.CopyTo(array, 0);
            foreach (var item in array)
            {
                listView.Dispatcher.Invoke(() => { listView.Items.Remove(item); });
            }
        };
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (tasks.Count(x => x.IsCompleted) < tasks.Count)
        {
            MessageBox.Show("Кажется, висит несколько задач. Подождите завершения их выполнения.");
            e.Cancel = true;
            return;
        }

        base.OnClosing(e);
    }

    private void FillSheduleFiles(List<FilePathModel>? FilePaths)
    {
        if (FilePaths == null)
        {
            sheduleFiles.Clear();
            WordFilesListView.Items.Clear();
            Console.WriteLine("Нет графиков в папке по умолчанию");
            return;
        }

        foreach (var filePathModel in FilePaths)
        {
            sheduleFiles.Add(filePathModel);
        }
    }

    private void FillExcelFiles(List<FilePathModel>? FilePaths)
    {
        if (FilePaths == null)
        {
            excelFiles.Clear();
            ExcelFilesListView.Items.Clear();
            Console.WriteLine("Нет ведомостей в папке по умолчанию");
            return;
        }

        foreach (var filePathModel in FilePaths)
        {
            excelFiles.Add(filePathModel);
        }
    }

    private void FillTemplate(string? templatePath)
    {
        if (templatePath == null)
        {
            this.templatePath = string.Empty;
            TemplateListView.Items.Clear();
            Console.WriteLine("Нет шаблона уведомлений в папке по умолчанию");
            return;
        }

        this.templatePath = templatePath;
        TemplateListView.Items.Add(new FilePathModel(templatePath));
        Console.WriteLine($"Выбран файл шаблона уведомления: {templatePath}");
    }

    private async Task WorkBody(FilePathModel excelFile)
    {
        var groupName = excelFile.FileName.Split('.').First();
        try
        {
            var wordFilePath = sheduleFiles.FirstOrDefault(x => x.WordPath.Contains(groupName));
            var disciplines = await WordExtensions.DisciplinesAttestationFill(wordFilePath.WordPath);

            var notifies = await ExcelExtensions.ExcelParse(excelFile.WordPath, disciplines);

            var targetFile = groupName + $@" - {DateTime.Now.ToShortDateString()}";
            var outputDirectory = Directory.CreateDirectory(resultsDirectory + @$"\{targetFile}").FullName;

            Console.WriteLine($"Создание Word уведомлений для {groupName}");
            await WordExtensions.WordGenerate(notifies, outputDirectory, templatePath, OutputFiles, tasks);
        }
        catch (Exception e)
        {
            Console.WriteLine($"Невозможно обработать: {groupName}. Файла либо нет, либо произошла ошибка.");
            // OutputFiles.Items.Add($"{groupName} не сформирован(пропущено)");
            return;
        }
    }

    #region Buttons

    private async void SelectShedulesBtn(object sender, RoutedEventArgs e)
    {
        var fileDialog = new OpenFileDialog
        {
#if DEBUG
            InitialDirectory = resourcesDirectory + @"\Word",
#endif
            Multiselect = true,
            Filter = "Графики аттестации|*.docx|All Files|*.*",
            Title = "Выберите файлы"
        };

        if (fileDialog.ShowDialog() != true) return;

        sheduleFiles.Clear();
        FillSheduleFiles(fileDialog.FileNames.Select(x => new FilePathModel(x)).ToList());
    }

    private async void SelectAttestationsBtn(object sender, RoutedEventArgs e)
    {
        var fileDialog = new OpenFileDialog
        {
#if DEBUG
            InitialDirectory = resourcesDirectory + @"\Excel",
#endif
            Multiselect = true,
            Filter = "Сводные ведомости|*.xlsx|All Files|*.*",
            Title = "Выберите файлы"
        };

        if (fileDialog.ShowDialog() != true) return;

        excelFiles.Clear();
        FillExcelFiles(fileDialog.FileNames.Select(x => new FilePathModel(x)).ToList());
    }

    private async void SelectTemplateBtn(object sender, RoutedEventArgs e)
    {
        var fileDialog = new OpenFileDialog
        {
#if DEBUG
            InitialDirectory = resourcesDirectory,
#endif
            Filter = "Шаблон уведомления|*.docx|All Files|*.*",
            Title = "Выберите файлы"
        };

        if (fileDialog.ShowDialog() != true) return;

        FillTemplate(fileDialog.FileName);
    }

    private async void RunBtn(object sender, RoutedEventArgs e)
    {
        if (IsRunning
            || string.IsNullOrWhiteSpace(templatePath)
            || excelFiles.Count == 0
            || sheduleFiles.Count == 0)
        {
            return;
        }

        tasks.Enqueue(Task.Run(async () =>
        {
            IsRunning = true;
            OutputFiles.Dispatcher.Invoke(() => { OutputFiles.Items.Clear(); });

            var innerTasks = new List<Task>();
            Console.WriteLine("Запуск обработки Excel файлов");
            foreach (var excelFile in excelFiles)
            {
                if (excelFile.FileName.Contains("~$")) continue;
                innerTasks.Add(Task.Run(async () => { await WorkBody(excelFile); }));
            }

            await Task.WhenAll(innerTasks);
            IsRunning = false;
            Console.WriteLine("Конец обработки.");
        }));
    }

    #endregion

    private void OutputFiles_OnMouseRightButtonDown(object sender, MouseButtonEventArgs e)
    {
        OutputFiles.ContextMenu.IsOpen = true;
    }
    
    /// <summary>
    /// Медленно, но экономим оперативку
    /// </summary>
    private async void ConvertToPDF(object sender, RoutedEventArgs e)
    {
        var selectedItems = OutputFiles.SelectedItems;
        await Task.Run(() =>
        {
            foreach (var selectedItem in selectedItems.Cast<FilePathModel>().ToList())
            {
                selectedItem.PdfPath = selectedItem.WordPath.Replace(@"\Word\", @"\PDF\") + ".pdf";
                WordExtensions.ConvertDocxToPdf(selectedItem.WordPath, selectedItem.PdfPath);
            }

            Console.WriteLine("Элементы преобразованы в pdf.");
        });
    }
    /// <summary>
    /// Быстро, но оперативе пизда
    /// </summary>
    private async void ConvertToPDFAsync(object sender, RoutedEventArgs e)
    {
        var selectedItems = OutputFiles.SelectedItems.Cast<FilePathModel>().ToList();

        var maxThreads = Environment.ProcessorCount;  
        using (var semaphore = new SemaphoreSlim(maxThreads))
        {
            var tasks = selectedItems.Select(async selectedItem =>
            {
                await semaphore.WaitAsync();  
                try
                {
                    selectedItem.PdfPath = selectedItem.WordPath.Replace(@"\Word\", @"\PDF\") + ".pdf";
                    await Task.Run(() => WordExtensions.ConvertDocxToPdf(selectedItem.WordPath, selectedItem.PdfPath));
                }
                finally
                {
                    semaphore.Release();  
                }
            });
            
            await Task.WhenAll(tasks);
        }

        Console.WriteLine("Все элементы преобразованы в PDF.");
    }
}