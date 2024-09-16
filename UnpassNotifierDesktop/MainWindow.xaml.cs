using System.ComponentModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using UnpassNotifierDesktop.Classes.Extenstions;

namespace UnpassNotifierDesktop;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    private string templatePath { get; set; }
    private List<string> sheduleFiles { get; set; } = [];
    private List<string> excelFiles { get; set; } = [];
    private string resultsDirectory { get; }
    private string resourcesDirectory { get; }
    private Queue<Task> tasks { get; set; } = new Queue<Task>();

    public MainWindow()
    {
        InitializeComponent();

        var directories = Directory.GetDirectories(Environment.CurrentDirectory);
        resourcesDirectory = directories.FirstOrDefault(x => x.Contains("Resources")) ??
                             Directory.CreateDirectory(Environment.CurrentDirectory + @"\Resources").FullName;
        resultsDirectory = directories.FirstOrDefault(x => x.Contains("Result"))
                           ?? Directory.CreateDirectory(Environment.CurrentDirectory + @"\Result").FullName;

        ExcelFiles.KeyDown += RemoveOnKeyDown();
        WordFiles.KeyDown += RemoveOnKeyDown();
        ExcelFiles.MouseDoubleClick += (o, args) => { ExcelFiles.Items.Remove((o as ListView).SelectedItem); };
        WordFiles.MouseDoubleClick += (o, args) => { WordFiles.Items.Remove((o as ListView).SelectedItem); };
        TemplateBox.MouseDoubleClick += (sender, args) => { TemplateBox.Text = ""; };
    }

    private KeyEventHandler RemoveOnKeyDown()
    {
        return (sender, args) =>
        {
            if (sender is ListView listView)
            {
                if (args.Key is Key.Delete or Key.Back)
                {
                    foreach (var item in ExcelFiles.SelectedItems)
                        listView.Items.Remove(item);
                }
            }
        };
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        Task.WaitAll(tasks.ToArray());
        
        base.OnClosing(e);
    }


    private void SelectShedulesBtn(object sender, RoutedEventArgs e)
    {
        var fileDialog = new OpenFileDialog
        {
            Multiselect = true,
            Filter = "Графики аттестации|*.docx|All Files|*.*",
            Title = "Выберите файлы"
        };

        if (fileDialog.ShowDialog() == true)
        {
            sheduleFiles = fileDialog.FileNames.ToList();
            foreach (var sheduleFile in sheduleFiles)
            {
                WordFiles.Items.Add(sheduleFile);
                Console.WriteLine($"Выбран файл графика: {sheduleFile}");
            }
        }
    }

    private void SelectAttestationsBtn(object sender, RoutedEventArgs e)
    {
        var fileDialog = new OpenFileDialog
        {
            Multiselect = true,
            Filter = "Сводные ведомости|*.xlsx|All Files|*.*",
            Title = "Выберите файлы"
        };

        if (fileDialog.ShowDialog() == true)
        {
            excelFiles = fileDialog.FileNames.ToList();
            foreach (var excelFile in excelFiles)
            {
                ExcelFiles.Items.Add(excelFile);
                Console.WriteLine($"Выбран файл ведомости: {excelFile}");
            }
        }
    }

    private void SelectTemplateBtn(object sender, RoutedEventArgs e)
    {
        var fileDialog = new OpenFileDialog
        {
            Filter = "Шаблон уведомления|*.docx|All Files|*.*",
            Title = "Выберите файлы"
        };

        if (fileDialog.ShowDialog() == true)
        {
            templatePath = fileDialog.FileName;
            TemplateBox.Text = templatePath;
            Console.WriteLine($"Выбран файл шаблона уведомления: {templatePath}");
        }
    }

    private async void RunBtn(object sender, RoutedEventArgs e)
    {
        Console.WriteLine("Запуск обработки Excel файлов");
        foreach (var excelFile in excelFiles)
        {
            if (excelFile.Contains("~$")) continue;
            tasks.Enqueue(Task.Run(async () =>
            {
                var groupName = excelFile.Split(@"\").Last().Split('.').First();

                var wordFilePath = sheduleFiles.FirstOrDefault(x => x.Contains(groupName));
                var disciplines = await WordExtensions.DisciplinesAttestationFill(wordFilePath);
                if (disciplines == null)
                {
                    Console.WriteLine(
                        $"Программа не смогла прочесть данные из графика: {wordFilePath ?? groupName}. Файла либо нет, либо произошла ошибка.");
                    return;
                }

                var notifies = await ExcelExtensions.ExcelParse(excelFile, disciplines);

                var targetFile = groupName + $@" - {DateTime.Now.ToShortDateString()}";
                var outputDirectory = Directory.CreateDirectory(resultsDirectory + @$"\{targetFile}").FullName;

                Console.WriteLine($"Создание Word уведомлений для {groupName}");
                await WordExtensions.WordGenerate(notifies, outputDirectory, templatePath);
                return;
            }));
        }

        Console.WriteLine("Конец работы.");
        return;
    }
}