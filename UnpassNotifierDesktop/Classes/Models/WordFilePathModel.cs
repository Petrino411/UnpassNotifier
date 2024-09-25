using System.IO;

namespace UnpassNotifierDesktop.Classes.Models;

public class WordFilePathModel
{
    public string WordPath { get; }
    public string FileName { get; }
    public string PdfPath { get; set; } = string.Empty;

    public WordFilePathModel(string wordPath)
    {
        WordPath = wordPath;
        FileName = Path.GetFileName(wordPath);
    }

    public WordFilePathModel(string wordPath, string pdfPath) : this(wordPath)
    {
        PdfPath = pdfPath;
    }

    public override string ToString()
    {
        return string.IsNullOrEmpty(PdfPath) ? FileName : $"{FileName} + PDF";
    }
}