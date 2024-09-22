using System.IO;

namespace UnpassNotifierDesktop.Classes.Models;

public class FilePathModel
{
    public string WordPath { get; }
    public string FileName { get; }
    public string PdfPath { get; set; } = string.Empty;

    public FilePathModel(string wordPath)
    {
        WordPath = wordPath;
        FileName = Path.GetFileName(wordPath);
    }

    public FilePathModel(string wordPath, string pdfPath) : this(wordPath)
    {
        PdfPath = pdfPath;
    }

    public override string ToString()
    {
        return string.IsNullOrEmpty(PdfPath) ? FileName : $"{FileName} + PDF";
    }
}