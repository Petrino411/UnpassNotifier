using System.IO;

namespace UnpassNotifierDesktop.Classes.Models;

public class FilePathModel
{
    public string FileName { get; }
    public string FilePath { get; }

    public FilePathModel(string filePath)
    {
        FilePath = filePath;
        FileName = Path.GetFileName(filePath);
    }

    public FileInfo GetFileInfo()
    {
        return new FileInfo(FilePath);
    }

    public override string ToString()
    {
        return FileName;
    }
}