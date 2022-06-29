﻿using System.IO.Compression;

namespace PP2Audio
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0) { return; }
            string ppFilePath = args[0];
            if (ppFilePath == null) { return; }
            if (!File.Exists(ppFilePath)) { return; }
            string zipFilePath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(ppFilePath) + ".zip");
            string zipFolderPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(ppFilePath));
            string mediaFolderPath = zipFolderPath + @"\ppt\media";
            string audioFolderPath = Path.Combine(Path.GetDirectoryName(ppFilePath)!, Path.GetFileNameWithoutExtension(ppFilePath));
            if (File.Exists(zipFilePath)) { File.Delete(zipFilePath); }
            if (Directory.Exists(zipFolderPath)) { Directory.Delete(zipFolderPath, true); }
            File.Copy(ppFilePath, zipFilePath);
            ZipFile.ExtractToDirectory(zipFilePath, zipFolderPath);
            string[] audioFilePaths = Directory.GetFiles(mediaFolderPath, "*.m4a", SearchOption.AllDirectories);
            Array.Sort(audioFilePaths);
            if (!Directory.Exists(audioFolderPath)) { Directory.CreateDirectory(audioFolderPath); }
            audioFilePaths.ToList().ForEach(x => File.Move(x, Path.Combine(audioFolderPath, Path.GetFileName(x))));
        }
    }
}