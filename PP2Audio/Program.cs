using Microsoft.Win32;
using System.Diagnostics;
using System.IO.Compression;

namespace PP2Audio
{
    internal class Program
    {
        static void Main(string[] args)
        {
            foreach (string arg in args)
            {
                switch (arg)
                {
#pragma warning disable CA1416 // プラットフォームの互換性を検証
                    case "-install":
                        RegistryKey registryKey = Registry.CurrentUser.CreateSubKey(@"Software\Classes\SystemFileAssociations\.pptx\Shell\PP2Audio\");
                        registryKey.SetValue(null, "PP2Audioで開く");
                        registryKey = registryKey.CreateSubKey("Command");
                        registryKey.SetValue(null, $"\"{Process.GetCurrentProcess().MainModule!.FileName}\" \"%1\"");
                        continue;
                    case "-uninstall":
                        Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\SystemFileAssociations\.pptx\Shell\PP2Audio\");
                        continue;
#pragma warning restore CA1416 // プラットフォームの互換性を検証
                }
                string ppFilePath = arg;
                if (!File.Exists(ppFilePath)) { continue; }
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
}