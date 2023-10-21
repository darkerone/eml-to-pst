using Aspose.Email;
using Aspose.Email.Mapi;
using Aspose.Email.Storage.Pst;

namespace ConsoleApp1
{
    /// <summary>
    /// Permet de transformer tous les fichiers eml des sous dossiers (de premier niveau) en un fichier pst.
    /// Il faut définir les chemins d'entrée et de sortie en dur dans le code
    /// </summary>
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, World!");

            var dirPath = "D:\\Franck\\Mails\\Storage Folders\\CRE";
            var outFileName = $"D:\\Franck\\Mails\\CRE.pst";

            if (File.Exists(outFileName))
                File.Delete(outFileName);

            using (var personalStorage = PersonalStorage.Create(outFileName, FileFormatVersion.Unicode))
            {
                var directories = Directory.GetDirectories(dirPath);
                if (directories.Any())
                {
                    for (int i = 0; i < directories.Length; i++)
                    {
                        SetFilesIntoBox(personalStorage, directories[i], i, directories.Length);
                    }
                        
                }
                else
                    SetFilesIntoBox(personalStorage, dirPath, 0, 1);
            }
        }

        static void SetFilesIntoBox(PersonalStorage personalStorage, string directoryPath, int currentDirIndex, int dirCount)
        {
            FolderInfo pathBox = personalStorage.RootFolder.GetSubFolder(Path.GetFileName(directoryPath));
            if(pathBox == null)
            {
                pathBox = personalStorage.RootFolder.AddSubFolder(Path.GetFileName(directoryPath));
            }

            Console.WriteLine($"Create box: {pathBox.DisplayName}");
            var files = Directory.GetFiles(directoryPath, "*.eml");
            int filesCount = files.Length;
            int currentFileIndex = 0;
            foreach (var f in files)
            {
                using (var message = MailMessage.Load(f))
                {
                    pathBox.AddMessage(MapiMessage.FromMailMessage(message, MapiConversionOptions.UnicodeFormat));
                    Console.WriteLine($"{currentDirIndex + 1}/{dirCount} - {currentFileIndex + 1}/{filesCount} : {Path.GetFileName(f)}");
                }
                currentFileIndex++;
            }
        }

    }
}



