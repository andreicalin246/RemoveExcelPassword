using System.Diagnostics;
using System.IO.Compression;
using System.Text.RegularExpressions;

namespace RemoveExcelPassword
{
    internal static class PasswordRemover
    {
        public static void Run()
        {
            var excelFilePath = GetFilePath();

            if (IsValidFilePath(excelFilePath))
            {
                var paths = UnpackExcel(excelFilePath);

                RemovePassword(paths.DestinationDirectory);

                PackExcel(excelFilePath, paths.DestinationDirectory, paths.PathRoot);
            }
        }

        private static void PackExcel(string? excelFilePath, string? destinationDirectory, string? pathRoot)
        {
            var fileName = Path.GetFileNameWithoutExtension(excelFilePath);
            var fileExtension = Path.GetExtension(excelFilePath);
            var outputFile = $"{pathRoot}\\{fileName}-unprotected{fileExtension}";
            
            if (File.Exists(outputFile)) 
                File.Delete(outputFile);
            ZipFile.CreateFromDirectory(destinationDirectory, outputFile);
            Directory.Delete(destinationDirectory, true);
        }

        private static void RemovePassword(string? destinationDirectory)
        {
            var xmlFiles = Directory.GetFiles(destinationDirectory, "*.xml", SearchOption.AllDirectories);
            foreach (var xmlFile in xmlFiles)
            {
                var fileText = File.ReadAllText(xmlFile);
                fileText = Regex.Replace(fileText, "<sheetProtection[\\s\\S]*?\\/>", string.Empty);
                File.WriteAllText(xmlFile, fileText);
            }
        }

        private static (string? PathRoot, string? DestinationDirectory) UnpackExcel(string? excelFilePath)
        {
            var pathRoot = Path.GetDirectoryName(excelFilePath);
            var destinationDirectory = $"{pathRoot}\\{Guid.NewGuid()}";
            ZipFile.ExtractToDirectory(excelFilePath, destinationDirectory);

            return (pathRoot, destinationDirectory);
        }

        private static bool IsValidFilePath(string? excelFilePath)
        {
            if (excelFilePath == null)
            {
                Console.WriteLine("File path null!");
                return false;
            }

            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine("File not found!");
                return false;
            }

            if (string.Compare(Path.GetExtension(excelFilePath), ".xlsx", true) != 0)
            {
                Console.WriteLine("Unsupported file type!");
                return false;
            }

            return true;
        }

        private static string? GetFilePath()
        {
            Console.Write("Excel File Path:");
            return Console.ReadLine();
        }
    }
}
