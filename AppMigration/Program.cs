using System;
using System.Diagnostics;
using System.IO;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Management.Automation;
using System.Net;
using System.Threading;

namespace AppMigration
{
    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse); 
    public class Package
    {
        public string PackageIdentifier { get; set; }
    }

    public class SourceDetails
    {
        public string Argument { get; set; }
        public string Identifier { get; set; }
        public string Name { get; set; }
        public string Type { get; set; }
    }

    public class Source
    {
        public List<Package> Packages { get; set; }
        public SourceDetails SourceDetails { get; set; }
    }

    public class WingetPackages
    {
        [JsonProperty("$schema")]
        public string Schema { get; set; }
        public DateTime CreationDate { get; set; }
        public List<Source> Sources { get; set; }
        public string WinGetVersion { get; set; }
    }


    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Welcome to the Application Migration Tool");
            Console.WriteLine($"{Environment.NewLine}Built on WinGet");
            Console.WriteLine($"{Environment.NewLine}This tool tries to make setting up a new machine easier by backing up your installed applications to Box");
            Console.WriteLine("");
//          Console.ResetColor();
            while (!EnsureWingetInstalled())
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Winget needs to be installed in order to proceed.");
                Console.ResetColor();
            }
            bool showMenu = true;
            while (showMenu)
            {
                showMenu = MainMenu();
            }
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Press any key to exit.");
            Console.ResetColor();
            Console.ReadKey(true);
        }

        public static bool EnsureWingetInstalled()
        {
            string version = "v1.2.10271";
            if (!ExistsOnPath("winget.exe"))
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Do you want to install WinGet now?");
                Console.WriteLine("[Y] Yes  [N] No  [X] Exit");
                Console.ResetColor();
                switch (Console.ReadLine().ToLower())
                {
                    case "y":
                        return InstallWinget(version);
                    case "n":
                        return false;
                    case "x":
                        System.Environment.Exit(1607);
                        return false;
                    default:
                        return false;
                }
            }
            return true;
        }

        public static bool InstallWinget(string version)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Installing Winget...");
            WebClient webClient = new WebClient();
            string filename = "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle";
            string remoteUri = "https://github.com/microsoft/winget-cli/releases/download/" + version + "/" + filename;
            webClient.DownloadFile(remoteUri,filename);
            Process p = new Process();
            p.StartInfo.FileName = "powershell.exe";
            p.StartInfo.Arguments = $@"Add-AppPackage -Path {filename} -InstallAllResources";
            p.Start();
            p.WaitForExit(120000);
            if (ExistsOnPath("winget.exe"))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Winget is installed!");
                Console.ResetColor();
                return true;
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Winget failed to install :(");
                Console.ResetColor();
                return false;
            }
            
        }

        public static bool ExistsOnPath(string fileName)
        //Taken from https://stackoverflow.com/questions/3855956/check-if-an-executable-exists-in-the-windows-path
        {
            return GetFullPath(fileName) != null;
        }

        public static string GetFullPath(string fileName)
        //Taken from https://stackoverflow.com/questions/3855956/check-if-an-executable-exists-in-the-windows-path
        {
            if (File.Exists(fileName))
                return Path.GetFullPath(fileName);

            var values = Environment.GetEnvironmentVariable("PATH");
            foreach (var path in values.Split(Path.PathSeparator))
            {
                var fullPath = Path.Combine(path, fileName);
                if (File.Exists(fullPath))
                    return fullPath;
            }
            return null;
        }

        public static bool MainMenu()
        //Taken from https://wellsb.com/csharp/beginners/create-menu-csharp-console-application
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"Main Menu{Environment.NewLine}");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("Select one of the following options:");
            Console.WriteLine("(1) Export your application list to Box Drive");
            Console.WriteLine("(2) Reinstall your applications");
            Console.WriteLine("(3) Exit");
            Console.Write("\r\nSelect an option: ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            switch (Console.ReadLine())
            {
                case "1":
                    return ExportApps();
                case "2":
                    return ImportApps();
                case "3":
                    System.Environment.Exit(0);
                    return false;
                default:
                    return true;
            }
        }

        public static bool ExportApps()
        {
            // Need to make sure Box exists
            // If Box exists, make sure folder exists
            // If folder already exists, asks if we can save the output there
            // Run winget and save the output
            if (!BoxDriveExists())
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Box Drive is not installed.  Please install Box Drive and try again.");
                Console.WriteLine($"{Environment.NewLine}Press any key to continue...{Environment.NewLine}");
                Console.ResetColor();
                Console.ReadKey(true);
                System.Environment.Exit(1607);
                return true;
            }
            else
            {
                string folder = ExportFolderExists();
                BackupAppList(folder);
                //Console.ForegroundColor = ConsoleColor.Green;
                //Console.WriteLine($"All winget installable packages have been exported to Box.{Environment.NewLine}{Environment.NewLine}");
                //Console.ResetColor();
                return true;
            }
        }

        

        public static bool BoxDriveExists()
        {
            string[] paths = { Environment.GetEnvironmentVariable("USERPROFILE"), "Box"};
            string fullpath = Path.Combine(paths);
            //Console.WriteLine(fullpath);
            return Directory.Exists(fullpath);
        }
        
        public static string ExportFolderExists()
        {
            string backupfolder = "winget-export";
            string[] paths = { Environment.GetEnvironmentVariable("USERPROFILE"), "Box", backupfolder };
            string fullpath = Path.Combine(paths);

            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("We are about to backup your application files to Box.");
            Console.WriteLine($"They will be uploaded to the following folder in Box:");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write($"      { backupfolder}");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"{Environment.NewLine}But don't worry about remembering that.  As long as you are signed into Box on your new PC this tool will find them automatically.");
            Console.Write($"{Environment.NewLine}The applications below in ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("YELLOW");
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.Write(" cannot be backed up by winget and may need to be manually reinstalled.");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"{Environment.NewLine}Press any key to continue...{Environment.NewLine}");
            Console.ResetColor();
            Console.ReadKey(true);

            DirectoryInfo di = Directory.CreateDirectory(fullpath);
            return di.FullName;
        }

        public static void BackupAppList(string path)
        {
            string f = "winget-export.json";
            string fullpath = Path.Combine(path, f);
            Process p = new Process();
            p.StartInfo.FileName = "winget.exe";
            p.StartInfo.Arguments = $@"export -o {fullpath} --accept-source-agreements";
            p.Start();
            p.WaitForExit(120000);
            WingetPackages wingetPackages = ReadJson(fullpath);
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"{Environment.NewLine}The following packages were backed up:");
            int count = 0;
            foreach (var source in wingetPackages.Sources)
            {
                foreach (var package in source.Packages)
                {
                    Console.WriteLine(package.PackageIdentifier);
                    count = count + 1;
                    Thread.Sleep(50);
                }
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"{Environment.NewLine}{count} Total Packages were backed up and can be reinstalled with WinGet.{Environment.NewLine}{Environment.NewLine}");
            Console.ResetColor();
        }

        public static WingetPackages ReadJson(string path)
        {
            StreamReader r = new StreamReader(path);
            string jsonString = r.ReadToEnd();
            var myJsonObject = JsonConvert.DeserializeObject<WingetPackages>(jsonString);            
            return myJsonObject;
        }

        public static bool ImportApps()
        {
            // First ensure Box is installed 
            // Second ensure the user signed into Box
            // Get the file
            // Install all the packages one by one
            // if everything is successful return false to close out the menu
            if (EnsureBoxInstalled())
            {
                string backupfolder = "winget-export";
                string[] paths = { Environment.GetEnvironmentVariable("USERPROFILE"), "Box", backupfolder };
                string path = Path.Combine(paths);
                string f = "winget-export.json";
                string fullpath = Path.Combine(path, f);
                WingetPackages wingetPackages = ReadJson(fullpath);
                InstallPackagesWithWinget(wingetPackages);
                return true;
            }
            return true;
        }

        

        public static bool EnsureBoxInstalled()
        {
            // Download Box
            if (!BoxDriveExists())
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Installing Box Drive...");
                WebClient webClient = new WebClient();
                string filename = "Box-x64.msi";
                string remoteUri = "https://e3.boxcdn.net/box-installers/desktop/releases/win/Box-x64.msi";
                webClient.DownloadFile(remoteUri, filename);
                // Install Box
                Process p = new Process();
                p.StartInfo.FileName = "msiexec.exe";
                p.StartInfo.Arguments = $@"/i {filename}";
                p.Start();
                p.WaitForExit(120000);
                Console.WriteLine("Box is installed!  You need to sign in first to get your backup file.");
                Console.WriteLine("Press any key to continue after signing into Box...");
                Console.ReadKey(true);
                return true;
            }
            return true;
        }

        public static void InstallPackagesWithWinget(WingetPackages wingetPackages)
        {
            int count = 0;
            int failures = 0;
            foreach (var source in wingetPackages.Sources)
            {
                foreach (var package in source.Packages)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"{Environment.NewLine}Installing application {package.PackageIdentifier} using winget");
                    Process p = new Process();
                    p.StartInfo.FileName = "winget.exe";
                    p.StartInfo.Arguments = $@"install {package.PackageIdentifier} -s {source.SourceDetails.Name}";
                    p.Start();
                    p.WaitForExit(600000);
                    if (p.ExitCode == 0)
                    {
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"{package.PackageIdentifier} successfully installed!");
                        Console.ResetColor();
                        count = count + 1;
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"{package.PackageIdentifier} failed to install :(");
                        Console.ResetColor();
                        failures = failures + 1;
                    }
                }
            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"{count} Total Packages were reinstalled with WinGet!");
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"{failures} Total Packages failed to get installed.");
            Console.ResetColor();
        }
    }
}
