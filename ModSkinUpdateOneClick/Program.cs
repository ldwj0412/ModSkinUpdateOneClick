using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO.Compression;
using Syroot.Windows.IO;
using System.IO;
using System.Diagnostics;
using System.Collections.ObjectModel;
using System.Management.Automation;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;
using OpenQA.Selenium.Edge;

namespace ModSkinUpdate
{
    class Program
    {

        static int initialCount;
        static Timer timer;
        //static string directoryToMon = @"C:\Downloads";
        static void Main(string[] args)
        {
            initialCount = CountZipFileInDir();
         

            web web =new web(); 

            web.directToWeb("https://leagueskin.net/p/download-mod-skin-2020-chn");
            web.clickForDownload();

            FileDownloaded();

            web.quitDriver();

            ZipAction zip = new ZipAction();
            var path = zip.unzipFile();

            SilentInstallation.ProcessFolder(path);

            Environment.Exit(0);

        }


        public static void FileDownloaded()
        {
            int newZipFileCount = CountZipFileInDir();

            Console.WriteLine(newZipFileCount + " " + initialCount);

            if(newZipFileCount == initialCount)
            {
                Thread.Sleep(2000);
                FileDownloaded();
            }
            else
            {
                return;
            }/*

            while (newZipFileCount == initialCount)
            {
                Thread.Sleep(3000);
                FileDownloaded();
            }*/

        }


        public static int CountZipFileInDir()
        {
         

            DirectoryInfo dir = new DirectoryInfo(KnownFolders.Downloads.Path);
/*            if (!dir.Exists)
            {
                Console.WriteLine("no directory");
                return 0;
            }*/
            return dir.GetFiles("*.zip").Count();


        }







        class web
        {   
            private readonly EdgeDriver driver;

            public web()
            {
                //driver = new ChromeDriver();
                

                new DriverManager().SetUpDriver(new EdgeConfig(), VersionResolveStrategy.MatchingBrowser);
                var options = new EdgeOptions();           
                options.AddArgument("--headless=new");
                var edgeDriverService = EdgeDriverService.CreateDefaultService();

                // instatntiate the driver here  
               driver = new EdgeDriver(edgeDriverService, options);
            }

            public void directToWeb(string url)
            {

                driver.Navigate().GoToUrl(url);

            }

            public void clickForDownload()
            {
                new WebDriverWait(driver, TimeSpan.FromSeconds(15)).Until(ExpectedConditions.ElementExists(By.XPath("/html/body/div/div/div/div[2]/div/div[7]/div/span/center[2]/a")));
                new WebDriverWait(driver, TimeSpan.FromSeconds(15)).Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/div/div/div/div[2]/div/div[7]/div/span/center[2]/a")));

                var download = driver.FindElement(By.XPath("/html/body/div/div/div/div[2]/div/div[7]/div/span/center[2]/a"));
                IJavaScriptExecutor Buttonexecutor = (IJavaScriptExecutor)driver;
                Buttonexecutor.ExecuteScript("arguments[0].click();", download);




            }


            public void quitDriver()
            {

               // changeInFolder();
                driver.Quit();
            }




        }








        class ZipAction
        {

            private string downloadpath = KnownFolders.Downloads.Path;
            private string desktoppath = KnownFolders.Desktop.Path;
            private string unzipFileName;
            private readonly string pattern = "MODSKIN_*";

            public string UnzipFileName
            {
                get { return unzipFileName; }
                set { unzipFileName = value; }
            }

            public string DesktopPath
            {
                get { return desktoppath; }
            }


            private string getTheLatestDownloadFilePath()
            {

                var directory = new DirectoryInfo(downloadpath);

                var myFile = directory.GetFiles()
                             .OrderByDescending(f => f.LastWriteTime)
                             .First().FullName;

                return myFile;
            }

            public string unzipFile()
            {

                FindAndDeleteFolders(desktoppath, pattern,"folder");

                var zipFilePath = getTheLatestDownloadFilePath();

                UnzipFileName = Path.GetFileNameWithoutExtension(zipFilePath);


                if (Directory.Exists(desktoppath + "/" + UnzipFileName))
                    // Ask user confirmation if needed
                    Directory.Delete(desktoppath + "/" + UnzipFileName, true);

                ZipFile.ExtractToDirectory(zipFilePath, desktoppath + "/" + UnzipFileName,true);

                FindAndDeleteFolders(downloadpath, pattern,"file");

                return desktoppath + "/" + UnzipFileName;

            }


            private static void FindAndDeleteFolders(string directory, string searchPattern,string deleteType)
            {
                try
                {
                    switch (deleteType)
                    {
                        case "file":
                            {
                                string[] files = Directory.GetFiles(directory, $"{searchPattern}");
                                foreach (string file in files)
                                {
                                    File.Delete(file);
                                    Console.WriteLine($"Deleted file: {file}");
                                }
                                break;
                            }
                        case "folder":
                            {
                                // Get all folders with names containing "MODSKIN_"
                                string[] folders = Directory.GetDirectories(directory, searchPattern, SearchOption.AllDirectories);

                                foreach (string folder in folders)
                                {
                                    // Delete the folder and its contents
                                    Directory.Delete(folder, true);
                                    Console.WriteLine($"Deleted folder: {folder}");
                                }
                                break;
                            }
                    }
                    

                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error: {e.Message}");
                }
            }

            


        }



        public class SilentInstallation
        {


            public static void ProcessFolder(string path)
            {
                string SOURCEFOLDERPATH = path;

                if (Directory.Exists(SOURCEFOLDERPATH))
                {
                    Console.WriteLine("Directory exists at: {0}", SOURCEFOLDERPATH);
                    if (Directory.GetFiles(SOURCEFOLDERPATH, "*.exe").Length > 0)
                    {
                        int count = Directory.GetFiles(SOURCEFOLDERPATH, "*.exe").Length;
                        string[] files = Directory.GetFiles(SOURCEFOLDERPATH, "*.exe");

                        foreach (var file in files)
                        {
                            var fileName =Path.GetFileName(file);
                            var fileNameWithPath = SOURCEFOLDERPATH + "\\" + fileName;
                           // Console.WriteLine("File Name: {0}", fileName);
                            //Console.WriteLine("File name with path : {0}", fileNameWithPath);
                            //Deploy application  
                            //Console.WriteLine("Wanna install {0} application on this VM? ", fileName);

                        
                            DeployApplications(fileNameWithPath);
                           
                        }
                    }

                }
                else
                    Console.WriteLine("Directory does not exist at: {0}", SOURCEFOLDERPATH);

            }





            private static void DeployApplications(string executableFilePath)
            {
                PowerShell powerShell = null;
                //Console.WriteLine(" ");
                Console.WriteLine("Deploying application...");
                try
                {
                    using (powerShell = PowerShell.Create())
                    {
                        //here “executableFilePath” need to use in place of “  
                        //'C:\\ApplicationRepository\\FileZilla_3.14.1_win64-setup.exe'”  
                        //but I am using the path directly in the script.  

                        string code = String.Format("$setup=Start-Process '{0}' -ArgumentList ' / S ' -Wait -PassThru", executableFilePath);
                        powerShell.AddScript(code);


                        Collection<PSObject> PSOutput = powerShell.Invoke(); 
                        foreach (PSObject outputItem in PSOutput)
                        {

                            if (outputItem != null)
                            {

                                Console.WriteLine(outputItem.BaseObject.GetType().FullName);
                                Console.WriteLine(outputItem.BaseObject.ToString() + "\n");
                            }
                        }

                        if (powerShell.Streams.Error.Count > 0)
                        {
                            string temp = powerShell.Streams.Error.First().ToString();
                            Console.WriteLine("Error: {0}", temp);

                        }
                        else
                            Console.WriteLine("Installation has completed successfully.");

                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error occured: {0}", ex.InnerException);
                    //throw;  
                }
                finally
                {
                    if (powerShell != null)
                        powerShell.Dispose();
                }

                Environment.Exit(0);

            }
        }







    }



}







    

