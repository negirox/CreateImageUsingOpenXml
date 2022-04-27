using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace InteropTest
{
    class Program
    {
        static void Main(string[] args)
        {
            //var tasks = new List<Task>();
            InteropConversion fIleGeneration = new InteropConversion();
            string templatePath = @"C:\Users\mukes\Desktop\New folder\LAW\";
            string destination = @"C:\Users\mukes\Desktop\New folder\LAW\converted";
            ////fIleGeneration.convertDOCtoPDF(@"C:\Users\mukes\Desktop\Siena\Canada\ArchivalTest\ArchivalTest1\cdactemplate.docx", @"C:\Users\mukes\Desktop\Siena\Canada\ArchivalTest\ArchivalTest1\");
            FindAndKillProcess("WINWORD");
            //fIleGeneration.convertDOCtoPDF(@"C:\Users\mukes\Desktop\New folder\LAW\test\law.docx", destination);
            //  fIleGeneration.CreatePDFusingI(@"C:\Users\mukes\Desktop\New folder\LAW\test\law.docx", destination);
            var tasks = new List<Task>();
            var timer = new Stopwatch();
            timer.Start();
            DirectoryInfo d = new DirectoryInfo(templatePath);
            FileInfo[] Files = d.GetFiles("*.docx"); //Getting Text files

            //foreach (FileInfo file in Files)
            //{
            //    string filePath = file.FullName;
            //    tasks.Add(Task.Run(() => fIleGeneration.convertDOCtoPDF(filePath, destination)));
            //}
            //for (int i = 0; i < 5; i++)
            //{
            //    tasks.Add(Task.Run(() => fIleGeneration.convertDOCtoPDF(templatePath,destination)));
            //}
            //Task.WhenAll(tasks);
            Task.WaitAll(tasks.ToArray());
            TimeSpan timeTaken = timer.Elapsed;
            string timeTakenString = "Time taken: " + timeTaken.ToString(@"m\:ss\.fff");
            Console.WriteLine("File converted " + timeTakenString);
            // ExecutePowerShell();
            Console.Read();
        }

        public static void ExecutePowerShell() {
            FindAndKillProcess("powershell");
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = @"powershell.exe",
                Arguments = @"& 'C:\Users\mukes\Desktop\New folder\LAW\wordtopdf.ps1'",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            Process process = new Process();
            process.StartInfo = startInfo;
            var timer = new Stopwatch();
            timer.Start();
            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            TimeSpan timeTaken = timer.Elapsed;
            string timeTakenString = "Time taken: " + timeTaken.ToString(@"m\:ss\.fff");
            Console.WriteLine("File converted " + timeTakenString);
            Console.WriteLine(output);
        }

        public static void FindAndKillProcess(string name)

        {

            Process[] procList = Process.GetProcessesByName(name);
            Console.WriteLine($"process found {name} : " + procList.Length);
            for (int i = procList.Length - 1; i >= 0; i--)
            {
                //ShowProcessinfo(procList[i]);
                procList[i].Kill();
                Console.WriteLine($"process killed {name}");
            }

        }

        public static void ShowProcessinfo(Process process)
        {
            // Define variables to track the peak
            // memory usage of the process.
            long peakPagedMem = 0,
                 peakWorkingSet = 0,
                 peakVirtualMem = 0;

            // Start the process.
            using (Process myProcess = process)
            {
                // Display the process statistics until
                // the user closes the program.
                do
                {
                    if (!myProcess.HasExited)
                    {
                        // Refresh the current process property values.
                        myProcess.Refresh();

                        Console.WriteLine();

                        // Display current process statistics.

                        Console.WriteLine($"{myProcess} -");
                        Console.WriteLine("-------------------------------------");

                        Console.WriteLine($"  Physical memory usage     : {myProcess.WorkingSet64}");
                        Console.WriteLine($"  Base priority             : {myProcess.BasePriority}");
                        Console.WriteLine($"  Priority class            : {myProcess.PriorityClass}");
                        Console.WriteLine($"  User processor time       : {myProcess.UserProcessorTime}");
                        Console.WriteLine($"  Privileged processor time : {myProcess.PrivilegedProcessorTime}");
                        Console.WriteLine($"  Total processor time      : {myProcess.TotalProcessorTime}");
                        Console.WriteLine($"  Paged system memory size  : {myProcess.PagedSystemMemorySize64}");
                        Console.WriteLine($"  Paged memory size         : {myProcess.PagedMemorySize64}");
                        Console.WriteLine($"  Thread Count in Process        : {myProcess.Threads.Count}");
                        // Update the values for the overall peak memory statistics.
                        peakPagedMem = myProcess.PeakPagedMemorySize64;
                        peakVirtualMem = myProcess.PeakVirtualMemorySize64;
                        peakWorkingSet = myProcess.PeakWorkingSet64;

                        if (myProcess.Responding)
                        {
                            Console.WriteLine("Status = Running");
                        }
                        else
                        {
                            Console.WriteLine("Status = Not Responding");
                        }
                    }
                }
                while (!myProcess.WaitForExit(5000));

                Console.WriteLine();
                Console.WriteLine($"  Process exit code          : {myProcess.ExitCode}");

                // Display peak memory statistics for the process.
                Console.WriteLine($"  Peak physical memory usage : {peakWorkingSet}");
                Console.WriteLine($"  Peak paged memory usage    : {peakPagedMem}");
                Console.WriteLine($"  Peak virtual memory usage  : {peakVirtualMem}");
            }
        }
    }
}
