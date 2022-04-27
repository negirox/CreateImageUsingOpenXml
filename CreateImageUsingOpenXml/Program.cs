using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateImageUsingOpenXml
{
    class Program
    {
        static void Main(string[] args)
        {
            //ConvertHtmlToPDF convertHtmlToPDF = new ConvertHtmlToPDF(@"C:\Users\mukes\Desktop\wordtemplate.docx");
            //convertHtmlToPDF.ConvertFIle();
            var tasks = new List<Task>();
            FIleGeneration fIleGeneration = new FIleGeneration();
            for (int i = 0; i < 10; i++)
            {
                tasks.Add(Task.Run(() => fIleGeneration.CreatePDFusingI(@"C:\Users\mukes\Desktop\wordtemplate.docx", @"C:\Users\mukes\Desktop\")));
            }
            //Task.WhenAll(tasks);
            Task.WaitAll(tasks.ToArray());
            Console.WriteLine("File converted");
            Console.Read();
        }

        private static void FileGeneration()
        {
            FIleGeneration fIleGeneration = new FIleGeneration();
            fIleGeneration.GenerateFile();
            fIleGeneration.Generate();
            List<Emp> emps = new List<Emp>();
            emps.Add(new Emp { Name = "Mukesh", Age = "30", Gender = "Male" });
            emps.Add(new Emp { Name = "Mukesh1", Age = "31", Gender = "Male" });
            emps.Add(new Emp { Name = "Mukesh2", Age = "32", Gender = "Male" });
            emps.Add(new Emp { Name = "Mukesh3", Age = "33", Gender = "Male" });
            var empDict = emps.Select(x =>
            {
                return new { key = x.Name, value = x.Age };
            });
            var empD = emps.ToDictionary(x => x.Name);
            Console.WriteLine(JsonConvert.SerializeObject(empDict));
            Console.WriteLine(JsonConvert.SerializeObject(empD));
            foreach (var item in empDict)
            {
                Console.WriteLine($"key => {item.key} value = {item.value}");
            }
        }
    }
}
