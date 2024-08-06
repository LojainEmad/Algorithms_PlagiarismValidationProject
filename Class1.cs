using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using OfficeOpenXml;
using static PlagiarismValidationProject.Program;

namespace PlagiarismValidationProject
{
    public class TestClass
    {
        public void Test(string folderPath)
        {
            string[] files = Directory.GetFiles(folderPath, "*.xlsx");
            int filenum = 1;
            for (int i = 0; i < files.Length; i += 3)
            {
                string file = files[i];
                Console.WriteLine($"Processing file: {file}");

                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                int result = StartMyProgram(file, filenum);

                stopwatch.Stop();
                FileInfo fileInfo = new FileInfo(files[i+2]);
                int numRows = 0;
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    numRows = worksheet.Dimension.Rows;
                }
                Console.WriteLine("num of groups Expected : " + (numRows - 1) + "   output : " + result);
                Console.WriteLine($"Time taken for processing: {stopwatch.ElapsedMilliseconds} ms");
                Console.WriteLine(" ");
                Console.WriteLine(" ");
                filenum++;
            }
        }

        private int StartMyProgram(string filePath , int filenum)
        {
            Dictionary<KeyValuePair<string, string>, edgeinfo> edges = new Dictionary<KeyValuePair<string, string>, edgeinfo>();
            Dictionary<string, double> groupAvgSimilarity = new Dictionary<string, double>();
            ExcelHandler excelproceesor = new ExcelHandler();
            excelproceesor.ReadExcel(filePath, edges);
            GraphHandler h = new GraphHandler();
            h.ConstructGraph(edges);
            Stopwatch stopwatch1 = new Stopwatch();
            stopwatch1.Start();
            Dictionary<string, groupinfo> groups = h.MST();
            h.CalculateAvgSimilarity(groups);
            excelproceesor.WriteMSTFile("C:\\Users\\Hager Essam\\OneDrive\\Desktop\\Collage\\Algo\\Project\\Output", groups, filenum);
            stopwatch1.Stop();
            Console.WriteLine($"Time taken for MST and MST file: {stopwatch1.ElapsedMilliseconds} ms");
            //h.CalculateAvgSimilarity(groups);
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();
            excelproceesor.WriteStatFile("C:\\Users\\Hager Essam\\OneDrive\\Desktop\\Collage\\Algo\\Project\\Output", groups,filenum);
            stopwatch.Stop();
            Console.WriteLine($"Time taken for generating Stat file: {stopwatch.ElapsedMilliseconds} ms");
            //excelproceesor.WriteMSTFile("C:\\Users\\Hager Essam\\OneDrive\\Desktop\\Collage\\Algo\\Project\\Output", groups, filenum);
            return groups.Count();
        }
    }
}
