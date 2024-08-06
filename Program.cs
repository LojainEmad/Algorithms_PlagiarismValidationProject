using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Runtime.Remoting.Messaging;
using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;
using System.Windows.Media;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using static PlagiarismValidationProject.Program;
using static PlagiarismValidationProject.Program.ExcelHandler;

namespace PlagiarismValidationProject
{
    public class Program
    {
        public class edgeinfo
        {
            public int similarityv1;
            public int similarityv2;
            public string v1url;
            public string v2url;
            public int maxsim;
            public int linesmatched;
            public string groupparent;
            public string v1;
            public string v2;
            public edgeinfo(int sim1, int sim2, int lines, string v1, string v2,string hyperlink1, string hyperlink2, string parent)
            {
                similarityv1 = sim1;
                similarityv2 = sim2;
                linesmatched = lines;
                this.v1 = v1;
                this.v2 = v2;
                v1url = hyperlink1;
                v2url = hyperlink2;
                maxsim = Math.Max(sim1, sim2);
                groupparent = parent;
            }
        }
        public class groupinfo
        {
            public HashSet<string> vertices;
            public List<edgeinfo> edges;
            public double avgsimilarity;
            public int edgescount;
            public groupinfo()
            {
                vertices = new HashSet<string>();
                edges = new List<edgeinfo>();
                avgsimilarity = 0;
                edgescount = 0;
            }
        }
        public class DSU
        {
            public Dictionary<string, string> parent;
            private Dictionary<string, int> size;

            public DSU()
            {
                parent = new Dictionary<string, string>();
                size = new Dictionary<string, int>();
            }
            public void MakeSet(string id)
            {
                if (!parent.ContainsKey(id))
                {
                    parent[id] = id;
                    size[id] = 1;
                }
            }
            public string FindSet(string id)
            {
                if (parent[id] != id)
                {
                    parent[id] = FindSet(parent[id]);
                }
                return parent[id];
            }

            public void Union(string id1, string id2)
            {
                string root1 = FindSet(id1);
                string root2 = FindSet(id2);

                if (root1 == root2)
                    return;

                if (size[root1] < size[root2])
                {
                    parent[root1] = root2;
                    size[root2] += size[root1];
                }
                else
                {
                    parent[root2] = root1;
                    size[root1] += size[root2];
                }
            }
        }

        public class ExcelHandler
        {
            public static int ExtractFileNumber(string input) //O(k)
            {
                Match match = Regex.Match(input, @"\d+");
                return Int32.Parse(match.Value);
            }
            public void ReadExcel(string filePath, Dictionary<KeyValuePair<string, string>, edgeinfo> edges) //O(N)
            {
                using (var excelPackage = new ExcelPackage(new System.IO.FileInfo(filePath)))
                {
                    var worksheet = excelPackage.Workbook.Worksheets[0];

                    KeyValuePair<string, string> k1;
                    ExcelRange file1Cell, file2Cell;
                    string file1Data, file2Data, file1name, file2name, file1Url, file2Url;
                    int openParenIndex1, closeParenIndex1, openParenIndex2, closeParenIndex2, file1percentInt, file2percentInt, linesMatchedData, row = 2;
                    while (!string.IsNullOrEmpty(worksheet.Cells[row, 1].Text))
                    {
                        file1Cell = worksheet.Cells[row, 1]; // File 1 data
                        file1Data = file1Cell.Text;
                        if (file1Cell.Hyperlink != null && file1Cell.Hyperlink.AbsoluteUri != null)
                            file1Url = file1Cell.Hyperlink.AbsoluteUri.ToString();  //O(k)
                        else
                            file1Url = null;
                        openParenIndex1 = file1Data.IndexOf('('); //O(k)
                        closeParenIndex1 = file1Data.IndexOf(')'); //O(k)
                        file1percentInt = Int32.Parse(file1Data.Substring(openParenIndex1 + 1, closeParenIndex1 - openParenIndex1 - 2)); //O(1)
                        file1name = file1Data.Replace(file1Data.Substring(openParenIndex1, file1Data.Length - openParenIndex1), ""); //O(k)

                        file2Cell = worksheet.Cells[row, 2]; // File 2 data
                        file2Data = file2Cell.Text;
                        if (file2Cell.Hyperlink != null && file2Cell.Hyperlink.AbsoluteUri != null)
                            file2Url = file2Cell.Hyperlink.AbsoluteUri.ToString();
                        else
                            file2Url = null;
                        openParenIndex2 = file2Data.IndexOf('(');//O(k)
                        closeParenIndex2 = file2Data.IndexOf(')');//O(k)
                        file2percentInt = Int32.Parse(file2Data.Substring(openParenIndex2 + 1, closeParenIndex2 - openParenIndex2 - 2));//O(1)
                        file2name = file2Data.Replace(file2Data.Substring(openParenIndex2, file2Data.Length - openParenIndex2), "");//O(k)

                        k1 = new KeyValuePair<string, string>(file1name, file2name);

                        linesMatchedData = (int)double.Parse(worksheet.Cells[row, 3].Text.ToString()); //O(k)
                        edges[k1] = new edgeinfo(file1percentInt, file2percentInt, linesMatchedData, file1name, file2name,file1Url,file2Url, " ");
                        row++;
                    }
                }
            }
            public void WriteStatFile(string folderPath, Dictionary<string, groupinfo> groups, int filenum)
            {
                using (var excelPackage = new ExcelPackage())
                {

                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Components");
                    List<KeyValuePair<string, groupinfo>> sortedgroups = groups.OrderByDescending(x => x.Value.avgsimilarity).ToList(); // OrderByDescending O(g log g)  Tolist O(g)
                    groups = sortedgroups.ToDictionary(x => x.Key, x => x.Value); // O(g)
                    //groups = groups.OrderByDescending(x => x.Value.avgsimilarity).ToDictionary(x => x.Key, x => x.Value); // OrderByDescending O(g log g)  Tolist O(g)
                    List<int> vertsnum = new List<int>();
                    worksheet.Cells[1, 1].Value = "Component Index";
                    worksheet.Cells[1, 2].Value = "Vertices";
                    worksheet.Cells[1, 3].Value = "Average Similarity";
                    worksheet.Cells[1, 4].Value = "Component Count";

                    int row = 2;
                    //int gid = 1;
                    foreach (var group in groups) // #iter = g 
                    {
                        foreach (var vertix in group.Value.vertices) //#iter = M , Body O(k)
                            vertsnum.Add(ExtractFileNumber(vertix));
                        vertsnum.Sort(); //O(M log M)

                        worksheet.Cells[row, 1].Value = row-1;
                        worksheet.Cells[row, 2].Value = string.Join(", ", vertsnum); //O(M × k)
                        worksheet.Cells[row, 3].Value = group.Value.avgsimilarity;
                        worksheet.Cells[row, 4].Value = group.Value.vertices.Count;
                        row++;
                        vertsnum.Clear(); //O(M)
                    }
                    worksheet.Cells.AutoFitColumns(); //O(l) l is the num of characters in each column   O(number of columns × num of characters in each column)
                    string filePath = Path.Combine(folderPath, filenum + "-StatFileO" + ".xlsx"); //O(p), where 'p' is the total length of the combined paths.
                    FileInfo excelFile = new FileInfo(filePath);
                    excelPackage.SaveAs(excelFile); //O(s), where 's' is the total size of the data being written.
                }
            }
            public void WriteMSTFile(string folderPath, Dictionary<string, groupinfo> groups, int filenum)
            {
                using (var excelPackage = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("MST"); //O(1)
                    List<KeyValuePair<string, groupinfo>> sortedKeyValuePairs = groups.OrderByDescending(x => x.Value.avgsimilarity).ToList();
                    groups = sortedKeyValuePairs.ToDictionary(x => x.Key, x => x.Value);
                    //groups = groups.OrderByDescending(x => x.Value.avgsimilarity).ToDictionary(x => x.Key, x => x.Value); // OrderByDescending O(g log g)  Tolist O(g)
                    worksheet.Cells[1, 1].Value = "File 1";
                    worksheet.Cells[1, 2].Value = "File 2";
                    worksheet.Cells[1, 3].Value = "Lines Matched";
                    int row = 2;
                    string StyleName = "UrlStyle";
                    ExcelNamedStyleXml HyperStyle = worksheet.Workbook.Styles.CreateNamedStyle(StyleName);
                    HyperStyle.Style.Font.UnderLine = true;
                    HyperStyle.Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                    foreach (var group in groups) // #iter = g
                    {
                        group.Value.edges = group.Value.edges.OrderByDescending(x => x.linesmatched).ThenByDescending(x => x.maxsim).ToList(); //O(N log N) + O(N log N) + O(N)

                        foreach (var edge in group.Value.edges) // O(N)
                        {
                            if(edge.v1url != null)
                                worksheet.Cells[row, 1].Hyperlink = new Uri(edge.v1url);
                            if (edge.v2url != null)
                                worksheet.Cells[row, 2].Hyperlink = new Uri(edge.v2url);

                            worksheet.Cells[row, 1].Value = edge.v1 + " (" + edge.similarityv1 + "%)";
                            worksheet.Cells[row, 2].Value = edge.v2 + " (" + edge.similarityv2 + "%)";

                            worksheet.Cells[row, 1].StyleName = StyleName;
                            worksheet.Cells[row, 2].StyleName = StyleName;

                            worksheet.Cells[row, 3].Value = edge.linesmatched;
                            row++;
                        }
                    }
                    worksheet.Cells.AutoFitColumns(); //O(l) l is the num of characters in each column   O(number of columns × num of characters in each column)
                    string filePath = Path.Combine(folderPath, filenum + "-mstFileO" + ".xlsx"); //O(p), where 'p' is the total length of the combined paths.
                    FileInfo excelFile = new FileInfo(filePath); 
                    excelPackage.SaveAs(excelFile); //O(s), where 's' is the total size of the data being written.
                }
            }
        }

        public class GraphHandler
        {
            private Dictionary<string, List<edgeinfo>> graph = new Dictionary<string, List<edgeinfo>>();
            public void ConstructGraph(Dictionary<KeyValuePair<string, string>, edgeinfo> edges)
            {
                DSU dsu = new DSU(); // O(1))

                foreach (var edge in edges) //#iter = N   O(1)
                {
                    if (!dsu.parent.ContainsKey(edge.Key.Key))
                        dsu.MakeSet(edge.Key.Key);
                    if (!dsu.parent.ContainsKey(edge.Key.Value))
                        dsu.MakeSet(edge.Key.Value);
                }
                foreach (var edge in edges) //#iter = N
                {
                    dsu.Union(edge.Value.v1, edge.Value.v2);
                    edge.Value.groupparent = dsu.FindSet(edge.Value.v1);
                }
                foreach (var edge in edges)//#iter = N
                {
                    string groupParent = dsu.FindSet(edge.Value.v1);
                    if (!graph.ContainsKey(groupParent))
                        graph[groupParent] = new List<edgeinfo>();
                    graph[groupParent].Add(new edgeinfo(edge.Value.similarityv1, edge.Value.similarityv2, 
                        edge.Value.linesmatched, edge.Key.Key, edge.Key.Value,edge.Value.v1url,edge.Value.v2url, edge.Value.groupparent));
                }
            }
            public Dictionary<string, groupinfo> MST()
            {
                DSU dsu = new DSU();
                Dictionary<string, groupinfo> groups = new Dictionary<string, groupinfo>();
                groupinfo currentGroup = null;
                foreach (var graphid in graph.Keys) //#iter = g
                {
                    List<edgeinfo> SortedEdges = graph[graphid].OrderByDescending(x => x.maxsim).ThenByDescending(x => x.linesmatched).ToList(); // OrderByDescending O(N log N)  Tolist O(N)

                    foreach (var edge in SortedEdges) //#iter = N  O(1)
                    {
                        dsu.MakeSet(edge.v1);
                        dsu.MakeSet(edge.v2);
                    }
                    currentGroup = new groupinfo(); //O(1)
                    foreach (var edge in SortedEdges) //#iter = N
                    {
                        var root1 = dsu.FindSet(edge.v1);
                        var root2 = dsu.FindSet(edge.v2);
                        if (root1 != root2)
                        {
                            dsu.Union(edge.v1, edge.v2);
                            currentGroup.edges.Add(new edgeinfo(edge.similarityv1, edge.similarityv2, edge.linesmatched, edge.v1, edge.v2, edge.v1url, edge.v2url, edge.groupparent));
                            currentGroup.vertices.Add(edge.v1);
                            currentGroup.vertices.Add(edge.v2);
                        }
                    }
                    groups[graphid] = currentGroup; //O(1)
                }
                return groups;
            }
            public void CalculateAvgSimilarity(Dictionary<string, groupinfo> groups)
            {
                foreach (var g in graph) //#iter:g
                {
                    foreach (var edge in g.Value) // #iter: N  O(1)
                    {
                        groups[g.Key].avgsimilarity += (edge.similarityv1 + edge.similarityv2);
                        groups[g.Key].edgescount += 2;
                    }
                }
                foreach (var group in groups) //#iter:g
                {
                    group.Value.avgsimilarity = Math.Round((group.Value.avgsimilarity / group.Value.edgescount), 1); //O(1)
                }
            }
        }
        public static void Main()
        {
            TestClass t = new TestClass();
            bool cont = true;
            while (cont)
            {
                Console.WriteLine("Choose 1) Sample cases 2) Easy cases 3) Medium cases 4) Hard cases :\n");
                int Choice = int.Parse(Console.ReadLine());
                switch (Choice)
                {
                    case 1:
                        Console.WriteLine("Sample Cases :\n");
                        t.Test("C:\\Users\\Hager Essam\\OneDrive\\Desktop\\Collage\\Algo\\Project\\Test Cases\\Sample");
                        break;
                    case 2:
                        Console.WriteLine("Easy Cases :\n");
                        t.Test("C:\\Users\\Hager Essam\\OneDrive\\Desktop\\Collage\\Algo\\Project\\Test Cases\\Complete\\Easy");
                        break;
                    case 3:
                        Console.WriteLine("Medium Cases :\n");
                        t.Test("C:\\Users\\Hager Essam\\OneDrive\\Desktop\\Collage\\Algo\\Project\\Test Cases\\Complete\\Medium");
                        break;
                    case 4:
                        Console.WriteLine("Hard Cases :\n");
                        t.Test("C:\\Users\\Hager Essam\\OneDrive\\Desktop\\Collage\\Algo\\Project\\Test Cases\\Complete\\Hard");
                        break;
                    default:
                        cont = false;
                        break;
                }
            }
        }
    }
}