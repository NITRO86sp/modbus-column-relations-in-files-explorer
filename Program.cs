using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using LinqToExcel;

namespace ConsoleApplication1
{
    class Program
    {
        static void Main(string[] args)
        {
            var hDic = new SortedDictionary<string, List<string>>();
            var dirName = @"G:\Dropbox\`ITU SDT\4th Semester (thesis)\THESIS from PC\MBusCom\";
            var fileNames = Directory.EnumerateFiles(dirName,"*", SearchOption.TopDirectoryOnly).ToList();

            foreach (var f in fileNames)
            {
                var xlFactory = new ExcelQueryFactory(f);
                xlFactory.ReadOnly = true;
                var wsNames = xlFactory.GetWorksheetNames();
                var cNames = xlFactory.GetColumnNames(wsNames.First());

                foreach (var c in cNames) {
                    List<string> currentValue= null;
                    if (hDic.ContainsKey(c)) {
                        try
                        {
                            hDic.TryGetValue(c, out currentValue);
                            currentValue.Add(Path.GetFileName(f));
                            hDic[c]= currentValue;

                        }
                        
                        catch (Exception e)
                        {Console.WriteLine(e);}

                    } else {
                        hDic.Add(c, new List<string> { Path.GetFileName(f) });
                    }
                    
                }
            }
            var shDic=hDic.OrderByDescending(i => i.Value.Count());
            foreach (var c in shDic)
            {
                Console.WriteLine(c.Key);
                foreach (var f in c.Value) {
                Console.Write(f + " ");
                }
                Console.WriteLine("");
            }
            Console.ReadKey();
        }
    }
}
