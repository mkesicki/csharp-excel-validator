using OfficeOpenXml;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace ExcelValidator.Excel
{
    internal class Excel
    {
        private ConcurrentDictionary<String, ExcelWorksheet> Sheets = new ConcurrentDictionary<String, ExcelWorksheet>();
        private List<String> LogFiles = new List<String>();
        private Dictionary<String, List<Validator.IValidator>> Validators;
        private List<String> Files = new List<String>();
        private List<Thread> Threads = new List<Thread>();

        private List<StreamWriter> Logs = new List<StreamWriter>();

        private String UUID;

        private FileInfo resultsDir;

        private int RowsCount;
        private int BatchSize;
        private int Cores;

        public Excel(String excel, String path)
        {
            Console.WriteLine("Open Excel file.");

            FileInfo excelFile = new FileInfo(excel);
            resultsDir = new FileInfo(path);

            Configure(excelFile, resultsDir);
        }

        private void Configure(FileInfo excelFile, FileInfo dir)
        {
            using (ExcelPackage package = new ExcelPackage(excelFile))
            {
                Cores = Environment.ProcessorCount;
                Console.WriteLine("The number of processors on this computer is {0}.", Cores);

                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                RowsCount = worksheet.Dimension.End.Row;
                BatchSize = RowsCount / Cores;

                Console.WriteLine("Found {0} rows in Excel file.", RowsCount.ToString());
            }

            UUID = Guid.NewGuid().ToString();
            for (int i = 0; i < Cores; i++)
            {
                FileInfo file = excelFile.CopyTo(dir.DirectoryName + "\\tmp\\" + Path.GetRandomFileName());

                Files.Add(file.FullName);
                FileInfo fi = new FileInfo(dir.ToString() + "\\" + UUID + "_" + i.ToString() + ".log");
                StreamWriter log = new StreamWriter(fi.FullName);
                LogFiles.Add(fi.FullName);
                Logs.Add(log);
            }
        }

        public void Parse(Object core)
        {
            int cpu = (int)core;
            int firstRow = cpu * BatchSize + 1;
            int lastRow = ((cpu + 1) == Cores) ? RowsCount + 1 : firstRow + BatchSize;
            Console.WriteLine("Thread #{0} started: firstRow {1}; lastRow {2}", core, firstRow, lastRow);

            using (ExcelPackage package = new ExcelPackage(new FileInfo(Files[cpu])))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                StreamWriter log = Logs[cpu];

                for (int rowIndex = firstRow; rowIndex < lastRow; rowIndex++)
                {
                    for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                    {
                        String column = ColumnIndexToColumnLetter(j);
                        String value = "";
                        if (worksheet.Cells[rowIndex, j].Value != null)
                        {
                            value = worksheet.Cells[rowIndex, j].Value.ToString();
                        }

                        if (Validators.ContainsKey("defaults"))
                        {
                            foreach (Validator.IValidator validator in Validators["defaults"])
                            {
                                if (false == validator.IsValid(value))
                                {
                                    Console.WriteLine("Found error in cell[{0}{1}]", column, rowIndex);
                                    log.WriteLine("Error in cell [{0}{1}], value \"{2}\" is not valid. Validator message: \"{3}\"", column, rowIndex, value, validator.Message);
                                }
                            }
                        }

                        if (Validators.ContainsKey(column))
                        {
                            foreach (Validator.IValidator validator in Validators[column])
                            {
                                if (false == validator.IsValid(value))
                                {
                                    Console.WriteLine("Found error in cell[{0}{1}]", column, rowIndex);
                                    log.WriteLine("Error in cell [{0}{1}], value \"{2}\" is not valid. Validator message: \"{3}\"", column, rowIndex, value, validator.Message);
                                }
                            }
                        }
                    }
                }
            }
        }

        public String Validate(Dictionary<String, List<Validator.IValidator>> config)
        {
            Validators = config;
            for (int i = 0; i < Cores; i++)
            {
                ParameterizedThreadStart start = new ParameterizedThreadStart(Parse);
                Thread t = new Thread(start);
                t.Start(i);
                Threads.Add(t);
            }

            bool loop = true;

            while (loop)
            {
                int count = 0;
                for (int i = 0; i < Cores; i++)
                {
                    if (Threads[i].IsAlive == false)
                    {
                        count++;
                        Logs[i].Close();
                    }
                }
                if (count == Cores) { loop = false; }
            }

            Console.WriteLine("Store log file");
            String logAll = "";
            foreach (String fileName in LogFiles)
            {
                logAll = logAll + System.IO.File.ReadAllText(fileName);
            }

            String logFileName = resultsDir + "\\" + UUID + "_all.log";
            System.IO.File.WriteAllText(logFileName, logAll);

            return logFileName;
        }

        public String ColumnIndexToColumnLetter(int colIndex)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (div - mod) / 26;
            }
            return colLetter;
        }
    }
}