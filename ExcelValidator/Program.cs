using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace ExcelValidator
{
    internal class Program
    {
        private static int Main(string[] args)
        {
            String excel;
            String config;

            int cores = Environment.ProcessorCount;

            if (args.Length < 2)
            {
                Console.WriteLine(" You need to pass 2 parametrs:\n" +
                    " - path to Excel file\n" +
                    " - path to yaml config file with validation rules");

                return 1;
            }

            String location = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            String path = location + "/tmp";

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            if (File.Exists(path) == false)

                if (File.Exists(args[0]) == false)
                {
                    Console.WriteLine("Excel file do not exists!");
                    return 1;
                }

            if (File.Exists(args[1]) == false)
            {
                Console.WriteLine("Config yaml file do not exists!");
                return 1;
            }

            excel = args[0];
            config = args[1];

            Console.WriteLine("Parse excel file: {0} with config: {1}", excel, config);

            Dictionary<String, List<Validator.IValidator>> Validators = Yaml.ParseConfig(config);

            Console.WriteLine("Found: {0} validator(s)", Validators.Count);

            Excel.Excel sheet = new Excel.Excel(excel, path);
            String log = sheet.Validate(Validators);

            Console.WriteLine("log {0}", log);

            return 0;
        }
    }
}