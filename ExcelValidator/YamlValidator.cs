using System;
using System.Collections.Generic;
using System.Reflection;

namespace ExcelValidator
{
    internal class YamlParser
    {
        public List<YamlValidator> defaults { get; set; }
        public Dictionary<String, List<YamlValidator>> columns { get; set; }
    }
}

internal class YamlOptions
{
    public String type { get; set; }
    public Boolean trim { get; set; }
    public Int16 min { get; set; }
    public Int16 max { get; set; }
    public String minMessage { get; set; }
    public String maxMessage { get; set; }
}

internal class YamlValidator
{
    public String name { get; set; }
    public String message { get; set; }
    public List<YamlOptions> options { get; set; }

    public Dictionary<String, String> GetOptions()
    {
        Dictionary<String, String> results = new Dictionary<String, String>();

        if (options != null && options.Count > 0)
        {
            foreach (YamlOptions option in options)
            {
                PropertyInfo[] properties = option.GetType().GetProperties();

                foreach (var property in properties)
                {
                    if (property.GetValue(option) != null)
                    {
                        results.Add(property.Name, Convert.ToString(property.GetValue(option)));
                    }
                }
            }
        }

        return results;
    }
}