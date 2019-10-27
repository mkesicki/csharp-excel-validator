using System;
using System.Collections.Generic;
using System.IO;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace ExcelValidator
{
    internal class Yaml
    {
        private static Dictionary<String, List<Validator.IValidator>> Validators;

        public static Dictionary<String, List<Validator.IValidator>> ParseConfig(String path)
        {
            Validators = new Dictionary<String, List<Validator.IValidator>>();

            using (var reader = File.OpenText(path))
            {
                var deserializer = new DeserializerBuilder()
              .WithNamingConvention(new CamelCaseNamingConvention())
              .Build();

                var validators = deserializer.Deserialize<YamlParser>(reader);

                List<Validator.IValidator> defaults = new List<Validator.IValidator>();

                if (validators.defaults != null)
                {
                    foreach (YamlValidator item in validators.defaults)
                    {
                        defaults.Add(Validator.AbstractValidator.CreateValidator(item.name, item.GetOptions(), item.message));
                    }

                    Validators.Add("defaults", defaults);
                }

                List<Validator.IValidator> columns;

                if (validators.columns != null)
                {
                    foreach (KeyValuePair<String, List<YamlValidator>> kvp in validators.columns)
                    {
                        columns = new List<Validator.IValidator>();
                        foreach (YamlValidator validator in kvp.Value)
                        {
                            columns.Add(Validator.AbstractValidator.CreateValidator(validator.name, validator.GetOptions(), validator.message));
                        }
                        Validators.Add(kvp.Key, columns);
                    }
                }
            }

            return Validators;
        }
    }
}