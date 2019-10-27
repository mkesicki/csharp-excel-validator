using System;
using System.Collections.Generic;

namespace ExcelValidator.Validator
{
    internal class TypeValidator : AbstractValidator, IValidator
    {
        private String Type { get; set; }

        public override Boolean IsValid(String value)
        {
            try
            {
                switch (Type.ToLower())
                {
                    case "boolean":
                        bool.Parse(value);
                        return true;

                    case "integer":
                        int.Parse(value);
                        return true;

                    case "double":
                        double.Parse(value);
                        return true;

                    case "date":
                        DateTime date = new DateTime();
                        bool valid = DateTime.TryParse(value, out date);
                        if (valid) { return true; }
                        break;
                }
            }
            catch (Exception)
            {
                return false;
            }

            return false;
        }

        public TypeValidator()
        {
            Name = "Type";
        }

        public TypeValidator(String name, String message, Dictionary<String, String> options) : base(name, message, options)
        {
            if (options.ContainsKey("type"))
            {
                Type = options["type"].ToLower();
            }
            else
            {
                throw new Exception("You need to set type for TypeValidator.");
            }
        }
    }
}