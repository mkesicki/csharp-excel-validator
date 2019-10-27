using System;
using System.Collections.Generic;

namespace ExcelValidator.Validator
{
    internal class LengthValidator : AbstractValidator, IValidator
    {
        private String MinMessage { get; set; }
        private String MaxMessage { get; set; }

        private int Min { get; set; }
        private int Max { get; set; }

        public override Boolean IsValid(String value)
        {
            if (Min > 0 && value.Length < Min)
            {
                Message = MinMessage;
                return false;
            }

            if (Max > 0 && value.Length > Max)
            {
                Message = MaxMessage;
                return false;
            }

            return true;
        }

        public LengthValidator()
        {
            Name = "Length";
        }

        public LengthValidator(String name, String message, Dictionary<String, String> options) : base(name, message, options)
        {
            if (options.ContainsKey("minMessage"))
            {
                MinMessage = options["minMessage"];
            }

            if (options.ContainsKey("maxMessage"))
            {
                MaxMessage = options["maxMessage"];
            }

            if (options.ContainsKey("min"))
            {
                Min = Int32.Parse(options["min"]);
            }

            if (options.ContainsKey("max"))
            {
                Max = Int32.Parse(options["max"]);
            }
        }
    }
}