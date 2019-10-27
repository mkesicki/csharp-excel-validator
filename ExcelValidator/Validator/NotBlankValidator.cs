using System;
using System.Collections.Generic;

namespace ExcelValidator.Validator
{
    internal class NotBlankValidator : AbstractValidator, IValidator
    {
        public override Boolean IsValid(String value)
        {
            return String.IsNullOrEmpty(value) != true;
        }

        public NotBlankValidator()
        {
            Name = "NotBlank";
        }

        public NotBlankValidator(String name, String message, Dictionary<String, String> options) : base(name, message, options)
        {
        }
    }
}