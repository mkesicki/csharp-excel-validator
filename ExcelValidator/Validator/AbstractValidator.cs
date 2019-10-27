using System;
using System.Collections.Generic;

namespace ExcelValidator.Validator
{
    internal abstract class AbstractValidator : IValidator
    {
        public String Message { get; set; }
        public String Name { get; set; }

        public Dictionary<String, String> Options { get; set; }

        public abstract Boolean IsValid(string Value);

        public AbstractValidator(String name, String message, Dictionary<String, String> options)
        {
            Options = options;
            Name = name;
            Message = message;
        }

        public AbstractValidator(String name, String message = "")
        {
            Name = name;
            if (!message.Equals(""))
            {
                Message = message;
            }
        }

        public AbstractValidator()
        {
        }

        public static IValidator CreateValidator(String name, Dictionary<String, String> options, String message = "")
        {
            String objectType = "ExcelValidator.Validator." + name + "Validator";
            Type type = Type.GetType(objectType);

            return (IValidator)Activator.CreateInstance(type, name, message, options);
        }
    }
}