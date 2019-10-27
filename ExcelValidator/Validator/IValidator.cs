using System;

namespace ExcelValidator.Validator
{
    public interface IValidator
    {
        String Message { get; set; }

        String Name { get; set; }

        Boolean IsValid(String Value);
    }
}