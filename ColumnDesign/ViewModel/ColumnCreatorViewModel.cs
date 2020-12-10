using System.Globalization;
using System.Windows.Controls;

namespace ColumnDesign.ViewModel
{
    public class ColumnCreatorViewModel
    {
        public int WidthX { get; set; }
        public int LengthY { get; set; }
        public int HeightZ { get; set; }
    }

    public class ValueRangeRule : ValidationRule
    {
        public int Min { get; set; }
        public int Max { get; set; }

        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            var val = 0;
            if (value != null && ((string) value).Length > 0)
            {
                if (!int.TryParse((string) value, out val))
                {
                    return new ValidationResult(false, "Illegal characters");
                }
            }
            if (val < Min || val > Max)
            {
                return new ValidationResult(false, $"Please enter value in the range: {Min}-{Max}.");
            }

            return ValidationResult.ValidResult;
        }
    }
}