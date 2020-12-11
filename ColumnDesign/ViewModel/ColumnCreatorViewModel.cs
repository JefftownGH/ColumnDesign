using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using ColumnDesign.Annotations;

namespace ColumnDesign.ViewModel
{
    public sealed class ColumnCreatorViewModel : INotifyPropertyChanged
    {
        private string _widthX;
        private string _sheetName;
        private string _lengthY;
        private string _sSheetName;
        private string _sWidthX;
        private string _sLengthY;

        public string WidthX
        {
            get => _widthX;
            set
            {
                _widthX = value;
                OnPropertyChanged(nameof(WidthX));
                OnPropertyChanged(nameof(SheetName));
            }
        }

        public string SWidthX
        {
            get => _sWidthX;
            set
            {
                _sWidthX = value;
                OnPropertyChanged(nameof(SWidthX));
                OnPropertyChanged(nameof(SSheetName));
            }
        }

        public string LengthY
        {
            get => _lengthY;
            set
            {
                _lengthY = value;
                OnPropertyChanged(nameof(LengthY));
                OnPropertyChanged(nameof(SheetName));
            }
        }

        public string SLengthY
        {
            get => _sLengthY;
            set
            {
                _sLengthY = value;
                OnPropertyChanged(nameof(SLengthY));
                OnPropertyChanged(nameof(SSheetName));
            }
        }

        public string HeightZ { get; set; }
        public string SHeightZ { get; set; }
        public string Quantity { get; set; }
        public string SQuantity { get; set; }

        public string SheetName
        {
            get
            {
                if (WidthX.Equals("") || LengthY.Equals("")) return "";
                return $"{WidthX}\" x {LengthY}\" GATES COLUMN FORM";
            }
            set => _sheetName = value;
        }

        public string SSheetName
        {
            get
            {
                if (SWidthX.Equals("") || SLengthY.Equals("")) return "";
                return $"{WidthX}\" x {LengthY}\" SCISSOR CLAMP COLUMN FAB";
            }
            set => _sSheetName = value;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
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