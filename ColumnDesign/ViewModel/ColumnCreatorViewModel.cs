using System;
using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using ColumnDesign.Annotations;
using ColumnDesign.Methods;
using ColumnDesign.UI;
using static ColumnDesign.Modules.ConvertFeetInchesToNumber;

namespace ColumnDesign.ViewModel
{
    public sealed class ColumnCreatorViewModel : INotifyPropertyChanged
    {
        private readonly EventHandlerWithWpfArg _eExternalMethodWpfArg;
        private readonly ColumnCreatorView _view;
        private string _widthX;
        private string _sheetName;
        private string _lengthY;
        private string _sSheetName;
        private string _sWidthX;
        private string _sLengthY;
        private string _heightZ;
        private string _sHeightZ;
        private string _quantity;
        private string _sQuantity;
        private bool _windowX;
        private bool _windowY;
        private string _plywoodType;
        private string _date;
        private string _sDate;
        private ICommand _drawGatesCommand;
        private ICommand _drawScissorsCommand;

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

        public string HeightZ
        {
            get => _heightZ;
            set
            {
                _heightZ = value;
                OnPropertyChanged(nameof(HeightZ));
            }
        }

        public string SHeightZ
        {
            get => _sHeightZ;
            set
            {
                _sHeightZ = value;
                OnPropertyChanged(nameof(SHeightZ));
            }
        }


        public string Quantity
        {
            get => _quantity;
            set
            {
                _quantity = value;
                OnPropertyChanged(nameof(Quantity));
            }
        }

        public string SQuantity
        {
            get => _sQuantity;
            set
            {
                _sQuantity = value;
                OnPropertyChanged(nameof(SQuantity));
            }
        }

        public string PlywoodType
        {
            get => _plywoodType;
            set
            {
                _plywoodType = value;
                OnPropertyChanged(nameof(PlywoodType));
            }
        }

        public bool WindowX
        {
            get => _windowX;
            set
            {
                _windowY = false;
                _windowX = value;
                OnPropertyChanged(nameof(WindowX));
                OnPropertyChanged(nameof(WindowY));
            }
        }

        public bool WindowY
        {
            get => _windowY;
            set
            {
                _windowX = false;
                _windowY = value;
                OnPropertyChanged(nameof(WindowY));
                OnPropertyChanged(nameof(WindowX));
            }
        }

        public string Date
        {
            get => _date;
            set
            {
                _date = value;
                OnPropertyChanged(nameof(Date));
            }
        }

        public string SDate
        {
            get => _sDate;
            set
            {
                _sDate = value;
                OnPropertyChanged(nameof(SDate));
            }
        }

        public ColumnCreatorViewModel(EventHandlerWithWpfArg eExternalMethodWpfArg, ColumnCreatorView view)
        {
            _eExternalMethodWpfArg = eExternalMethodWpfArg;
            _view = view;
        }

        public ICommand DrawGatesCommand
        {
            get { return _drawGatesCommand ??= new DrawButtonHandler(DrawGates, () => DrawGatesCanExecute); }
        }

        public ICommand DrawScissorsCommand
        {
            get { return _drawScissorsCommand ??= new DrawButtonHandler(DrawScissors, () => DrawScissorsCanExecute); }
        }

        private void DrawGates()
        {
            _eExternalMethodWpfArg.Raise(_view, DrawingTypes.Gates);
        }

        private void DrawScissors()
        {
            _eExternalMethodWpfArg.Raise(_view, DrawingTypes.Scissors);
        }

        private static bool DrawGatesCanExecute => true;

        private static bool DrawScissorsCanExecute => true;

        public string SheetName
        {
            get
            {
                if (WidthX == null) return "";
                if (LengthY == null) return "";
                if (WidthX.Equals("")) return "";
                if (LengthY.Equals("")) return "";
                return $"{ConvertToNum(WidthX)}\" x {ConvertToNum(LengthY)}\" GATES COLUMN FORM";
            }
            set
            {
                _sheetName = value;
                OnPropertyChanged(nameof(SheetName));
            }
        }

        public string SSheetName
        {
            get
            {
                if (SWidthX == null) return "";
                if (SLengthY == null) return "";
                if (SWidthX.Equals("")) return "";
                if (SLengthY.Equals("")) return "";
                return $"{ConvertToNum(SWidthX)}\" x {ConvertToNum(SLengthY)}\" SCISSOR CLAMP COLUMN FAB";
            }
            set
            {
                _sSheetName = value;
                OnPropertyChanged(nameof(SSheetName));
            }
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
            if (value == null || ((string) value).Length == 0)
            {
                return new ValidationResult(false, $"Please enter value in the range: {Min}-{Max}.");
            }

            var val = ConvertToNum((string) value);
            if (val == 0d)
            {
                return new ValidationResult(false, "Illegal characters");
            }

            if (val < Min || val > Max)
            {
                return new ValidationResult(false, $"Please enter value in the range: {Min}-{Max}.");
            }

            return ValidationResult.ValidResult;
        }
    }

    [ValueConversion(typeof(string), typeof(Visibility))]
    public class CheckPlywoodVisibilityConverter : System.Windows.Markup.MarkupExtension, IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value switch
            {
                "HDO" => Visibility.Hidden,
                _ => Visibility.Visible
            };
        }

        public object ConvertBack(object value, Type targetType, object parameter,
            System.Globalization.CultureInfo culture)
        {
            return null;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }

    [ValueConversion(typeof(bool), typeof(Visibility))]
    public class WinDimSeamsVisibilityConverter : System.Windows.Markup.MarkupExtension, IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            foreach (var value in values)
            {
                if (value is not bool isVisible) break;
                if (isVisible) return Visibility.Visible;
            }

            return Visibility.Collapsed;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            return null;
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}