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
using static ColumnDesign.Modules.ConvertNumberToFeetInches;
using static ColumnDesign.Modules.sUpdatePly_Function;
using static ColumnDesign.Modules.UpdatePly_Function;

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
        private string _slblAxis;
        private string _winDim2;
        private string _winDim1;
        private string _boxPlySeams;
        private string _sBoxPlySeams;
        private string _sPlywoodType;

        public string WidthX
        {
            get => _widthX;
            set
            {
                _widthX = value;
                _boxPlySeams = "";
                OnPropertyChanged(nameof(BoxPlySeams));
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
                _sBoxPlySeams = "";
                sCheckHeight(ConvertToNum(SWidthX), ConvertToNum(SLengthY), ConvertToNum(SHeightZ), _view);
                OnPropertyChanged(nameof(SBoxPlySeams));
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
                _boxPlySeams = "";
                OnPropertyChanged(nameof(BoxPlySeams));
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
                _sBoxPlySeams = "";
                sCheckHeight(ConvertToNum(SWidthX), ConvertToNum(SLengthY), ConvertToNum(SHeightZ), _view);
                OnPropertyChanged(nameof(SBoxPlySeams));
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
                _boxPlySeams = "";
                OnPropertyChanged(nameof(BoxPlySeams));
                OnPropertyChanged(nameof(HeightZ));
                OnPropertyChanged(nameof(SheetName));
            }
        }

        public string SHeightZ
        {
            get => _sHeightZ;
            set
            {
                _sHeightZ = value;
                _sBoxPlySeams = "";
                sCheckHeight(ConvertToNum(SWidthX), ConvertToNum(SLengthY), ConvertToNum(SHeightZ), _view);
                OnPropertyChanged(nameof(SBoxPlySeams));
                OnPropertyChanged(nameof(SHeightZ));
                OnPropertyChanged(nameof(SSheetName));
            }
        }


        public string Quantity
        {
            get => _quantity;
            set
            {
                _quantity = value;
                _boxPlySeams = "";
                OnPropertyChanged(nameof(BoxPlySeams));
                OnPropertyChanged(nameof(Quantity));
            }
        }

        public string SQuantity
        {
            get => _sQuantity;
            set
            {
                _sQuantity = value;
                _sBoxPlySeams = "";
                sCheckHeight(ConvertToNum(SWidthX), ConvertToNum(SLengthY), ConvertToNum(SHeightZ), _view);
                OnPropertyChanged(nameof(SBoxPlySeams));
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

        public string SPlywoodType
        {
            get => _sPlywoodType;
            set
            {
                _sPlywoodType = value;
                OnPropertyChanged(nameof(SPlywoodType));
            }
        }

        public bool WindowX
        {
            get => _windowX;
            set
            {
                _windowY = false;
                _windowX = value;
                _slblAxis = "X";
                OnPropertyChanged(nameof(WindowX));
                OnPropertyChanged(nameof(WindowY));
                OnPropertyChanged(nameof(SlblAxis));
                UpdatePly(_view, this);
            }
        }

        public bool WindowY
        {
            get => _windowY;
            set
            {
                _windowX = false;
                _windowY = value;
                _slblAxis = "Y";
                OnPropertyChanged(nameof(WindowY));
                OnPropertyChanged(nameof(WindowX));
                OnPropertyChanged(nameof(SlblAxis));
                UpdatePly(_view, this);
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

        public string SlblAxis
        {
            get => _slblAxis;
            set
            {
                _slblAxis = value;
                OnPropertyChanged(nameof(SlblAxis));
            }
        }

        public string WinDim1
        {
            get => _winDim1;
            set
            {
                _winDim1 = value;
                OnPropertyChanged(nameof(WinDim1));
            }
        }

        public string WinDim2
        {
            get => _winDim2;
            set
            {
                _winDim2 = value;
                OnPropertyChanged(nameof(WinDim2));
            }
        }

        public string BoxPlySeams
        {
            get => _boxPlySeams;
            set
            {
                _boxPlySeams = value;
                OnPropertyChanged(nameof(BoxPlySeams));
            }
        }

        public string SBoxPlySeams
        {
            get => _sBoxPlySeams;
            set
            {
                _sBoxPlySeams = value;
                OnPropertyChanged(nameof(SBoxPlySeams));
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
            _eExternalMethodWpfArg.Raise(_view, this, DrawingTypes.Gates);
        }

        private void DrawScissors()
        {
            _eExternalMethodWpfArg.Raise(_view, this, DrawingTypes.Scissors);
        }

        private static bool DrawGatesCanExecute => true;

        private static bool DrawScissorsCanExecute => true;

        public string SheetName
        {
            get
            {
                if (WidthX == null) return "";
                if (LengthY == null) return "";
                if (HeightZ == null) return "";
                if (WidthX.Equals("")) return "";
                if (LengthY.Equals("")) return "";
                if (HeightZ.Equals("")) return "";
                return
                    $"{ConvertToNum(WidthX)}\" x {ConvertToNum(LengthY)}\" x {ConvertFtIn(ConvertToNum(HeightZ))} GATES COLUMN FORM";
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
                if (SHeightZ == null) return "";
                if (SWidthX.Equals("")) return "";
                if (SLengthY.Equals("")) return "";
                if (SHeightZ.Equals("")) return "";
                return
                    $"{ConvertToNum(SWidthX)}\" x {ConvertToNum(SLengthY)}\" x {ConvertFtIn(ConvertToNum(SHeightZ))} SCISSOR CLAMP COLUMN FAB";
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
                return ValidationResult.ValidResult;
            }

            if (((string) value).Substring(0,1).Equals("-") )
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