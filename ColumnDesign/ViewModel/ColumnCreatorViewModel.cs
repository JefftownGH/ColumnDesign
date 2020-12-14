using System.ComponentModel;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
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
        private string _quantity;
        private string _sQuantity;
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

        private ICommand _drawGatesCommand;
        private ICommand _drawScissorsCommand;

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
                if (WidthX.Equals("") || LengthY.Equals("")) return "";
                return $"{ConvertToNum(WidthX)}\" x {ConvertToNum(LengthY)}\" GATES COLUMN FORM";
            }
            set => _sheetName = value;
        }

        public string SSheetName
        {
            get
            {
                if (SWidthX.Equals("") || SLengthY.Equals("")) return "";
                return $"{ConvertToNum(SWidthX)}\" x {ConvertToNum(SLengthY)}\" SCISSOR CLAMP COLUMN FAB";
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
}