using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ColumnDesign.Methods;
using ColumnDesign.Modules;
using ColumnDesign.ViewModel;
using TextBox = System.Windows.Controls.TextBox;
using Visibility = System.Windows.Visibility;

namespace ColumnDesign.UI
{
    public partial class ColumnCreatorView
    {
        private readonly Document _doc;
        private readonly UIDocument _uiDoc;
        private readonly UIApplication _uiApp;
        private readonly ColumnCreatorViewModel _vm;

        public ColumnCreatorView(UIApplication uiApp, EventHandlerWithWpfArg eExternalMethodWpfArg)
        {
            _uiApp = uiApp;
            _uiDoc = uiApp.ActiveUIDocument;
            _doc = _uiDoc.Document;
            _vm = new ColumnCreatorViewModel(eExternalMethodWpfArg, this);
            DataContext = _vm;
            InitializeComponent();
            InitializeFields();
            EventManager.RegisterClassHandler(typeof(TextBox), KeyDownEvent, new KeyEventHandler(TextBox_KeyDown));
        }

        private static void TextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter & ((TextBox) sender).AcceptsReturn == false) MoveToNextUiElement(e);
        }

        private static void MoveToNextUiElement(RoutedEventArgs e)
        {
            const FocusNavigationDirection focusDirection = FocusNavigationDirection.Next;
            var request = new TraversalRequest(focusDirection);
            if (Keyboard.FocusedElement is not UIElement elementWithFocus) return;
            if (elementWithFocus.MoveFocus(request)) e.Handled = true;
        }

        private void InitializeFields()
        {
            var now = DateTime.Now.ToString("dd/M/yyyy", CultureInfo.InvariantCulture);
            _vm.Date = now;
            _vm.SDate = now;
            _vm.PlywoodType = "HDO";
            _vm.SlblAxis = "X";
        }

        private void Window_OnChecked(object sender, RoutedEventArgs e)
        {
            Picking.IsEnabled = false;
            Regular.IsChecked = true;
        }

        private void Window_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Picking.IsEnabled = true;
        }

        private void GridAxis_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _vm.SlblAxis = _vm.SlblAxis.Equals("X") ? "Y" : "X";
            UpdatePly_Function.UpdatePly(this, _vm);
        }

        private void Window_Click(object sender, RoutedEventArgs e)
        {
            UpdatePly_Function.UpdatePly(this, _vm);
        }

        private void BoxPlySeams_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            UpdatePly_Function.UpdatePly(this, _vm);
        }

        private void SBoxPlySeams_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            sUpdatePly_Function.sUpdatePly(this, _vm);
        }

        private void SaveSettings_OnMouseDown(object sender, RoutedEventArgs routedEventArgs)
        {
            using var file = new StreamWriter(GlobalNames.ConfigColumnLocation, false);
            file.WriteLine(SheetName.Text);
            file.WriteLine(ProjectTitle.Text);
            file.WriteLine(ProjectAddress.Text);
            file.WriteLine(Date.Text);
            file.WriteLine(SheetIssuedFor.Text);
            file.WriteLine(DrawnBy.Text);
            file.WriteLine(JobN.Text);
            file.WriteLine(SheetNumber.Text);
            file.WriteLine(Suffix.Text);
            file.WriteLine(Area.Text);
            file.Close();
        }

        private void SSaveSettings_OnMouseDown(object sender, RoutedEventArgs routedEventArgs)
        {
            using var file = new StreamWriter(GlobalNames.ConfigScissorsLocation, false);
            file.WriteLine(SSheetName.Text);
            file.WriteLine(SProjectTitle.Text);
            file.WriteLine(SProjectAddress.Text);
            file.WriteLine(SDate.Text);
            file.WriteLine(SSheetIssuedFor.Text);
            file.WriteLine(SDrawnBy.Text);
            file.WriteLine(SJobN.Text);
            file.WriteLine(SSheetNumber.Text);
            file.WriteLine(SSuffix.Text);
            file.WriteLine(SArea.Text);
            file.Close();
        }

        private void LoadSettings_OnMouseDown(object sender, RoutedEventArgs routedEventArgs)
        {
            if (!File.Exists(GlobalNames.ConfigColumnLocation)) return;
            using var file = new StreamReader(GlobalNames.ConfigColumnLocation);
            _vm.SheetName = file.ReadLine() ?? string.Empty;
            ProjectTitle.Text = file.ReadLine() ?? string.Empty;
            ProjectAddress.Text = file.ReadLine() ?? string.Empty;
            _vm.Date= file.ReadLine() ?? string.Empty;
            SheetIssuedFor.Text = file.ReadLine() ?? string.Empty;
            DrawnBy.Text = file.ReadLine() ?? string.Empty;
            JobN.Text = file.ReadLine() ?? string.Empty;
            SheetNumber.Text = file.ReadLine() ?? string.Empty;
            Suffix.Text = file.ReadLine() ?? string.Empty;
            Area.Text = file.ReadLine() ?? string.Empty;
            file.Close();
        }
        private void SLoadSettings_OnMouseDown(object sender, RoutedEventArgs routedEventArgs)
        {
            if (!File.Exists(GlobalNames.ConfigScissorsLocation)) return;
            using var file = new StreamReader(GlobalNames.ConfigScissorsLocation);
            _vm.SSheetName= file.ReadLine() ?? string.Empty;
            SProjectTitle.Text = file.ReadLine() ?? string.Empty;
            SProjectAddress.Text = file.ReadLine() ?? string.Empty;
            _vm. SDate = file.ReadLine() ?? string.Empty;
            SSheetIssuedFor.Text = file.ReadLine() ?? string.Empty;
            SDrawnBy.Text = file.ReadLine() ?? string.Empty;
            SJobN.Text = file.ReadLine() ?? string.Empty;
            SSheetNumber.Text = file.ReadLine() ?? string.Empty;
            SSuffix.Text = file.ReadLine() ?? string.Empty;
            SArea.Text = file.ReadLine() ?? string.Empty;
            file.Close();
        }

        private void WidthX_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is not TextBox box) return;
            WidthXHint.Visibility = box.Text.Length == 0 ? Visibility.Visible : Visibility.Hidden;
        }
        private void LengthY_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is not TextBox box) return;
            LengthYHint.Visibility = box.Text.Length == 0 ? Visibility.Visible : Visibility.Hidden;
        }
        private void HeightZ_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is not TextBox box) return;
            HeightZHint.Visibility = box.Text.Length == 0 ? Visibility.Visible : Visibility.Hidden;
        }
        private void Quantity_OnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is not TextBox box) return;
            QuantityHint.Visibility = box.Text.Length == 0 ? Visibility.Visible : Visibility.Hidden;
        }
    }
}