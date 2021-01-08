using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ColumnDesign.Methods;
using ColumnDesign.Modules;
using ColumnDesign.ViewModel;

namespace ColumnDesign.UI
{
    public partial class ColumnCreatorView
    {
        private readonly Document _doc;
        private readonly UIDocument _uiDoc;
        private readonly ColumnCreatorViewModel _vm;

        public ColumnCreatorView(UIApplication uiApp, EventHandlerWithWpfArg eExternalMethodWpfArg)
        {
            _uiDoc = uiApp.ActiveUIDocument;
            _doc = _uiDoc.Document;
            _vm = new ColumnCreatorViewModel(eExternalMethodWpfArg, this);
            DataContext = _vm;
            InitializeComponent();
            InitializeFields();
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
    }
}