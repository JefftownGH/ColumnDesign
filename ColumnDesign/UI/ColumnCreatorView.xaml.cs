using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ColumnDesign.Methods;
using ColumnDesign.ViewModel;

namespace ColumnDesign.UI
{
    public partial class ColumnCreatorView
    {
        private readonly Document _doc;
        private readonly UIDocument _uiDoc;

        public ColumnCreatorView(UIApplication uiApp, EventHandlerWithWpfArg eExternalMethodWpfArg)
        {
            _uiDoc = uiApp.ActiveUIDocument;
            _doc = _uiDoc.Document;
            DataContext = new ColumnCreatorViewModel(eExternalMethodWpfArg, this);
            InitializeComponent();
            InitializeFields();
        }

        private void InitializeFields()
        {
            var now = DateTime.Now.ToString("dd/M/yyyy", CultureInfo.InvariantCulture);
            Date.Text = now;
            SDate.Text = now;
        }
        private void Window_OnChecked(object sender, RoutedEventArgs e)
        {
            if (WindowX.IsChecked == true) WindowY.IsChecked = false;
            if (WindowY.IsChecked == true) WindowX.IsChecked = false;
            Picking.IsEnabled = false;
            Regular.IsChecked = true;
        }
        
        private void Window_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Picking.IsEnabled = true;
        }
    }
}