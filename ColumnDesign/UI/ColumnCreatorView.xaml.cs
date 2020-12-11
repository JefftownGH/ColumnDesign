using System;
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
        private readonly EventHandlerWithWpfArg _mExternalMethodWpfArg;

        public ColumnCreatorView(UIApplication uiApp, EventHandlerWithWpfArg eExternalMethodWpfArg)
        {
            _uiDoc = uiApp.ActiveUIDocument;
            _doc = _uiDoc.Document;
            DataContext = new ColumnCreatorViewModel();
            InitializeComponent();
            InitializeFields();
            _mExternalMethodWpfArg = eExternalMethodWpfArg;
        }

        private void InitializeFields()
        {
            var now = DateTime.Now.ToString("dd/M/yyyy");
            Date.Text = now;
            SDate.Text = now;
        }

        private void ButtonDrawGates_OnClick(object sender, RoutedEventArgs e)
        {
            _mExternalMethodWpfArg.Raise(this);
        }
        private void Window_OnChecked(object sender, RoutedEventArgs e)
        {
            if (WindowY.IsChecked == true) WindowY.IsChecked = false;
            if (WindowX.IsChecked == true) WindowX.IsChecked = false;
            Picking.IsEnabled = false;
            Regular.IsChecked = true;
        }
        
        private void Window_OnUnchecked(object sender, RoutedEventArgs e)
        {
            Picking.IsEnabled = true;
        }
    }
}