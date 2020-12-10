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
            _mExternalMethodWpfArg = eExternalMethodWpfArg;
            InitializeFields();
        }

        private void InitializeFields()
        {
            Date.Text = DateTime.Today.ToString("dd/MM/yyyy");
        }

        private void ButtonDrawGates_OnClick(object sender, RoutedEventArgs e)
        {
            _mExternalMethodWpfArg.Raise(this);
        }
        private void WindowX_OnChecked(object sender, RoutedEventArgs e)
        {
            if (WindowY.IsChecked == true)
            {
                WindowY.IsChecked = false;
            }
        }
        private void WindowY_OnChecked(object sender, RoutedEventArgs e)
        {
            if (WindowX.IsChecked == true)
            {
                WindowX.IsChecked = false;
            }
        }
    }
}