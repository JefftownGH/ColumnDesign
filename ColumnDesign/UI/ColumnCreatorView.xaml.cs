using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ColumnDesign.Methods;

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
            InitializeComponent();
            _mExternalMethodWpfArg = eExternalMethodWpfArg;
        }

        private void ButtonDrawGates_OnClick(object sender, RoutedEventArgs e)
        {
            _mExternalMethodWpfArg.Raise(this);
        }
    }
}