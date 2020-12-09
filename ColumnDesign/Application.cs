using System;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;
using ColumnDesign.Methods;
using ColumnDesign.UI;

namespace ColumnDesign
{
    public class Application : IExternalApplication
    {
        public static Application ThisApp;

        private ColumnCreatorView _form;
        public Result OnStartup(UIControlledApplication application)
        {
            _form = null;
            ThisApp = this;
            var panel = application.CreateRibbonPanel("Column design");
            if (panel.AddItem(new PushButtonData("ColumnCreator", "Column\ncreator",
                Assembly.GetExecutingAssembly().Location,
                typeof(ColumnCreator).FullName)) is PushButton buttonColumnCreator)
            {
                buttonColumnCreator.ToolTip = "Automatic creation of worksheets";
                buttonColumnCreator.Image = new BitmapImage(new Uri(
                    "pack://application:,,,/ColumnDesign;component/Resources/ColumnCreator16.png"));
                buttonColumnCreator.LargeImage = new BitmapImage(new Uri(
                    "pack://application:,,,/ColumnDesign;component/Resources/ColumnCreator32.png"));
            }
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
        public void ShowForm(UIApplication uiApp)
        {
            if (_form != null && _form == null) return;
            var evWpf = new EventHandlerWithWpfArg();
            _form = new ColumnCreatorView(uiApp, evWpf);
            _form.Show();
        }
    }
}