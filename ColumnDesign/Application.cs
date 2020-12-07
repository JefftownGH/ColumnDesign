using System;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace ColumnDesign
{
    public class Application : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
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
    }
}