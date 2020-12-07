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
            if (panel.AddItem(new PushButtonData("AutoNumerate", "Column\ncreator",
                Assembly.GetExecutingAssembly().Location,
                typeof(ColumnCreator).FullName)) is PushButton buttonAutoNumerate)
            {
                buttonAutoNumerate.ToolTip = "Automatic creation of worksheets";
                buttonAutoNumerate.Image = new BitmapImage(new Uri(
                    "pack://application:,,,/ColumnDesign;component/Resources/ColumnCreator16.png"));
                buttonAutoNumerate.LargeImage = new BitmapImage(new Uri(
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