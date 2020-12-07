using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ColumnDesign
{
    public class ColumnCreator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            return Result.Succeeded;
        }
    }
}