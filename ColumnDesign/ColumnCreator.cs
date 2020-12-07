using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace ColumnDesign
{
    [Transaction(TransactionMode.Manual)]
    public class ColumnCreator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiDoc = commandData.Application.ActiveUIDocument;
            var doc = uiDoc.Document;
            CreateSheet(doc);
            return Result.Succeeded;
        }

        private static void CreateSheet(Document doc)
        {
            using var tr = new Transaction(doc, "Create new sheet");
            tr.Start();
            var collector = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks);
            Element titleBlock = null;
            foreach(var element in collector)
            {
                if(element.Name.Contains(GlobalNames.TitleBlockFamilyName))
                {
                    titleBlock = element;
                }
            }

            if (titleBlock != null)
            {
                var viewSheet = ViewSheet.Create(doc, titleBlock.Id);
            }
            tr.Commit();
        }
    }
}