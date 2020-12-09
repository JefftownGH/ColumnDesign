using Autodesk.Revit.DB;
using ColumnDesign.UI;

namespace ColumnDesign.Methods
{
    public static class Methods
    {
        public static void CreateSheet(Document doc, ColumnCreatorView ui)
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
                var sheet = ViewSheet.Create(doc, titleBlock.Id);
                sheet.get_Parameter(BuiltInParameter.SHEET_NAME)?.Set(ui.SheetName.Text);
                doc.ProjectInformation.Name = ui.ProjectTitle.Text;
                doc.ProjectInformation.Address = ui.ProjectAddress.Text;
                sheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE )?.Set(ui.Date.Text);
                sheet.get_Parameter(BuiltInParameter.SHEET_CHECKED_BY )?.Set(ui.SheetIssuedFor.Text);
                sheet.get_Parameter(BuiltInParameter.SHEET_DRAWN_BY )?.Set(ui.DrawnBy.Text);
                sheet.get_Parameter(BuiltInParameter.SHEET_DESIGNED_BY )?.Set(ui.JobN.Text);
                sheet.get_Parameter(BuiltInParameter.SHEET_NUMBER )?.Set(ui.SheetNumber.Text);
                doc.ProjectInformation.Status = ui.Suffix.Text;
            }
            tr.Commit();
        }
    }
}