using System;
using System.Linq;
using Autodesk.Revit.DB;
using ColumnDesign.UI;

namespace ColumnDesign.Methods
{
    public static class Methods
    {
        private static Document _doc;

        public static void CreateGates(Document doc, ColumnCreatorView ui)
        {
            _doc = doc;
            const DrawingTypes type = DrawingTypes.Gates;
            using var tr = new Transaction(doc, $"Create new {type.ToString().ToLower()} sheet");
            tr.Start();
            var sheet = CreateSheet(doc, ui, DrawingTypes.Gates);
            var draftingView = CreateDraftingView(doc, sheet.Name);
            var draftingLocation = CalculateDraftingLocation(sheet);
            var clampPlanBack824Location = new XYZ(-1, 1, 0);
            var clampPlanOp824Location = new XYZ(0, 0, 0);
            doc.Create.NewFamilyInstance(clampPlanBack824Location,
                GetFamilySymbolByName(GlobalNames.WtClampPlanBack824), draftingView);
            doc.Create.NewFamilyInstance(clampPlanOp824Location, 
                GetFamilySymbolByName(GlobalNames.WtClampPlanOp824), draftingView);
           Viewport.Create(doc, sheet.Id, draftingView.Id, new XYZ(draftingLocation.U, draftingLocation.V, 0));
            tr.Commit();
        }

        private static FamilySymbol GetFamilySymbolByName(string name)
        {
            var symbol = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .First(s => s.Name.Equals(name));
            if (!symbol.IsActive) symbol.Activate();
            return symbol;
        }

        private static UV CalculateDraftingLocation(View sheet)
        {
            var maxO = sheet.Outline.Max;
            var minO = sheet.Outline.Min;
            return new UV((maxO.U - maxO.U) / 2, (maxO.V - minO.V) / 2);
        }

        public static void CreateScissors(Document doc, ColumnCreatorView ui)
        {
            const DrawingTypes type = DrawingTypes.Scissors;
            using var tr = new Transaction(doc, $"Create new {type.ToString().ToLower()} sheet");
            tr.Start();
            var sheet = CreateSheet(doc, ui, DrawingTypes.Scissors);
            var dv = CreateDraftingView(doc, sheet.Name);
            tr.Commit();
        }

        private static ViewDrafting CreateDraftingView(Document doc, string sheetName)
        {
            var vd = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(q => q.ViewFamily == ViewFamily.Drafting);

            if (vd == null) return null;
            var draftView = ViewDrafting.Create(doc, vd.Id);
            draftView.Name = $"{sheetName}_{DateTime.Now:MM-dd-yyyy-HH-mm-ss}";
            draftView.Scale = 25;
            return draftView;
        }

        private static ViewSheet CreateSheet(Document doc, ColumnCreatorView ui, DrawingTypes type)
        {
            var titleBlock = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .First(element => element.Name.Contains(GlobalNames.WtTitleBlockFamilyName));

            if (titleBlock == null) return null;
            var sheet = ViewSheet.Create(doc, titleBlock.Id);
            switch (type)
            {
                case DrawingTypes.Gates:
                    sheet.get_Parameter(BuiltInParameter.SHEET_NAME)?.Set(ui.SheetName.Text);
                    doc.ProjectInformation.Name = ui.ProjectTitle.Text;
                    doc.ProjectInformation.Address = ui.ProjectAddress.Text;
                    sheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE)?.Set(ui.Date.Text);
                    sheet.get_Parameter(BuiltInParameter.SHEET_CHECKED_BY)?.Set(ui.SheetIssuedFor.Text);
                    sheet.get_Parameter(BuiltInParameter.SHEET_DRAWN_BY)?.Set(ui.DrawnBy.Text);
                    sheet.get_Parameter(BuiltInParameter.SHEET_DESIGNED_BY)?.Set(ui.JobN.Text);
                    sheet.get_Parameter(BuiltInParameter.SHEET_NUMBER)?.Set(ui.SheetNumber.Text);
                    doc.ProjectInformation.Status = ui.Suffix.Text;
                    break;
                case DrawingTypes.Scissors:
                    sheet.get_Parameter(BuiltInParameter.SHEET_NAME)?.Set(ui.SSheetName.Text);
                    doc.ProjectInformation.Name = ui.ProjectTitle.Text;
                    doc.ProjectInformation.Address = ui.ProjectAddress.Text;
                    sheet.get_Parameter(BuiltInParameter.SHEET_ISSUE_DATE)?.Set(ui.SDate.Text);
                    sheet.get_Parameter(BuiltInParameter.SHEET_CHECKED_BY)?.Set(ui.SSheetIssuedFor.Text);
                    sheet.get_Parameter(BuiltInParameter.SHEET_DRAWN_BY)?.Set(ui.SDrawnBy.Text);
                    sheet.get_Parameter(BuiltInParameter.SHEET_DESIGNED_BY)?.Set(ui.SJobN.Text);
                    sheet.get_Parameter(BuiltInParameter.SHEET_NUMBER)?.Set(ui.SSheetNumber.Text);
                    doc.ProjectInformation.Status = ui.SSuffix.Text;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(type), type, null);
            }

            return sheet;
        }
    }
}