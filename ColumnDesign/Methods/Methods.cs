using System;
using System.Linq;
using Autodesk.Revit.DB;
using ColumnDesign.UI;
using static ColumnDesign.Modules.ConvertFeetInchesToNumber;
using static ColumnDesign.Modules.ImportMatrix_Function;

namespace ColumnDesign.Methods
{
    public static class Methods
    {
        private static Document _doc;
        private static ColumnCreatorView _ui;

        public static void CreateGates(Document doc, ColumnCreatorView ui)
        {
            _doc = doc;
            _ui = ui;
            const DrawingTypes type = DrawingTypes.Gates;
            using var tr = new Transaction(doc, $"Create new {type.ToString().ToLower()} sheet");
            tr.Start();
            var sheet = CreateSheet(doc, ui, DrawingTypes.Gates);
            var draftingView = CreateDraftingView(doc, sheet.Name);
            var draftingLocation = CalculateDraftingLocation(sheet);
            CreateClamps(draftingView);
            Viewport.Create(doc, sheet.Id, draftingView.Id, new XYZ(draftingLocation.U, draftingLocation.V, 0));
            tr.Commit();
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

        private static void CreateClamps(View draftingView)
        {
            double clamp_spacing;
            double swing_ang;
            double win_clamp_top_max = 5;
            double win_clamp_bot_max = 4;
            var widthX = ConvertToNum(_ui.WidthX.Text);
            var lengthY = ConvertToNum(_ui.LengthY.Text);
            var heightZ = ConvertToNum(_ui.HeightZ.Text);
            var bigSide = new[] {widthX, lengthY, heightZ}.Max();
            double clamp_size;
            double clamp_L;
            string clamp_name;
            string clamp_block_bk;
            string clamp_block_op;
            string clamp_block_pr;

            if (_ui.RbColumn824.IsChecked == true || _ui.RbColumnCustom.IsChecked == true && bigSide <= 24)
            {
                clamp_size = 1d;
                clamp_name = "8/24";
                clamp_L = 36d;
                clamp_block_bk = GlobalNames.WtClampPlanBack824;
                clamp_block_op = GlobalNames.WtClampPlanOp824;
                clamp_block_pr = GlobalNames.WtClampProfile824;
                var clampPlanBack824Location = new XYZ(-1, 1, 0);
                var clampPlanOp824Location = new XYZ(0, 0, 0);
                _doc.Create.NewFamilyInstance(clampPlanBack824Location,
                    GetFamilySymbolByName(GlobalNames.WtClampPlanBack824), draftingView);
                _doc.Create.NewFamilyInstance(clampPlanOp824Location,
                    GetFamilySymbolByName(GlobalNames.WtClampPlanOp824),
                    draftingView);
            }
            else if (_ui.RbColumn2436.IsChecked == true || _ui.RbColumnCustom.IsChecked == true && bigSide <= 36)
            {
                clamp_size = 2;
                clamp_L = 48;
                clamp_name = "12/36";
                clamp_block_bk = GlobalNames.WtClampPlanBack1236;
                clamp_block_op = GlobalNames.WtClampPlanOp1236;
                clamp_block_pr = GlobalNames.WtClampProfile1236;
            }
            else if (_ui.RbColumn3648.IsChecked == true || _ui.RbColumnCustom.IsChecked == true && bigSide <= 48)
            {
                clamp_size = 3;
                clamp_L = 60;
                clamp_name = "24/48";
                clamp_block_bk = GlobalNames.WtClampPlanBack2448;
                clamp_block_op = GlobalNames.WtClampPlanOp2448;
                clamp_block_pr = GlobalNames.WtClampProfile2448;
            }
            else throw new ArgumentException("Error: Column is too wide (>48\") in one dimension");

            if (_ui.WindowX.IsChecked != true && _ui.WindowY.IsChecked != true) return;
            if (_ui.WindowY.IsChecked != true) clamp_block_pr += "_FLIPPED";
            clamp_block_op += "_WIN";
            swing_ang = -Math.PI / 4;
            const string filepath = GlobalNames.WtFileLocationPrefix + "Columns\\clamp_matrix.csv";
            var clamp_matrix = ImportMatrix(filepath);
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
            return new UV((maxO.U - minO.U) / 2, (maxO.V - minO.V) / 2);
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