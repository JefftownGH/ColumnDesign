using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ColumnDesign.Modules;
using ColumnDesign.UI;
using ColumnDesign.ViewModel;
using static ColumnDesign.Modules.ConvertFeetInchesToNumber;
using static ColumnDesign.Modules.ConvertNumberToFeetInches;
using static ColumnDesign.Modules.ImportMatrixFunction;
using static ColumnDesign.Modules.ReadSizesFunction;
using static ColumnDesign.Modules.UpdatePly_Function;
using static ColumnDesign.Modules.zScissorClampModule;

namespace ColumnDesign.Methods
{
    public static class Methods
    {
        private static ColumnCreatorView _ui;
        private static ColumnCreatorViewModel _vm;
        private static Document _doc;
        private static UIDocument _uiDoc;
        private static double win_clamp_top_max = 5;
        private static int win_clamp_bot_max = 4;
        private const int bot_clamp_gap = 8;
        private static double WinPos;
        private const double WinGap = 6;
        private const double WinStudOff = 1;

        public static void CreateGates(UIDocument uiDoc, ColumnCreatorView ui, ColumnCreatorViewModel vm)
        {
            _uiDoc = uiDoc;
            _doc = uiDoc.Document;
            _ui = ui;
            _vm = vm;
            const DrawingTypes type = DrawingTypes.Gates;
            using var tr = new Transaction(_doc, $"Create new {type.ToString().ToLower()} sheet");
            tr.Start();
            LoadFamilies(_doc);
            var sheet = CreateSheet(_doc, ui, DrawingTypes.Gates);
            var draftingView = CreateDraftingView(_doc, sheet.Name, sheet.SheetNumber);
            var zero = CalculateDraftingLocation(_doc, draftingView);
            Viewport.Create(_doc, sheet.Id, draftingView.Id, XYZ.Zero);
            CreateClamps(draftingView);
            _doc.Delete(zero.Id);
            tr.Commit();
            uiDoc.ActiveView = sheet;
        }

        public static void CreateScissors(UIDocument uiDoc, ColumnCreatorView ui)
        {
            _uiDoc = uiDoc;
            _doc = uiDoc.Document;
            _ui = ui;
            const DrawingTypes type = DrawingTypes.Scissors;
            using var tr = new Transaction(_doc, $"Create new {type.ToString().ToLower()} sheet");
            tr.Start();
            LoadFamilies(_doc);
            var sheet = CreateSheet(_doc, ui, DrawingTypes.Scissors);
            var draftingView = CreateDraftingView(_doc, sheet.Name, sheet.SheetNumber);
            var zero = CalculateDraftingLocation(_doc, draftingView);
            Viewport.Create(_doc, sheet.Id, draftingView.Id, XYZ.Zero);
            CreateScissorClamp(_ui, _uiDoc, draftingView);
            _doc.Delete(zero.Id);
            tr.Commit();
            uiDoc.ActiveView = sheet;
        }

        private static void CreateClamps(View draftingView)
        {
            double x;
            double y;
            double z;
            int n_col;
            int n_studs_x;
            int n_studs_y;
            int n_clamps = 0;
            int n_top_clamps;
            int n_reinf_angles;
            int stud_type;
            int clamp_size;
            int brace_size;
            int col_wt;
            string ply_name;
            string stud_name;
            string stud_name_full;
            string stud_block;
            string stud_block_spax;
            string stud_block_bolt;
            string stud_face_block;
            string SQ_NAME;
            string SQ_NAME_SLINGS;
            bool window = false;
            bool WinX;
            bool WinY;
            var ptTemp = new double[2];

            WinX = _vm.WindowX;
            WinY = _vm.WindowY;
            if (WinX || WinY) window = true;
            WinPos = ConvertToNum(_ui.WinDim2.Text);
            int[,] stud_matrix;
            var ply_thk = 0.75;
            var chamf_thk = 0.75;
            var min_stud_gap = 2.125;
            var max_wt_single_top = 2000;
            var stud_base_gap = 0.25;
            double bot_clamp_gap = 8;
            var stud_start_offset = 0.125;

            x = ConvertToNum(_ui.WidthX.Text);
            y = ConvertToNum(_ui.LengthY.Text);
            z = ConvertToNum(_ui.HeightZ.Text);
            n_col = (int) ConvertToNum(_ui.Quantity.Text);
            ply_name = _ui.PlywoodType.Text;
            if (WinX)
            {
                x = ConvertToNum(_ui.LengthY.Text);
                y = ConvertToNum(_ui.WidthX.Text);
                WinX = false;
                WinY = true;
            }

            if (z <= 16 * 12)
            {
                stud_type = 1;
                stud_name = "2X4";
                stud_name_full = "2X4";
                stud_block = "VBA_2X4";
                stud_block_spax = "VBA_2X4_SPAX";
                stud_block_bolt = "VBA_2X4_BOLT";
                stud_face_block = "VBA_2X4_FACE";
            }
            else if (z > 16 * 12)
            {
                stud_type = 2;
                stud_name = "LVL";
                stud_name_full = "1.5\" X 3.5\" LVL";
                stud_block = "VBA_LVL";
                stud_block_spax = "VBA_LVL_SPAX";
                stud_block_bolt = "VBA_LVL_BOLT";
                stud_face_block = "VBA_LVL_FACE";
            }
            else
            {
                throw new Exception("Error: Invalid stud type.");
            }

            string stud_block_bolt_hidden = "";
            string stud_block_spax_hidden;

            if (window)
            {
                stud_block_spax_hidden = stud_block_spax + "_HIDDEN";
                stud_block_bolt_hidden = stud_block_bolt + "_HIDDEN";
            }

            if (window && (z - WinPos) > 48)
            {
                var dialog = new TaskDialog("Warning")
                {
                    MainContent = "Pour window exceeds 48\" in height. This is allowed but not recommended",
                    AllowCancellation = true,
                    CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel
                };
                if (dialog.Show() == TaskDialogResult.Cancel) return;
            }

            if (window && _ui.Picking.IsChecked == true)
            {
                var dialog = new TaskDialog("Warning")
                {
                    MainContent =
                        "You specified a picking loop instead of a regular squaring corner to be used and you have a pour window. The picking loop will be removed from 1 corner to allow the gates clamp to swing open",
                    AllowCancellation = true,
                    CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel
                };
                if (dialog.Show() == TaskDialogResult.Cancel) return;
            }

            if (_ui.Picking.IsChecked == true && (x < 14 || y < 14))
            {
                throw new Exception(
                    "Columns with a side <14'' use regular squaring corners rotated 180 degrees. Picking loop squaring corners cannot be rotated 180 degrees.");
            }

            if (((int) x - x) != 0 || ((int) y - y) != 0)
            {
                var dialog = new TaskDialog("Warning")
                {
                    MainContent =
                        "Column plan dimensions are not divisible by 1 inch. Gates clamps have holes spaced 1 inch on center\nDo you know what you're doing?",
                    AllowCancellation = true,
                    CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
                };
                if (dialog.Show() == TaskDialogResult.No) return;
            }

            if (x < 14 || y < 14)
            {
                SQ_NAME = "VBA_GATES_SQUARING_CORNER_INV";
                SQ_NAME_SLINGS = "VBA_GATES_SQUARING_CORNER_INV";
            }
            else
            {
                SQ_NAME = "VBA_GATES_SQUARING_CORNER";
                SQ_NAME_SLINGS = "VBA_GATES_SQUARING_CORNER_SLINGS";
            }

            double[] insertionPnt = new double[2];


            // #####################################################################################
            //                             S T U D S
            // #####################################################################################
            int row_num = 0;
            int col_num_x = 0;
            int col_num_y = 0;
            int n_studs_total;
            stud_matrix = ImportMatrix(@$"{GlobalNames.FilesLocationPrefix}Columns\n_stud_matrix.csv");
            for (var i = 0; i < stud_matrix.GetLength(0); i++)
            {
                if (stud_matrix[i, 0] >= z)
                {
                    row_num = i;
                    break;
                }
            }

            for (var i = 0; i < stud_matrix.GetLength(1); i++)
            {
                if (Math.Abs(stud_matrix[0, i] - x) < 0.001)
                {
                    col_num_x = i;
                    break;
                }
            }

            for (var i = 0; i < stud_matrix.GetLength(1); i++)
            {
                if (Math.Abs(stud_matrix[0, i] - y) < 0.001)
                {
                    col_num_y = i;
                    break;
                }
            }

            if (col_num_x == 0)
            {
                for (var i = 0; i < stud_matrix.GetLength(1); i++)
                {
                    if (stud_matrix[0, i] > x)
                    {
                        if (stud_matrix[row_num, i] >= stud_matrix[row_num, i - 1])
                        {
                            col_num_x = i;
                            break;
                        }
                        else
                        {
                            col_num_x = i - 1;
                            break;
                        }
                    }
                }
            }

            if (col_num_y == 0)
            {
                for (var i = 0; i < stud_matrix.GetLength(1); i++)
                {
                    if (stud_matrix[0, i] > y)
                    {
                        if (stud_matrix[row_num, i] >= stud_matrix[row_num, i - 1])
                        {
                            col_num_y = i;
                            break;
                        }
                        else
                        {
                            col_num_y = i - 1;
                            break;
                        }
                    }
                }
            }

            n_studs_x = stud_matrix[row_num, col_num_x];
            n_studs_y = stud_matrix[row_num, col_num_y];
            n_studs_total = n_studs_x * 2 + n_studs_y * 2;

            double avg_gap_x;
            double avg_gap_y;
            double push_studs_x = 0;
            int push_studs_y = 0;
            avg_gap_x = (x + ply_thk - stud_start_offset - n_studs_x * 3.5) / (n_studs_x - 1);
            if (avg_gap_x < min_stud_gap)
            {
                push_studs_x = 1;
            }

            avg_gap_y = (y + ply_thk - stud_start_offset - n_studs_y * 3.5) / (n_studs_y - 1);
            if (avg_gap_y < min_stud_gap)
            {
                push_studs_y = 1;
            }

            double[] stud_spacing_x = new double[n_studs_x];
            double[] stud_spacing_y = new double[n_studs_y];
            double available_2x2_x;
            double available_2x2_y;
            double min_2x2_gap = 1.625;
            int AFB_x2;
            for (var i = 0; i < n_studs_x; i++)
            {
                stud_spacing_x[i] = stud_start_offset + i * (3.5 + avg_gap_x);
            }

            for (var i = 0; i < n_studs_y; i++)
            {
                stud_spacing_y[i] = stud_start_offset + i * (3.5 + avg_gap_y);
            }

            if (Math.Abs(push_studs_x - 1) < 0.001)
            {
                if (avg_gap_x * (n_studs_x - 1) >= 2 * min_stud_gap)
                {
                    stud_spacing_x[1] = stud_start_offset + 3.5 + min_stud_gap;
                    stud_spacing_x[stud_spacing_x.Length - 2] =
                        stud_spacing_x[stud_spacing_x.Length - 1] - 3.5 - min_stud_gap;
                    AFB_x2 = 1;
                }
                else
                {
                    stud_spacing_x[stud_spacing_x.Length - 2] =
                        stud_spacing_x[stud_spacing_x.Length - 1] - 3.5 - min_stud_gap;
                    AFB_x2 = 0;
                }
            }

            if (push_studs_y == 1)
            {
                if (avg_gap_y * (n_studs_y - 1) >= 2 * min_stud_gap)
                {
                    stud_spacing_y[1] = stud_start_offset + 3.5 + min_stud_gap;
                    stud_spacing_y[stud_spacing_y.Length - 2] =
                        stud_spacing_y[stud_spacing_y.Length - 1] - 3.5 - min_stud_gap;
                }
                else
                {
                    stud_spacing_y[stud_spacing_y.Length - 2] =
                        stud_spacing_y[stud_spacing_y.Length - 1] - 3.5 - min_stud_gap;
                }
            }

            for (var i = 0; i < n_studs_x - 1; i++)
            {
                if (stud_spacing_x[i + 1] - stud_spacing_x[i] < 3.5)
                {
                    stud_spacing_x[i] = stud_spacing_x[i + 1] - 3.5;
                }
            }

            for (var i = 0; i < n_studs_y - 1; i++)
            {
                if (stud_spacing_y[i + 1] - stud_spacing_y[i] < 3.5)
                {
                    stud_spacing_y[i] = stud_spacing_y[i + 1] - 3.5;
                }
            }

            for (var i = 0; i < n_studs_x - 1; i++)
            {
                if (stud_spacing_x[i + 1] - stud_spacing_x[i] < 3.5)
                {
                    TaskDialog.Show("Error", "Error: Studs overlap, check and manually correct drawing.");
                    goto EndOfStudChecks;
                }
            }

            for (var i = 0; i < n_studs_y - 1; i++)
            {
                if (stud_spacing_y[i + 1] - stud_spacing_y[i] < 3.5)
                {
                    TaskDialog.Show("Error", "Error: Studs overlap, check and manually correct drawing.");
                    goto EndOfStudChecks;
                }
            }

            if (window)
            {
                for (var i = 0; i < n_studs_x - 1; i++)
                {
                    if (stud_spacing_x[i + 1] - stud_spacing_x[i] < 5)
                    {
                        TaskDialog.Show("Error",
                            "Warning: Insufficient clearance between studs for a 2x2 window lock. Check and manually correct drawing if necessary.");
                        goto EndOfStudChecks;
                    }
                }

                for (var i = 0; i < n_studs_y - 1; i++)
                {
                    if (stud_spacing_y[i + 1] - stud_spacing_y[i] < 5)
                    {
                        TaskDialog.Show("Error",
                            "Warning: Insufficient clearance between studs for a 2x2 window lock. Check and manually correct drawing if necessary.");
                        goto EndOfStudChecks;
                    }
                }
            }

            EndOfStudChecks:

            // #####################################################################################
            //                           P L Y W O O D
            // #####################################################################################

            double ply_width_x;
            double ply_width_y;
            int ply_bot_n;
            double ply_top_ht_min = 6;
            double max_ply_ht;
            var ply_seams = new double[1];
            var ply_seams_win = new double[1];

            ply_seams = ReadSizes(_ui.BoxPlySeams.Text);
            ply_seams_win = ReadSizes(_ui.BoxPlySeams.Text);
            if (ValidatePlySeams(_ui, ply_seams, x, y, z) == 0)
            {
                throw new Exception("Plywood layout invalid. You should never see this message...How did you do this?");
            }

            var temp_1 = 0d;

            for (var i = 0; i < ply_seams.Length; i++)
            {
                temp_1 += ply_seams[i];
                if (Math.Abs(temp_1 - WinPos) < 0.001)
                {
                    break;
                }
                else if (temp_1 > WinPos)
                {
                    ply_seams_win[i] -= temp_1 - WinPos;
                    Array.Resize(ref ply_seams_win, ply_seams.Length + 1);
                    ply_seams_win[i + 1] = ply_seams[i] - ply_seams_win[i];
                    if (i < ply_seams.Length - 1)
                    {
                        for (var j = i + 1; j < ply_seams.Length; j++)
                        {
                            ply_seams_win[j + 1] = ply_seams[j];
                        }
                    }

                    break;
                }
            }

            ply_width_x = x + 1.5;
            ply_width_y = y + 1.5;
            int n_studs_w = 0;
            double ply_width_w = 0;
            double[] stud_spacing_w = new double[1];
            int n_studs_e;
            double ply_width_e = 0;
            double[] stud_spacing_e = new double[1];
            if (WinX)
            {
                n_studs_w = n_studs_x;
                n_studs_e = n_studs_x;
                ply_width_w = ply_width_x;
                ply_width_e = ply_width_x;
                Array.Resize(ref stud_spacing_w, stud_spacing_x.Length);
                Array.Resize(ref stud_spacing_e, stud_spacing_x.Length);
                for (var i = 0; i < stud_spacing_x.Length; i++)
                {
                    stud_spacing_w[i] = stud_spacing_x[i];
                    stud_spacing_e[i] = stud_spacing_x[i];
                }
            }
            else if (WinY)
            {
                n_studs_w = n_studs_y;
                n_studs_e = n_studs_y;
                ply_width_w = ply_width_y;
                ply_width_e = ply_width_y;
                Array.Resize(ref stud_spacing_w, stud_spacing_y.Length);
                Array.Resize(ref stud_spacing_e, stud_spacing_y.Length);
                for (var i = 0; i < stud_spacing_y.Length; i++)
                {
                    stud_spacing_w[i] = stud_spacing_y[i];
                    stud_spacing_e[i] = stud_spacing_y[i];
                }
            }
            else
            {
                n_studs_e = n_studs_x;
                ply_width_e = ply_width_x;
                Array.Resize(ref stud_spacing_e, stud_spacing_x.Length);
                for (var i = 0; i < stud_spacing_x.Length; i++)
                {
                    stud_spacing_e[i] = stud_spacing_x[i];
                }
            }

            if (ply_width_x > 48 || ply_width_y > 48)
            {
                max_ply_ht = 48;
            }
            else
            {
                max_ply_ht = 96;
            }

            if (window)
            {
                var temp_2 = 0d;
                var ply_cnt = 0;
                temp_1 = 0;
                for (var i = 0; i < ply_seams_win.Length - 1; i++)
                {
                    temp_1 += ply_seams_win[i];
                    if (temp_1 >= WinPos)
                    {
                        ply_cnt++;
                        temp_2 += ply_seams_win[i + 1];
                    }
                }

                if (temp_2 <= max_ply_ht)
                {
                    ply_seams_win[ply_seams_win.Length - ply_cnt] = temp_2;
                    Array.Resize(ref ply_seams_win, ply_seams_win.Length - ply_cnt + 1);
                }
            }

            ply_bot_n = ply_seams.Count(t => Math.Abs(t - max_ply_ht) < 0.001);

            int unique_plys;
            double[] ply_widths = new double[2];
            double[,] ply_cuts = new double[3, 1];
            unique_plys = 0;
            ply_widths[0] = ply_width_x;
            ply_widths[1] = ply_width_y;
            for (var i = 0; i < ply_seams.Length; i++)
            {
                for (var j = 0; j < ply_widths.Length; j++)
                {
                    for (var l = 0; l < ply_cuts.GetLength(1); l++)
                    {
                        if (Math.Abs(ply_cuts[0, l] - ply_widths[j]) < 0.001 &&
                            Math.Abs(ply_cuts[1, l] - ply_seams[i]) < 0.001)
                        {
                            ply_cuts[2, l] += 2 - (WinX == true ? 1 : 0) * (j == 0 ? 1 : 0) -
                                              (WinY == true ? 1 : 0) * (j == 1 ? 1 : 0);
                            break;
                        }
                        else if (l == ply_cuts.GetLength(1) - 1)
                        {
                            unique_plys++;
                            ply_cuts = ResizeArray<double>(ply_cuts, 3, unique_plys);
                            ply_cuts[0, unique_plys - 1] = ply_widths[j];
                            ply_cuts[1, unique_plys - 1] = ply_seams[i];
                            ply_cuts[2, unique_plys - 1] = 2 - (WinX == true ? 1 : 0) * (j == 0 ? 1 : 0) -
                                                           (WinY == true ? 1 : 0) * (j == 1 ? 1 : 0);
                            break;
                        }
                    }
                }
            }

            if (ply_cuts[0, ply_cuts.GetLength(1) - 1] == 0)
            {
                ply_cuts = ResizeArray<double>(ply_cuts, 3, ply_cuts.GetLength(1));
                unique_plys--;
            }

            for (var i = 0; i < ply_seams_win.Length; i++)
            {
                for (var j = 0; j < ply_cuts.GetLength(1); j++)
                {
                    if (Math.Abs(ply_cuts[0, j] - ply_width_w) < 0.001 &&
                        Math.Abs(ply_cuts[1, j] - ply_seams_win[i]) < 0.001)
                    {
                        ply_cuts[2, j]++;
                        break;
                    }
                    else if (j == ply_cuts.GetLength(1))
                    {
                        unique_plys++;
                        ply_cuts = ResizeArray(ply_cuts, 3, unique_plys);
                        ply_cuts[0, unique_plys] = ply_width_w;
                        ply_cuts[1, unique_plys] = ply_seams_win[i];
                        ply_cuts[2, unique_plys]++;
                        break;
                    }
                }
            }

            if (ply_cuts[0, ply_cuts.GetLength(1) - 1] == 0)
            {
                ply_cuts = ResizeArray(ply_cuts, 3, ply_cuts.GetLength(1));
                unique_plys--;
            }

            // var msg = "";
            // for (var i = 0; i < ply_cuts.GetLength(1); i++)
            // {
            //     for (var j = 0; j < ply_cuts.GetLength(0); j++)
            //     {
            //         msg += $"{ply_cuts[j, i]}\t";
            //     }
            //
            //     msg += "\n";
            // }
            // TaskDialog.Show("Message", msg);

            // #####################################################################################
            //                                       C L A M P S
            // #####################################################################################
            var clamp_spacing = new int[] { };
            double swing_ang = 0;
            double win_clamp_top_max = 5;
            double win_clamp_bot_max = 4;
            double long_side = 0;
            double clamp_L;
            string clamp_name;
            string clamp_block_bk;
            string clamp_block_op;
            string clamp_block_pr;
            var bigSide = new[] {x, y}.Max();
            if (_ui.RbColumn824.IsChecked == true || _ui.RbColumnCustom.IsChecked == true && bigSide < 24)
            {
                clamp_size = 1;
                clamp_name = "8/24";
                clamp_L = 36;
                clamp_block_bk = "VBA_8-24_CLAMP_PLAN_BACK";
                clamp_block_op = "VBA_8-24_CLAMP_PLAN_OP";
                clamp_block_pr = "VBA_8-24_CLAMP_PROFILE";
            }
            else if (_ui.RbColumn2436.IsChecked == true || _ui.RbColumnCustom.IsChecked == true && bigSide <= 36)
            {
                clamp_size = 2;
                clamp_L = 48;
                clamp_name = "12/36";
                clamp_block_bk = "VBA_12-36_CLAMP_PLAN_BACK";
                clamp_block_op = "VBA_12-36_CLAMP_PLAN_OP";
                clamp_block_pr = "VBA_12-36_CLAMP_PROFILE";
            }
            else if (_ui.RbColumn3648.IsChecked == true || _ui.RbColumnCustom.IsChecked == true && bigSide <= 48)
            {
                clamp_size = 3;
                clamp_L = 60;
                clamp_name = "24/48";
                clamp_block_bk = "VBA_24-48_CLAMP_PLAN_BACK";
                clamp_block_op = "VBA_24-48_CLAMP_PLAN_OP";
                clamp_block_pr = "VBA_24-48_CLAMP_PROFILE";
            }
            else throw new Exception("Error: Column is too wide (>48\") in one dimension");

            if (window)
            {
                if (WinY)
                {
                    clamp_block_pr += "_FLIPPED";
                }

                clamp_block_op += "_WIN";
                swing_ang = -Math.PI / 4;
            }

            var clamp_matrix = ImportMatrix(@$"{GlobalNames.FilesLocationPrefix}Columns\clamp_matrix.csv");
            long_side = x <= y ? y : x;

            row_num = 0;
            for (var i = 0; i < clamp_matrix.GetLength(0); i++)
            {
                if ((clamp_matrix[i, 0] >= long_side))
                {
                    row_num = i;
                    break;
                }
            }

            Array.Resize(ref clamp_spacing, clamp_matrix.GetLength(1) - 1);
            var k = 0;
            for (var i = 0; i < clamp_matrix.GetLength(1) - 1; i++)
            {
                if (clamp_matrix[row_num, i + 1] == 0)
                {
                    Array.Resize(ref clamp_spacing, k);
                    break;
                }

                clamp_spacing[i] = clamp_matrix[row_num, i + 1];
                k++;
            }

            var ht_rem = z - bot_clamp_gap;
            for (var i = 0; i < clamp_spacing.Length - 1; i++)
            {
                if (clamp_spacing[i] <= ht_rem)
                {
                    n_clamps++;
                    ht_rem -= clamp_spacing[i];
                }
                else
                {
                    n_clamps++;
                    break;
                }
            }

            Array.Resize(ref clamp_spacing, n_clamps + 1);
            clamp_spacing[clamp_spacing.Length - 1] = (int) bot_clamp_gap;
            var tot_temp = clamp_spacing.Sum();
            clamp_spacing[clamp_spacing.Length - 2] -= tot_temp - (int) z;
            TestClampSpacing:
            var infExit = 0;
            for (var i = 1; i < clamp_spacing.Length; i++)
            {
                if (clamp_spacing[i] < 8)
                {
                    clamp_spacing[i - 1] -= 8 - clamp_spacing[i];
                    clamp_spacing[i] = 8;


                    infExit++;
                    if (infExit > 100)
                    {
                        throw new TimeoutException("Error: Infinite loop encountered while computing clamp spacings.");
                    }
                    else
                    {
                        goto TestClampSpacing;
                    }
                }
            }

            var clamp_spacing_con = new int[clamp_spacing.Length];
            clamp_spacing_con[0] = (int) z - clamp_spacing[0];
            for (var i = 1; i < clamp_spacing.Length; i++)
            {
                clamp_spacing_con[i] = clamp_spacing_con[i - 1] - clamp_spacing[i];
            }

            var temp_value = 0;
            WinPos = ConvertToNum(_ui.WinDim2.Text);
            if (window)
            {
                for (var i = 0; i < clamp_spacing_con.Length - 1; i++)
                {
                    if (clamp_spacing_con[i] <= WinPos)
                    {
                        if (WinPos - clamp_spacing[i] > win_clamp_bot_max)
                        {
                            n_clamps++;
                            Array.Resize(ref clamp_spacing_con, n_clamps + 1);
                            for (var j = i; j < n_clamps - 1; j++)
                            {
                                clamp_spacing_con[n_clamps - j + i - 1] = clamp_spacing_con[n_clamps - j + i - 2];
                            }

                            clamp_spacing_con[i] = (int) WinPos - (int) win_clamp_bot_max;
                            clamp_spacing_con[clamp_spacing_con.Length - 1] = 0;
                            if (clamp_spacing_con[i + 1] - clamp_spacing_con[i] < 8)
                            {
                                clamp_spacing_con[i + 1] = (clamp_spacing_con[i + 2] + clamp_spacing_con[i]) / 2;
                            }

                            goto LowerWinClampSet;
                        }

                        if (WinPos - clamp_spacing_con[i] <= win_clamp_bot_max)
                        {
                            clamp_spacing_con[i] = (int) WinPos - (int) win_clamp_bot_max;
                            goto LowerWinClampSet;
                        }
                    }
                }

                LowerWinClampSet:
                int io;
                for (var i = 0; i < clamp_spacing_con.Length - 1; i++)
                {
                    io = clamp_spacing_con.Length - i - 2;
                    if (clamp_spacing_con[io] > WinPos)
                    {
                        if (clamp_spacing_con[io] - WinPos > win_clamp_top_max)
                        {
                            n_clamps++;
                            Array.Resize(ref clamp_spacing_con, n_clamps + 1);
                            clamp_spacing_con[clamp_spacing_con.Length - 1] = (int) WinPos + (int) win_clamp_top_max;
                            Array.Sort(clamp_spacing_con, (a, b) => b.CompareTo(a));

                            if (clamp_spacing_con[io] - clamp_spacing_con[io + 1] < 8)
                            {
                                if (io >= 2)
                                {
                                    clamp_spacing_con[io] =
                                        (clamp_spacing_con[io - 1] + clamp_spacing_con[io + 1]) / 2;
                                }
                                else
                                {
                                    clamp_spacing_con[io] = clamp_spacing_con[io + 1] + 8;
                                }
                            }

                            goto UpperWinClampSet;

                            if (clamp_spacing_con[io] - WinPos <= win_clamp_top_max &&
                                clamp_spacing_con[io] - WinPos >= 0)
                            {
                                clamp_spacing_con[io] = (int) WinPos + (int) win_clamp_top_max;
                                goto UpperWinClampSet;
                            }
                        }
                    }
                }

                UpperWinClampSet:
                Array.Resize(ref clamp_spacing, n_clamps + 2);
                clamp_spacing[0] = (int) z - clamp_spacing_con[0];
                for (var i = 1; i < n_clamps; i++)
                {
                    clamp_spacing[i] = clamp_spacing_con[i - 1] - clamp_spacing_con[i];
                }

                clamp_spacing[n_clamps + 1] = clamp_spacing_con[clamp_spacing_con.Length - 2];
            }

            n_reinf_angles = 0;
            if (x > 40)
            {
                for (var i = 0; i < clamp_spacing_con.Length - 1; i++)
                {
                    if (z - clamp_spacing_con[i] >= 87)
                    {
                        n_reinf_angles++;
                    }
                }
            }

            if (y > 40)
            {
                for (var i = 0; i < clamp_spacing_con[i]; i++)
                {
                    if (z - clamp_spacing_con[i] >= 87)
                    {
                        n_reinf_angles++;
                    }
                }
            }

            n_reinf_angles *= 2;
            if (window)
            {
                if (clamp_spacing_con[1] < WinPos)
                {
                    throw new Exception(
                        "Warning:\nPour window is only secured by 1 clamp. Consult with an engineering manager");
                }
            }

            // #####################################################################################
            //                            M I S C E L L A N E O U S
            // #####################################################################################

            int n_bolts;
            int n_screws;
            int n_chamf;
            double brace_L_stored;
            string brace_name;
            string brace_block;

            var chamf_length = 12;
            if (z < 160)
            {
                brace_size = 1;
                brace_L_stored = 79.6;
                brace_name = "7'-TO-11'";
                brace_block = "VBA_AFB_7-11";
            }
            else
            {
                brace_size = 2;
                brace_L_stored = 128.6;
                brace_name = "11'-TO-19'";
                brace_block = "VBA_AFB_11-19";
            }

            int n = 1;
            while (clamp_spacing_con[n] > z * 0.7)
            {
                n++;
            }

            int brace_clamp;
            if (Math.Abs(clamp_spacing_con[n] - z * 0.7) < Math.Abs(clamp_spacing_con[n - 1] - z * 0.7))
            {
                brace_clamp = n;
            }
            else
            {
                brace_clamp = n - 1;
            }

            while (clamp_spacing_con[brace_clamp] + 0.9289 < brace_L_stored + 4)
            {
                brace_clamp--;
            }

            int chain_clamp = 0;
            while (clamp_spacing_con[chain_clamp] > clamp_spacing_con[brace_clamp] + 0.9289 - brace_L_stored + 1)
            {
                chain_clamp++;
            }

            chain_clamp -= 1;

            col_wt = CalcWeightFunction.wt_total(x, y, z, n_studs_x, n_studs_y, stud_type, n_clamps, clamp_size,
                brace_size);
            if (col_wt <= max_wt_single_top)
            {
                n_top_clamps = 1;
            }
            else
            {
                n_top_clamps = 2;
            }

            n_bolts = (n_studs_total * n_top_clamps) + (2 * (n_clamps - n_top_clamps));
            n_screws = (n_studs_total - 2) * (n_clamps - n_top_clamps);

            if (z == 144)
            {
                n_chamf = 4;
            }
            else
            {
                n_chamf = (int) (4 * Math.Floor((z / 12) / chamf_length) +
                                 Math.Floor(4 / (12 / (z / 12 - chamf_length * Math.Floor((z / 12) / chamf_length)))));
                if (z < chamf_length * 12)
                {
                    n_chamf = 4;
                }
            }

            // #####################################################################################
            //                         T E X T
            // #####################################################################################
            var qty_text = "";
            var pt_o = new double[2];
            var pt1 = new double[2];
            var pt2 = new double[2];
            var pt3 = new double[2];
            pt_o[0] = 20.23568507;
            pt_o[1] = 5.17943146;
            pt1[0] = pt_o[0] + 448.4;
            pt1[1] = pt_o[1] + 360;
            pt2[0] = pt1[0] + 0;
            pt2[1] = pt1[1] + 7;
            pt3[0] = pt1[0] + 0;
            pt3[1] = pt1[1] - 55;
            var text = $"• COLUMN SIZE = {x}\" X {y}\"\n\n" +
                       $"• NUMBER OF COLUMN FORMS = {n_col}-EA\n\n" +
                       $"• COLUMN FORM WEIGHT (APPROXIMATE) = {col_wt}-LBS\n\n" +
                       $"• PLYWOOD = 3/4'' PLYFORM (\"{ply_name}\"), CLASS-1 (MIN)\n\n" +
                       "• COLUMN FORMS AND CLAMP SPACING LAYOUTS FOR L4 X 3 X 1/4 GATES LOK-FAST COLUMN CLAMPS ARE DESIGNED FOR A POUR RATE = FULL LIQUID HEAD U.N.O.\n\n" +
                       "• CONTACT THE MCC ENGINEER PRIOR TO ANY CHANGES OR MODIFICATIONS TO THE DETAILS ON THIS SHEET.";
            var textNoteOptions = new TextNoteOptions
            {
                VerticalAlignment = VerticalTextAlignment.Top,
                HorizontalAlignment = HorizontalTextAlignment.Left,
                TypeId = new FilteredElementCollector(_doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .First(q => q.Name == "2.0 mm").Id
            };
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt1),
                UnitUtils.ConvertToInternalUnits(95, DisplayUnitType.DUT_MILLIMETERS), text, textNoteOptions);

            for (int i = 0; i < ply_cuts.GetLength(1); i++)
            {
                qty_text +=
                    $"• ({ply_cuts[2, i] * n_col}-EA) = ({n_col}-COL) X ({ply_cuts[2, i]}-EA/COL) @ {ConvertFtIn(ply_cuts[0, i])} WIDE X {ConvertFtIn(ply_cuts[1, i])} LONG 3/4'' PLYWOOD\n";
            }

            qty_text = $"PLYWOOD\n{qty_text}\n";

            if (window == false)
            {
                qty_text +=
                    $"STUDS\n• ({n_studs_total * n_col}-EA) = ({n_col}-COL) X ({n_studs_total}-EA/COL) @ {ConvertFtIn(z - 0.25)} {stud_name_full}";
            }
            else
            {
                qty_text += "STUDS\n" +
                            $"• ({(n_studs_total - n_studs_w) * n_col}-EA) = ({n_col}-COL) X ({(n_studs_total - n_studs_w)}-EA/COL) @ {ConvertFtIn(z - 0.25)} {stud_name_full}\n";
                qty_text +=
                    $"• ({n_studs_w * n_col}-EA) = ({n_col}-COL) X ({n_studs_w}-EA/COL) @ {ConvertFtIn(WinPos - WinStudOff - 0.25)} {stud_name_full}\n";
                qty_text +=
                    $"• ({n_studs_w * n_col}-EA) = ({n_col}-COL) X ({n_studs_w}-EA/COL) @ {ConvertFtIn(z - WinPos - WinStudOff)} {stud_name_full}";
            }

            qty_text += "\n\nCOLUMN CLAMPS\n";

            qty_text +=
                $"• ({n_clamps * n_col}-EA) = ({n_col}-COL) X ({n_clamps}-EA/COL) @ GATES {clamp_name} LOK-FAST CLAMP ASSEMBLIES (SETS).";
            if (n_reinf_angles > 0)
            {
                qty_text +=
                    $"\n• ({n_reinf_angles * n_col}-EA) = ({n_col}-COL) X ({n_reinf_angles}-EA/COL) @ GATES REINFORCING ANGLES";
            }

            var n_nuts = 0;
            qty_text += "\n\nFASTENERS\n" +
                        $"•   ({n_bolts * n_col}-EA) = ({n_col}-COL) X ({n_bolts}-EA/COL) @ 5/16'' X 3'' GATES FLAT HEAD BOLTS\n" +
                        $"•   ({n_bolts * n_col}-EA) = ({n_col}-COL) X ({n_bolts}-EA/COL) @ 5/16''-18 UNC NYLOK LOCK NUTS\n" +
                        $"•   ({n_screws * n_col}-EA) = ({n_col}-COL) X ({n_screws}-EA/COL) @ 1/4'' X 2-3/8'' GATES SPAX POWERLAG SCREWS\n";
            if (n_top_clamps >= 2 || _ui.Picking.IsChecked == true)
            {
                qty_text +=
                    $"•   ({n_col * 2}-EA) = ({n_col}-COL) X (2-EA/COL) @ 1/2''-13 UNC X +/-36'' LONG ALL-THREADED ROD\n";
                n_nuts += 8;
            }

            if ((window && (n_top_clamps >= 2 || _ui.Picking.IsChecked == true) && clamp_spacing_con[1] > WinPos) ||
                (window && n_top_clamps < 2 && _ui.Picking.IsChecked == false))
            {
                qty_text +=
                    $"•   ({n_col}-EA) = ({n_col}-COL) X (2-EA/COL) @ 1/2''-13 UNC X +/-14'' LONG ALL-THREADED ROD\n";
                n_nuts += 4;
            }

            if (n_nuts > 0)
            {
                qty_text +=
                    $"•   ({n_nuts * n_col}-EA) = ({n_col}-COL) X ({n_nuts}-EA/COL) @ 1/2''-13 UNC NYLOK LOCK NUTS\n" +
                    $"•   ({n_nuts * n_col}-EA) = ({n_col}-COL) X ({n_nuts}-EA/COL) @ 1/2'' STANDARD FLAT WASHER\n\n";
            }

            qty_text += "3/4'' GATES PLASTIC CHAMFER (BASED ON 12' CHAMFER LENGTHS)\n" +
                        $"•   ({n_col * n_chamf}-EA) = ({n_col}-COL) X ({n_chamf}-EA/COL) @ {chamf_length}'-0'' LONG PIECES";
            qty_text += "\n\nGATES ADJUSTABLE FORM BRACES (INCLUDING FORM BASE PLATES)\n" +
                        $"•   ({n_col * 3}-EA) = ({n_col}-COL) X (3EA/COL) {brace_name} LONG GATES AFB\n\n";
            if (col_wt <= 2400)
            {
                qty_text += "HOISTING SLINGS\n" +
                            $"•   ({n_col * 2}-EA) = ({n_col}-COL) X (2EA/COL) ENDLESS ROUND SLINGS; LIFTEX P/N ''ENR1'', PURPLE.\n" +
                            "SWL = 2400-LBS. PER SLING IN CHOKER CONFIGURATION.";
            }
            else
            {
                qty_text += "HOISTING SLINGS\n" +
                            $"•   ({n_col * 2}-EA) = ({n_col}-COL) X (2EA/COL) ENDLESS ROUND SLINGS; LIFTEX P/N ''ENR2'', GREEN.\n" +
                            "SWL = 4800-LBS. PER SLING IN CHOKER CONFIGURATION.";
            }

            textNoteOptions = new TextNoteOptions
            {
                VerticalAlignment = VerticalTextAlignment.Top,
                HorizontalAlignment = HorizontalTextAlignment.Left,
                TypeId = new FilteredElementCollector(_doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .First(q => q.Name == "2.0 mm").Id
            };
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt3),
                UnitUtils.ConvertToInternalUnits(95, DisplayUnitType.DUT_MILLIMETERS), qty_text, textNoteOptions);

            // #####################################################################################
            //                                   D R A W I N G
            // #####################################################################################
            var ptA = new double[2];
            var ptB = new double[2];
            var ptW = new double[2];
            var ptE = new double[2];
            var pt4 = new double[2];
            var pt5 = new double[2];
            var pt6 = new double[2];
            var pt7 = new double[2];
            var pt8 = new double[2];
            var pt9 = new double[2];
            var pt10 = new double[2];
            var pt11 = new double[2];
            var pt12 = new double[2];
            var pt13 = new double[2];
            var pt14 = new double[2];
            var pt15 = new double[2];
            var pt16 = new double[2];
            var pt17 = new double[2];
            var pt18 = new double[2];
            var pt20 = new double[2];
            var pt21 = new double[2];
            var pt22 = new double[2];
            var pt23 = new double[2];
            var pt24 = new double[2];
            var pt25 = new double[2];
            var pt26 = new double[2];
            var pt27 = new double[2];
            var pt28 = new double[2];
            var pt29 = new double[2];
            var pt30 = new double[2];
            var pt31 = new double[2];
            var pt32 = new double[2];
            var pt33 = new double[2];
            var pt34 = new double[2];
            var pt_blk = new double[2];
            bool DrawB = false;
            bool DrawW = false;
            if (window == false && Math.Abs(x - y) < 0.001)
            {
                ptA[0] = pt_o[0] + 245;
                ptA[1] = pt_o[1] + 25;
                ptE[0] = pt_o[0] + 100;
                ptE[1] = pt_o[1] + 25;
            }
            else if (window == true && Math.Abs(x - y) < 0.001)
            {
                ptW[0] = pt_o[0] + 300;
                ptW[1] = pt_o[1] + 25;
                ptA[0] = pt_o[0] + 185;
                ptA[1] = pt_o[1] + 25;
                ptE[0] = pt_o[0] + 54.5;
                ptE[1] = pt_o[1] + 25;
                DrawW = true;
            }
            else if (window == false && Math.Abs(x - y) > 0.001)
            {
                ptA[0] = pt_o[0] + 300;
                ptA[1] = pt_o[1] + 25;
                ptB[0] = pt_o[0] + 185;
                ptB[1] = pt_o[1] + 25;
                ptE[0] = pt_o[0] + 54.5;
                ptE[1] = pt_o[1] + 25;
                DrawB = true;
            }
            else if (window == true && Math.Abs(x - y) > 0.001)
            {
                ptW[0] = pt_o[0] + 340 - 0.5 * ply_width_w;
                ptW[1] = pt_o[1] + 25;
                ptA[0] = pt_o[0] + 267.25 - 0.5 * ply_width_x;
                ptA[1] = pt_o[1] + 25;
                ptB[0] = pt_o[0] + 190 - 0.5 * ply_width_y;
                ptB[1] = pt_o[1] + 25;
                ptE[0] = pt_o[0] + 54.5 - 0.5 * (ply_width_w - 24);
                ptE[1] = pt_o[1] + 25;
                DrawB = true;
                DrawW = true;
            }

            FamilyInstance family;
            for (var i = 0; i < n_studs_x; i++)
            {
                pt6[0] = ptA[0] + ply_width_x + chamf_thk - 3.5 - stud_spacing_x[i];
                pt6[1] = ptA[1] + stud_base_gap;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt6), GetFamilySymbolByName(stud_face_block),
                    draftingView);
                RotateFamily(family, 90);
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(z - stud_base_gap, DisplayUnitType.DUT_DECIMAL_INCHES));
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Middle,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    Rotation = Math.PI / 2,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "2.0 mm").Id
                };
                ptTemp[0] = pt6[0] + 1.5;
                ptTemp[1] = pt6[1] + (z - stud_base_gap) / 2;
                switch (stud_face_block)
                {
                    case "VBA_2X4_FACE":
                        TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
                        break;
                    case "VBA_LVL_FACE":
                        TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "1.5 X 3.5 LVL", textNoteOptions);
                        break;
                }
            }

            if (DrawB)
            {
                for (var i = 0; i < n_studs_y; i++)
                {
                    pt6[0] = ptB[0] + ply_width_y + chamf_thk - 3.5 - stud_spacing_y[i];
                    pt6[1] = ptB[1] + stud_base_gap;
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt6), GetFamilySymbolByName(stud_face_block),
                        draftingView);
                    RotateFamily(family, 90);
                    family.LookupParameter("Distance1")
                        ?.Set(UnitUtils.ConvertToInternalUnits(z - stud_base_gap, DisplayUnitType.DUT_DECIMAL_INCHES));
                    textNoteOptions = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Middle,
                        HorizontalAlignment = HorizontalTextAlignment.Center,
                        Rotation = Math.PI / 2,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "2.0 mm").Id
                    };
                    ptTemp[0] = pt6[0] + 1.5;
                    ptTemp[1] = pt6[1] + (z - stud_base_gap) / 2;
                    switch (stud_face_block)
                    {
                        case "VBA_2X4_FACE":
                            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
                            break;
                        case "VBA_LVL_FACE":
                            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "1.5 X 3.5 LVL",
                                textNoteOptions);
                            break;
                    }
                }
            }

            if (DrawW)
            {
                for (var i = 0; i < n_studs_w; i++)
                {
                    pt6[0] = ptW[0] + ply_width_w + chamf_thk - 3.5 - stud_spacing_w[i];
                    pt6[1] = ptW[1] + stud_base_gap;
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt6), GetFamilySymbolByName(stud_face_block),
                        draftingView);
                    RotateFamily(family, 90);
                    family.LookupParameter("Distance1")
                        ?.Set(UnitUtils.ConvertToInternalUnits(WinPos - WinStudOff - stud_base_gap,
                            DisplayUnitType.DUT_DECIMAL_INCHES));
                    textNoteOptions = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Middle,
                        HorizontalAlignment = HorizontalTextAlignment.Center,
                        Rotation = Math.PI / 2,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "2.0 mm").Id
                    };
                    ptTemp[0] = pt6[0] + 1.5;
                    ptTemp[1] = pt6[1] + (WinPos - WinStudOff - stud_base_gap) / 2;
                    switch (stud_face_block)
                    {
                        case "VBA_2X4_FACE":
                            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
                            break;
                        case "VBA_LVL_FACE":
                            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "1.5 X 3.5 LVL",
                                textNoteOptions);
                            break;
                    }

                    pt6[1] = ptW[1] + WinPos + WinStudOff + WinGap;
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt6), GetFamilySymbolByName(stud_face_block),
                        draftingView);
                    RotateFamily(family, 90);
                    family.LookupParameter("Distance1")
                        ?.Set(UnitUtils.ConvertToInternalUnits(z - WinPos - WinStudOff,
                            DisplayUnitType.DUT_DECIMAL_INCHES));
                    textNoteOptions = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Middle,
                        HorizontalAlignment = HorizontalTextAlignment.Center,
                        Rotation = Math.PI / 2,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "2.0 mm").Id
                    };
                    ptTemp = new double[2];
                    ptTemp[0] = pt6[0] + 1.5;
                    ptTemp[1] = pt6[1] + (z - WinPos - WinStudOff) / 2;
                    switch (stud_face_block)
                    {
                        case "VBA_2X4_FACE":
                            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
                            break;
                        case "VBA_LVL_FACE":
                            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "1.5 X 3.5 LVL",
                                textNoteOptions);
                            break;
                    }
                }
            }

            pt4[1] = ptA[1];
            foreach (var t in ply_seams)
            {
                pt4[0] = ptA[0];
                pt4[1] += t;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt4), GetFamilySymbolByName("VBA_PLY_SHEET"),
                    draftingView);
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_DECIMAL_INCHES));
                family.LookupParameter("Distance2")
                    ?.Set(UnitUtils.ConvertToInternalUnits(t, DisplayUnitType.DUT_DECIMAL_INCHES));

                if (DrawB)
                {
                    pt4[0] = ptB[0];
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt4), GetFamilySymbolByName("VBA_PLY_SHEET"),
                        draftingView);
                    family.LookupParameter("Distance1")
                        ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));
                    family.LookupParameter("Distance2")
                        ?.Set(UnitUtils.ConvertToInternalUnits(t, DisplayUnitType.DUT_DECIMAL_INCHES));
                }
            }

            if (DrawW)
            {
                pt4[0] = ptW[0];
                pt4[1] = ptW[1];
                foreach (var t in ply_seams_win)
                {
                    pt4[1] += t;
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt4), GetFamilySymbolByName("VBA_PLY_SHEET"),
                        draftingView);
                    family.LookupParameter("Distance1")
                        ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_w, DisplayUnitType.DUT_DECIMAL_INCHES));
                    family.LookupParameter("Distance2")
                        ?.Set(UnitUtils.ConvertToInternalUnits(t, DisplayUnitType.DUT_DECIMAL_INCHES));

                    if (Math.Abs(pt4[1] - ptW[1] - WinPos) < 0.001)
                    {
                        pt4[1] += WinGap;
                    }
                }
            }

            var ply_tot = 0d;
            for (var i = 0; i < ply_seams.Length; i++)
            {
                ply_tot += ply_seams[i];
                foreach (var t in stud_spacing_x)
                {
                    pt29[0] = ptA[0] + ply_width_x + chamf_thk - t - 1.75;
                    pt29[1] = ptA[1] + ply_tot - 2;
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt29), GetFamilySymbolByName("VBA_SCREW_HEAD"),
                        draftingView);
                    if (i != ply_seams.Length - 1)
                    {
                        pt29[0] = ptA[0] + ply_width_x + chamf_thk - t - 1.75;
                        pt29[1] = ptA[1] + ply_tot + 2;
                        _doc.Create.NewFamilyInstance(GetXYZByPoint(pt29), GetFamilySymbolByName("VBA_SCREW_HEAD"),
                            draftingView);
                    }
                }
            }

            foreach (var t in stud_spacing_x)
            {
                pt29[0] = ptA[0] + ply_width_x + chamf_thk - t - 1.75;
                pt29[1] = ptA[1] + 2;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt29), GetFamilySymbolByName("VBA_SCREW_HEAD"),
                    draftingView);
            }

            if (DrawB == false)
            {
                pt_blk[0] = ptA[0] + ply_width_x + chamf_thk - stud_spacing_x[0] - 1.75;
                pt_blk[1] = ptA[1] + ply_seams[0] - 2;
                if (ply_seams[0] < 48 && (ply_seams.Length) > 1)
                {
                    pt_blk[1] += ply_seams[1];
                }

                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk),
                    GetFamilySymbolByName("VBA_ELEVATION_NOTES_SCREWS"),
                    draftingView);
                pt_blk[0] = ptA[0] + ply_width_x;
                pt_blk[1] = ptA[1];
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk),
                    GetFamilySymbolByName("VBA_ELEVATION_NOTES"),
                    draftingView);
            }

            if (DrawB)
            {
                ply_tot = 0;
                for (var i = 0; i < ply_seams.Length; i++)
                {
                    ply_tot += ply_seams[i];
                    foreach (var t in stud_spacing_y)
                    {
                        pt29[0] = ptB[0] + ply_width_y + chamf_thk - t - 1.75;
                        pt29[1] = ptB[1] + ply_tot - 2;
                        _doc.Create.NewFamilyInstance(GetXYZByPoint(pt29), GetFamilySymbolByName("VBA_SCREW_HEAD"),
                            draftingView);
                        if (i != ply_seams.Length - 1)
                        {
                            pt29[0] = ptB[0] + ply_width_y + chamf_thk - t - 1.75;
                            pt29[1] = ptB[1] + ply_tot + 2;
                            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt29), GetFamilySymbolByName("VBA_SCREW_HEAD"),
                                draftingView);
                        }
                    }
                }

                for (var i = 0; i < stud_spacing_y.Length; i++)
                {
                    pt29[0] = ptB[0] + ply_width_y + chamf_thk - stud_spacing_y[i] - 1.75;
                    pt29[1] = ptB[1] + 2;
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt29), GetFamilySymbolByName("VBA_SCREW_HEAD"),
                        draftingView);
                }

                if (DrawW == false)
                {
                    pt_blk[0] = ptB[0] + ply_width_y + chamf_thk - stud_spacing_y[0] - 1.75;
                    pt_blk[1] = ptB[1] + ply_seams[0] - 2;
                    if (ply_seams[0] < 48 && ply_seams.Length > 1)
                    {
                        pt_blk[1] = pt_blk[1] + ply_seams[1];
                    }

                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk),
                        GetFamilySymbolByName("VBA_ELEVATION_NOTES_SCREWS"),
                        draftingView);
                    pt_blk[0] = ptB[0] + ply_width_y;
                    pt_blk[1] = ptB[1];
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk), GetFamilySymbolByName("VBA_ELEVATION_NOTES"),
                        draftingView);
                }
                else if (DrawW)
                {
                    pt_blk[0] = ptB[0] + ply_width_y + chamf_thk - stud_spacing_y[stud_spacing_y.Length - 1] - 1.75;
                    pt_blk[1] = ptB[1] + ply_seams[0] - 2;
                    if (ply_seams[0] < 48 && ply_seams.Length > 1)
                    {
                        pt_blk[1] += ply_seams[1];
                    }

                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk),
                        GetFamilySymbolByName("VBA_ELEVATION_NOTES_SCREWS_MIRRORED"),
                        draftingView);
                    pt_blk[0] = ptB[0] + ply_width_y;
                    pt_blk[1] = ptB[1];
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk),
                        GetFamilySymbolByName("VBA_ELEVATION_NOTES_MIRRORED"),
                        draftingView);
                    family.LookupParameter("Distance1")
                        ?.Set(UnitUtils.ConvertToInternalUnits(y + 12, DisplayUnitType.DUT_DECIMAL_INCHES));
                }
            }

            if (DrawW)
            {
                ply_tot = 0;
                for (var i = 0; i < ply_seams_win.Length; i++)
                {
                    ply_tot += ply_seams_win[i];
                    for (var j = 0; j < stud_spacing_w.Length; j++)
                    {
                        pt29[0] = ptW[0] + ply_width_w + chamf_thk - stud_spacing_w[j] - 1.75;
                        pt29[1] = ptW[1] + ply_tot - 2;
                        if (Math.Abs(ply_tot - WinPos) < 0.001)
                        {
                            pt29[1] -= WinStudOff;
                        }

                        if (ply_tot > WinPos)
                        {
                            pt29[1] += WinGap;
                        }

                        _doc.Create.NewFamilyInstance(GetXYZByPoint(pt29), GetFamilySymbolByName("VBA_SCREW_HEAD"),
                            draftingView);
                        if (i != ply_seams_win.Length - 1)
                        {
                            pt29[0] = ptW[0] + ply_width_w + chamf_thk - stud_spacing_w[j] - 1.75;
                            pt29[1] = ptW[1] + ply_tot + 2;
                            if (Math.Abs(ply_tot - WinPos) < 0.001)
                            {
                                pt29[1] += WinStudOff + WinGap;
                            }

                            if (ply_tot > WinPos)
                            {
                                pt29[1] += WinGap;
                            }

                            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt29), GetFamilySymbolByName("VBA_SCREW_HEAD"),
                                draftingView);
                        }
                    }
                }

                foreach (var t in stud_spacing_w)
                {
                    pt29[0] = ptW[0] + ply_width_w + chamf_thk - t - 1.75;
                    pt29[1] = ptW[1] + 2;
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt29), GetFamilySymbolByName("VBA_SCREW_HEAD"),
                        draftingView);
                }
            }

            pt18[0] = ptA[0] + ply_width_x;
            pt18[1] = ptA[1] + stud_base_gap;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt18), GetFamilySymbolByName("VBA_CHAMF"),
                draftingView);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(z - stud_base_gap, DisplayUnitType.DUT_DECIMAL_INCHES));
            if (DrawB)
            {
                pt18[0] = ptB[0] + ply_width_y;
                pt18[1] = ptB[1] + stud_base_gap;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt18), GetFamilySymbolByName("VBA_CHAMF"),
                    draftingView);
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(z - stud_base_gap, DisplayUnitType.DUT_DECIMAL_INCHES));
            }

            if (DrawW)
            {
                pt18[0] = ptW[0] + ply_width_w;
                pt18[1] = ptW[1] + stud_base_gap;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt18), GetFamilySymbolByName("VBA_CHAMF"),
                    draftingView);
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(WinPos - stud_base_gap, DisplayUnitType.DUT_DECIMAL_INCHES));
                pt18[0] = ptW[0] + ply_width_w;
                pt18[1] = ptW[1] + WinPos + WinGap;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt18), GetFamilySymbolByName("VBA_CHAMF"),
                    draftingView);
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(z - WinPos, DisplayUnitType.DUT_DECIMAL_INCHES));
            }

            if (window == false)
            {
                for (var i = 0; i < n_studs_e; i++)
                {
                    pt20[0] = ptE[0] + stud_spacing_e[i];
                    pt20[1] = ptE[1] + stud_base_gap;
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName(stud_face_block),
                        draftingView);
                    RotateFamily(family, 90);
                    family.LookupParameter("Distance1")
                        ?.Set(UnitUtils.ConvertToInternalUnits(z - stud_base_gap, DisplayUnitType.DUT_DECIMAL_INCHES));
                    family.LookupParameter("Solid")?.Set(1);
                    textNoteOptions = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Middle,
                        HorizontalAlignment = HorizontalTextAlignment.Center,
                        Rotation = Math.PI / 2,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "2.0 mm").Id
                    };
                    ptTemp[0] = pt20[0] + 1.5;
                    ptTemp[1] = pt20[1] + (z - stud_base_gap) / 2;
                    switch (stud_face_block)
                    {
                        case "VBA_2X4_FACE":
                            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
                            break;
                        case "VBA_LVL_FACE":
                            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "1.5 X 3.5 LVL",
                                textNoteOptions);
                            break;
                    }
                }
            }
            else if (window)
            {
                if (DrawW)
                {
                    for (var i = 0; i < n_studs_e; i++)
                    {
                        pt20[0] = ptE[0] + stud_spacing_e[i];
                        pt20[1] = ptE[1] + stud_base_gap;
                        family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20),
                            GetFamilySymbolByName(stud_face_block),
                            draftingView);
                        RotateFamily(family, 90);
                        family.LookupParameter("Distance1")
                            ?.Set(UnitUtils.ConvertToInternalUnits(WinPos - stud_base_gap - WinStudOff,
                                DisplayUnitType.DUT_DECIMAL_INCHES));
                        family.LookupParameter("Solid")?.Set(1);
                        textNoteOptions = new TextNoteOptions
                        {
                            VerticalAlignment = VerticalTextAlignment.Middle,
                            HorizontalAlignment = HorizontalTextAlignment.Center,
                            Rotation = Math.PI / 2,
                            TypeId = new FilteredElementCollector(_doc)
                                .OfClass(typeof(TextNoteType))
                                .Cast<TextNoteType>()
                                .First(q => q.Name == "2.0 mm").Id
                        };
                        ptTemp[0] = pt20[0] + 1.5;
                        ptTemp[1] = pt20[1] + (WinPos - stud_base_gap - WinStudOff) / 2;
                        switch (stud_face_block)
                        {
                            case "VBA_2X4_FACE":
                                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
                                break;
                            case "VBA_LVL_FACE":
                                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "1.5 X 3.5 LVL",
                                    textNoteOptions);
                                break;
                        }

                        pt21[0] = ptE[0] + stud_spacing_e[i];
                        pt21[1] = ptE[1] + WinPos + WinStudOff;
                        family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt21),
                            GetFamilySymbolByName(stud_face_block),
                            draftingView);
                        RotateFamily(family, 90);
                        family.LookupParameter("Distance1")
                            ?.Set(UnitUtils.ConvertToInternalUnits(z - WinPos - WinStudOff,
                                DisplayUnitType.DUT_DECIMAL_INCHES));
                        family.LookupParameter("Solid")?.Set(1);
                        textNoteOptions = new TextNoteOptions
                        {
                            VerticalAlignment = VerticalTextAlignment.Middle,
                            HorizontalAlignment = HorizontalTextAlignment.Center,
                            Rotation = Math.PI / 2,
                            TypeId = new FilteredElementCollector(_doc)
                                .OfClass(typeof(TextNoteType))
                                .Cast<TextNoteType>()
                                .First(q => q.Name == "2.0 mm").Id
                        };
                        ptTemp = new double[2];
                        ptTemp[0] = pt21[0] + 1.5;
                        ptTemp[1] = pt21[1] + (z - WinPos - WinStudOff) / 2;
                        switch (stud_face_block)
                        {
                            case "VBA_2X4_FACE":
                                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
                                break;
                            case "VBA_LVL_FACE":
                                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "1.5 X 3.5 LVL",
                                    textNoteOptions);
                                break;
                        }
                    }
                }
            }

            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(new double[] {ptE[0] + chamf_thk, ptE[1]}),
                GetFamilySymbolByName("VBA_RECTANGLE"),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("xB")
                .Set(UnitUtils.ConvertToInternalUnits(ply_width_e, DisplayUnitType.DUT_DECIMAL_INCHES));
            family.LookupParameter("yB").Set(UnitUtils.ConvertToInternalUnits(z, DisplayUnitType.DUT_DECIMAL_INCHES));
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(ptE), GetFamilySymbolByName("VBA_PLY"),
                draftingView);
            RotateFamily(family, 90);
            family.LookupParameter("Distance1")
                .Set(UnitUtils.ConvertToInternalUnits(z, DisplayUnitType.DUT_DECIMAL_INCHES));

            if (window)
            {
                pt20[0] = ptE[0] + chamf_thk;
                pt20[1] = ptE[1] + WinPos;
                pt21[0] = ptE[0] + ply_width_e + chamf_thk;
                pt21[1] = ptE[1] + WinPos;
                _doc.Create.NewDetailCurve(draftingView, Line.CreateBound(GetXYZByPoint(pt20), GetXYZByPoint(pt21)));
            }

            if (window)
            {
                if (z - WinPos < 18)
                {
                    pt20[1] = ptE[1] + z;
                }
                else
                {
                    pt20[1] = ptE[1] + WinPos + 18;
                }

                pt21[1] = pt20[1] - 36;
                var Inserted2x2Note = false;
                for (var i = 0; i < stud_spacing_w.Length - 1; i++)
                {
                    if (stud_spacing_w[i + 1] - stud_spacing_w[i] > min_2x2_gap)
                    {
                        pt20[0] = ptE[0] + stud_spacing_w[i] + 3.5;
                        pt21[0] = pt20[0] + 1.5;
                        var pt20F = new[] {pt20[0], pt20[1] - 36};
                        family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20F),
                            GetFamilySymbolByName("VBA_RECTANGLE"),
                            draftingView);
                        _doc.Regenerate();
                        family.LookupParameter("xB")
                            .Set(UnitUtils.ConvertToInternalUnits(1.5, DisplayUnitType.DUT_DECIMAL_INCHES));
                        family.LookupParameter("yB")
                            .Set(UnitUtils.ConvertToInternalUnits(36, DisplayUnitType.DUT_DECIMAL_INCHES));
                        if (Inserted2x2Note == false)
                        {
                            Inserted2x2Note = true;
                            for (var j = 0; j < n_clamps - 2; j++)
                            {
                                if (clamp_spacing_con[i] < WinPos)
                                {
                                    pt18[0] = ptE[0] + stud_spacing_w[stud_spacing_w.Length - 2] + 3.5 + 1.5;
                                    pt18[1] = ptE[1] + clamp_spacing_con[j] - 6;
                                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20),
                                        GetFamilySymbolByName("VBA_2X2_NOTES"),
                                        draftingView);
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            pt21[0] = ptE[0] + chamf_thk + ply_width_e;
            pt21[1] = ptE[1] + z;
            CreateDimension(pt21[0] - clamp_L, pt21[1],
                pt21[0] - clamp_L, pt21[1] - clamp_spacing[0], 0, 0, draftingView, _doc);
            for (var i = 0; i < n_clamps; i++)
            {
                pt21[1] -= clamp_spacing[i];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt21), GetFamilySymbolByName(clamp_block_pr),
                    draftingView);
                _doc.Regenerate();
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(1.5 + ply_thk + chamf_thk + ply_width_e,
                        DisplayUnitType.DUT_DECIMAL_INCHES));
                if (i != 0)
                {
                    CreateDimension(pt21[0] - clamp_L, pt21[1] + clamp_spacing[i], pt21[0] - clamp_L, pt21[1],
                        0, 0, draftingView, _doc);
                }

                pt22[0] = pt21[0] - clamp_L + 3.75 - 12;
                pt22[1] = ptE[1] + clamp_spacing_con[i] + 1.5;
                var clamp_str = $"{clamp_spacing_con[i]}\"";
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Top,
                    HorizontalAlignment = HorizontalTextAlignment.Left,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "2.5 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt22), clamp_str, textNoteOptions);
                pt25[0] = ptE[0] + chamf_thk + ply_width_e + 9;
                pt25[1] = pt21[1] + 5.75 - 3 * (WinY == true ? 1 : 0);
                switch (clamp_size)
                {
                    case 1 when ply_width_e - 1.5 >= 19:
                    case 2 when ply_width_e - 1.5 >= 31:
                    case 3 when ply_width_e - 1.5 >= 43:
                        pt25[1] -= 3;
                        break;
                }

                if (i < n_top_clamps)
                {
                    textNoteOptions = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Top,
                        HorizontalAlignment = HorizontalTextAlignment.Left,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "2.5 mm").Id
                    };
                    TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt25), "TOP CLAMP", textNoteOptions);
                }
                else
                {
                    textNoteOptions = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Top,
                        HorizontalAlignment = HorizontalTextAlignment.Left,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "2.5 mm").Id
                    };
                    TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt25), "CLAMP", textNoteOptions);
                }
            }

            CreateDimension(pt21[0] - clamp_L, ptE[1], pt21[0] - clamp_L, ptE[1] + bot_clamp_gap,
                0, 0, draftingView, _doc);
            if (n_top_clamps >= 2 || _ui.Picking.IsChecked == true)
            {
                pt27[0] = ptE[0] + chamf_thk + 2;
                pt27[1] = ptE[1] + z - clamp_spacing[0] + 2.25 - 0.5 * (WinY ? 1 : 0);
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_COIL_ROD"),
                    draftingView);
                RotateFamily(family, 180);
                if (ply_width_w >= 37.5)
                {
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_COIL_ROD_NOTES"),
                        draftingView);
                }
                else if (_ui.Picking.IsChecked == true)
                {
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_COIL_ROD_NOTES_C"),
                        draftingView);
                }
                else
                {
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_COIL_ROD_NOTES_B"),
                        draftingView);
                }

                pt27[1] -= 2;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_NUT_5-8"),
                    draftingView);
                pt27[1] -= 0.25;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_NUT_5-8"),
                    draftingView);
                RotateFamily(family, 180);
                pt27[1] -= clamp_spacing[1];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_NUT_5-8"),
                    draftingView);
                RotateFamily(family, 180);
                pt27[1] += 0.25;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_NUT_5-8"),
                    draftingView);
            }

            if (window)
            {
                if (n_top_clamps >= 2 || _ui.Picking.IsChecked == true)
                {
                    if (clamp_spacing[1] > WinPos)
                    {
                        pt27[0] = ptE[0] + chamf_thk + ply_width_w / 2;
                        pt27[1] = ptE[1] + WinPos + win_clamp_top_max + 2.625 - 0.5 * (WinY ? 1 : 0);
                        family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27),
                            GetFamilySymbolByName("VBA_COIL_ROD_14"),
                            draftingView);
                        RotateFamily(family, 180);
                        _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27),
                            GetFamilySymbolByName("VBA_COIL_ROD_NOTES_14"),
                            draftingView);
                        pt27[1] -= 2.375;
                        _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_NUT_5-8"),
                            draftingView);
                        pt27[1] -= 0.25;
                        family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27),
                            GetFamilySymbolByName("VBA_NUT_5-8"),
                            draftingView);
                        RotateFamily(family, 180);
                        pt27[1] -= win_clamp_top_max + win_clamp_bot_max - 0.25;
                        _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_NUT_5-8"),
                            draftingView);
                        pt27[1] -= 0.25;
                        family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27),
                            GetFamilySymbolByName("VBA_NUT_5-8"),
                            draftingView);
                        RotateFamily(family, 180);
                    }
                }
                else
                {
                    pt27[0] = ptE[0] + chamf_thk + ply_width_w / 2;
                    pt27[1] = ptE[1] + WinPos + win_clamp_top_max + 2.625 - 0.5 * (WinY ? 1 : 0);
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27),
                        GetFamilySymbolByName("VBA_COIL_ROD_14"),
                        draftingView);
                    RotateFamily(family, 180);
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_COIL_ROD_NOTES_14"),
                        draftingView);
                    pt27[1] -= 2.375;
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_NUT_5-8"),
                        draftingView);
                    pt27[1] -= 0.25;
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_NUT_5-8"),
                        draftingView);
                    RotateFamily(family, 180);
                    pt27[1] -= win_clamp_top_max + win_clamp_bot_max - 0.25;
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_NUT_5-8"),
                        draftingView);
                    pt27[1] -= 0.25;
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt27), GetFamilySymbolByName("VBA_NUT_5-8"),
                        draftingView);
                    RotateFamily(family, 180);
                }
            }

            if (x > 40)
            {
                for (int i = 0; i < clamp_spacing_con.Length - 1; i++)
                {
                    if (z - clamp_spacing_con[i] >= 87)
                    {
                        if (WinY)
                        {
                            pt31[0] = ptE[0] + chamf_thk + (ply_width_e - 36) / 2 + 36;
                            pt31[1] = ptE[1] + clamp_spacing_con[i] - 0.25;
                            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt31),
                                GetFamilySymbolByName("VBA_REINFORCING_ANGLE"),
                                draftingView);
                            RotateFamily(family, 180);
                        }
                        else
                        {
                            pt31[0] = ptE[0] + chamf_thk + (ply_width_e - 36) / 2;
                            pt31[1] = ptE[1] + clamp_spacing_con[i];
                            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt31),
                                GetFamilySymbolByName("VBA_REINFORCING_ANGLE"),
                                draftingView);
                        }
                    }
                }
            }

            if (_ui.Regular.IsChecked == true)
            {
                pt26[0] = ptE[0] - ply_thk - 1.25;
                pt26[1] = ptE[1] + z - clamp_spacing[0];
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt26), GetFamilySymbolByName("VBA_LIFTING_SLING_A"),
                    draftingView);
                pt26[0] = ptE[0] + chamf_thk + ply_width_e + 1.3;
                pt26[1] = ptE[1] + z - clamp_spacing[0] + 3;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt26), GetFamilySymbolByName("VBA_LIFTING_SLING_B"),
                    draftingView);
            }
            else if (_ui.Picking.IsChecked == true)
            {
                pt26[0] = ptE[0] - ply_thk - 1.5;
                pt26[1] = ptE[1] + z - clamp_spacing[0] - 0.25;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt26), GetFamilySymbolByName("VBA_LIFTING_CORNER_A"),
                    draftingView);
                pt26[0] = ptE[0] + chamf_thk + ply_width_e;
                pt26[1] = ptE[1] + z - clamp_spacing[0] - 0.25;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt26), GetFamilySymbolByName("VBA_LIFTING_CORNER_B"),
                    draftingView);
            }

            pt28[0] = ptE[0] + chamf_thk + ply_width_e + 4;
            pt28[1] = ptE[1] + clamp_spacing_con[brace_clamp];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt28), GetFamilySymbolByName("VBA_AFB_SIDE"),
                draftingView);
            pt28[0] = pt28[0] + 1.9375;
            pt28[1] = pt28[1] + 0.6789;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt28), GetFamilySymbolByName("VBA_BRACE_ANGLE"),
                draftingView);
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt28), GetFamilySymbolByName(brace_block),
                draftingView);
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt28), GetFamilySymbolByName("VBA_BRACE_NOTES"),
                draftingView);
            pt28[0] = pt28[0] - 5.9375;
            pt28[1] = ptE[1] + clamp_spacing_con[chain_clamp];
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt28), GetFamilySymbolByName("VBA_CHAIN"),
                draftingView);

            if (pt28[1] - pt_o[1] < 62)
            {
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(6.27 + 62 - (pt28[1] - pt_o[1]),
                        DisplayUnitType.DUT_DECIMAL_INCHES));
            }

            pt28[0] = ptE[0] + 4.68;
            pt28[1] = ptE[1] + clamp_spacing_con[brace_clamp];
            if (window == false)
            {
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt28), GetFamilySymbolByName("VBA_AFB_FACE"),
                    draftingView);
            }

            pt20[0] = ptE[0] - ply_thk - 1.5 - 3.5;
            pt20[1] = ptE[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName("VBA_2X4"),
                draftingView);
            pt20[0] = ptE[0] + chamf_thk + ply_width_e;
            pt20[1] = ptE[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName("VBA_2X4"),
                draftingView);
            pt20[0] = ptE[0] - ply_thk - 1.5 - 6;
            pt20[1] = ptE[1] + 3;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName("VBA_2X4_SIDE"),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_e + chamf_thk + ply_thk + 1.5 + 12,
                    DisplayUnitType.DUT_DECIMAL_INCHES));
            textNoteOptions = new TextNoteOptions
            {
                VerticalAlignment = VerticalTextAlignment.Middle,
                HorizontalAlignment = HorizontalTextAlignment.Center,
                TypeId = new FilteredElementCollector(_doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .First(q => q.Name == "1.2 mm").Id
            };

            ptTemp[0] = pt20[0] + (ply_width_e + chamf_thk + ply_thk + 1.5 + 12) / 2;
            ptTemp[1] = pt20[1] - 0.7;
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);

            pt30[0] = ptE[0] + chamf_thk + ply_width_e + 6;
            pt30[1] = ptE[1] + 2.25;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30), GetFamilySymbolByName("VBA_DOWN_PLATE_NOTES"),
                draftingView);
            pt8[0] = ptA[0] + ply_width_x;
            pt8[1] = ptA[1] + z + 18 + (window ? 1 : 0) * WinGap;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt8), GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                draftingView);
            RotateFamily(family, 180);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_DECIMAL_INCHES));

            for (var j = 0; j < n_studs_x; j++)
            {
                pt11[0] = ptA[0] + ply_width_x + chamf_thk - 3.5 - stud_spacing_x[j];
                pt11[1] = ptA[1] + z + 18 + (window ? 1 : 0) * WinGap;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt11), GetFamilySymbolByName(stud_block),
                    draftingView);
            }

            if (DrawB)
            {
                pt9[0] = ptB[0] + ply_width_y;
                pt9[1] = ptB[1] + z + 18 + (window ? 1 : 0) * WinGap;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt9),
                    GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                    draftingView);
                RotateFamily(family, 180);
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));

                for (int j = 0; j < n_studs_y; j++)
                {
                    pt11[0] = ptB[0] + ply_width_y + chamf_thk - 3.5 - stud_spacing_y[j];
                    pt11[1] = ptB[1] + z + 18 + (window ? 1 : 0) * WinGap;
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt11), GetFamilySymbolByName(stud_block),
                        draftingView);
                }
            }

            if (DrawW)
            {
                pt10[0] = ptW[0] + ply_width_w;
                pt10[1] = ptW[1] + z + 18 + (window ? 1 : 0) * WinGap;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt10),
                    GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                    draftingView);
                RotateFamily(family, 180);
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_w, DisplayUnitType.DUT_DECIMAL_INCHES));
                for (int j = 0; j < n_studs_w; j++)
                {
                    pt11[0] = ptW[0] + ply_width_w + chamf_thk - 3.5 - stud_spacing_w[j];
                    pt11[1] = ptW[1] + z + 18 + (window ? 1 : 0) * WinGap;
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt11), GetFamilySymbolByName(stud_block),
                        draftingView);
                }
            }

            CreateDimension(pt8[0], pt8[1] - ply_thk, pt8[0] - ply_width_x, pt8[1] - ply_thk,
                0, -5, draftingView, _doc);
            CreateDimension(pt8[0] + chamf_thk, pt8[1] - ply_thk,
                pt8[0] - ply_width_x, pt8[1] - ply_thk, 0, -10, draftingView, _doc);
            for (int i = 0; i < n_studs_x; i++)
            {
                CreateDimension(pt8[0] + chamf_thk - stud_spacing_x[stud_spacing_x.Length - 1 - i], pt8[1],
                    pt8[0] - ply_width_x, pt8[1], 0, (i + 2) * 5, draftingView, _doc);
            }

            CreateDimension(pt8[0] - stud_spacing_x[stud_spacing_x.Length - 1] - 3.5 + chamf_thk, pt8[1],
                pt8[0] - ply_width_x, pt8[1], 0, 5, draftingView, _doc);

            CreateDimension(pt8[0] + chamf_thk - stud_start_offset, pt8[1],
                pt8[0] + chamf_thk, pt8[1], 0, (n_studs_x + 1) * 5, draftingView, _doc);

            if (DrawB)
            {
                CreateDimension(pt9[0], pt9[1] - ply_thk,
                    pt9[0] - ply_width_y, pt9[1] - ply_thk, 0, -5, draftingView, _doc);
                CreateDimension(pt9[0] + chamf_thk, pt9[1] - ply_thk,
                    pt9[0] - ply_width_y, pt9[1] - ply_thk, 0, -10, draftingView, _doc);
                for (int i = 0; i < n_studs_y; i++)
                {
                    CreateDimension(pt9[0] + chamf_thk - stud_spacing_y[stud_spacing_y.Length - 1 - i],
                        pt9[1],
                        pt9[0] - ply_width_y, pt9[1], 0, (i + 2) * 5, draftingView, _doc);
                }

                CreateDimension(pt9[0] - stud_spacing_y[stud_spacing_y.Length - 1] - 3.5 + chamf_thk, pt9[1],
                    pt9[0] - ply_width_y, pt9[1], 0, 5, draftingView, _doc);
                CreateDimension(pt9[0] + chamf_thk - stud_start_offset, pt9[1],
                    pt9[0] + chamf_thk, pt9[1], 0, (n_studs_y + 1) * 5, draftingView, _doc);
            }

            if (DrawW)
            {
                CreateDimension(pt10[0], pt10[1] - ply_thk,
                    pt10[0] - ply_width_w, pt10[1] - ply_thk, 0, -5, draftingView, _doc);
                CreateDimension(pt10[0] + chamf_thk, pt10[1] - ply_thk,
                    pt10[0] - ply_width_w, pt10[1] - ply_thk, 0, -10, draftingView, _doc);
                for (var i = 0; i < n_studs_w; i++)
                {
                    CreateDimension(pt10[0] + chamf_thk - stud_spacing_w[stud_spacing_w.Length - 1 - i], pt10[1],
                        pt10[0] - ply_width_w, pt10[1], 0, (i + 2) * 5, draftingView, _doc);
                }

                CreateDimension(pt10[0] - stud_spacing_w[stud_spacing_w.Length - 1] - 3.5 + chamf_thk,
                    pt10[1], pt10[0] - ply_width_w, pt10[1], 0, 5, draftingView, _doc);
                CreateDimension(pt10[0] + chamf_thk - stud_start_offset, pt10[1],
                    pt10[0] + chamf_thk, pt10[1], 0, (n_studs_w + 1) * 5, draftingView, _doc);
            }

            var PlyTemp = ptA[1] + ply_seams[0];
            var dimLine = CreateDimension(ptA[0] + ply_width_x, ptA[1],
                ptA[0] + ply_width_x, PlyTemp, 7, 0, draftingView, _doc);
            dimLine.Suffix = "PLYWOOD";
            if (ply_seams.Length >= 2)
            {
                for (int i = 1; i < ply_seams.Length; i++)
                {
                    PlyTemp += ply_seams[i];
                    dimLine = CreateDimension(ptA[0] + ply_width_x, PlyTemp - ply_seams[i],
                        ptA[0] + ply_width_x, PlyTemp, 7, 0, draftingView, _doc);
                    dimLine.Suffix = "PLYWOOD";
                }
            }

            dimLine = CreateDimension(ptA[0], ptA[1], ptA[0], ptA[1] + z, -7, 0, draftingView, _doc);
            dimLine.Suffix = "OVERALL HEIGHT";
            dimLine = CreateDimension(ptA[0], ptA[1] + z, ptA[0], ptA[1] + stud_base_gap, -2, 0, draftingView, _doc);
            dimLine.Suffix = "STUD";
            dimLine = CreateDimension(ptA[0] + ply_width_x, ptA[1] + z, ptA[0], ptA[1] + z, 0, 2, draftingView, _doc);
            dimLine.Suffix = "PLYWOOD";
            if (DrawB)
            {
                PlyTemp = ptB[1] + ply_seams[0];
                dimLine = CreateDimension(ptB[0] + ply_width_y, ptB[1], ptB[0] + ply_width_y, PlyTemp,
                    7, 0, draftingView, _doc);
                dimLine.Suffix = "PLYWOOD";
                if (ply_seams.Length >= 2)
                {
                    for (var i = 1; i < ply_seams.Length; i++)
                    {
                        PlyTemp += ply_seams[i];
                        dimLine = CreateDimension(ptB[0] + ply_width_y, PlyTemp - ply_seams[i],
                            ptB[0] + ply_width_y, PlyTemp, 7, 0, draftingView, _doc);
                        dimLine.Suffix = "PLYWOOD";
                    }
                }

                dimLine = CreateDimension(ptB[0], ptB[1], ptB[0], ptB[1] + z, -7, 0, draftingView, _doc);
                dimLine.Suffix = "OVERALL HEIGHT";
                dimLine = CreateDimension(ptB[0], ptB[1] + z, ptB[0], ptB[1] + stud_base_gap, -2, 0, draftingView,
                    _doc);
                dimLine.Suffix = "STUD";
                dimLine = CreateDimension(ptB[0] + ply_width_y, ptB[1] + z, ptB[0], ptB[1] + z, 0, 2, draftingView,
                    _doc);
                dimLine.Suffix = "PLYWOOD";
            }

            if (DrawW)
            {
                PlyTemp = ptW[1] + ply_seams_win[0];
                dimLine = CreateDimension(ptW[0] + ply_width_w, ptW[1], ptW[0] + ply_width_w, PlyTemp, 7, 0,
                    draftingView, _doc);
                dimLine.Suffix = "PLYWOOD";
                if (ply_seams_win.Length >= 2)
                {
                    for (int i = 1; i < ply_seams_win.Length; i++)
                    {
                        if (Math.Abs(PlyTemp - ptW[1] - WinPos) < 0.001)
                        {
                            PlyTemp += WinGap;
                        }

                        PlyTemp += ply_seams_win[i];
                        if (ply_seams_win[i] < 26)
                        {
                            dimLine = CreateDimension(ptW[0] + ply_width_w, PlyTemp - ply_seams_win[i],
                                ptW[0] + ply_width_w, PlyTemp, 7, 0, draftingView, _doc);
                            dimLine.Suffix = "PLYWOOD";
                        }
                        else
                        {
                            dimLine = CreateDimension(ptW[0] + ply_width_w, PlyTemp - ply_seams_win[i],
                                ptW[0] + ply_width_w, PlyTemp, 7, 0, draftingView, _doc);
                            dimLine.Suffix = "PLYWOOD";
                        }
                    }
                }

                dimLine = CreateDimension(ptW[0], ptW[1] + WinPos - WinStudOff,
                    ptW[0], ptW[1] + stud_base_gap, -7, 0, draftingView, _doc);
                dimLine.Suffix = "STUD";
                dimLine = CreateDimension(ptW[0], ptW[1] + WinPos + WinStudOff + WinGap,
                    ptW[0], ptW[1] + z + WinGap, -7, 0, draftingView, _doc);
                dimLine.Suffix = "STUD";
                CreateDimension(ptW[0], ptW[1] + WinPos - WinStudOff,
                    ptW[0], ptW[1] + WinPos, -7, 0, draftingView, _doc);
                CreateDimension(ptW[0], ptW[1] + WinPos + WinStudOff + WinGap,
                    ptW[0], ptW[1] + WinPos + WinGap, -7, 0, draftingView, _doc);
                dimLine = CreateDimension(ptW[0] + ply_width_w, ptW[1] + z + WinGap,
                    ptW[0], ptW[1] + z + WinGap, 0, 2, draftingView, _doc);
                dimLine.Suffix = "PLYWOOD";
            }

            if (window)
            {
                CreateDimension(ptE[0] + chamf_thk, ptE[1] + WinPos - 7,
                    ptE[0] + chamf_thk, ptE[1] + WinPos + win_clamp_top_max - 7, -5, 0, draftingView, _doc);
                CreateDimension(ptE[0] + chamf_thk, ptE[1] + WinPos - 7,
                    ptE[0] + chamf_thk, ptE[1] + WinPos - win_clamp_bot_max - 7, -5, 0, draftingView, _doc);
            }

            pt13[0] = pt_o[0] + 427 + x - chamf_thk;
            pt13[1] = pt_o[1] + 42 + ply_thk;
            pt30[0] = pt13[0] - x;
            pt30[1] = pt13[1] + y;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30), GetFamilySymbolByName("VBA_PLAN_NOTES1"),
                draftingView);
            pt30[0] = pt13[0];
            pt30[1] = pt13[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30), GetFamilySymbolByName("VBA_PLAN_NOTES2"),
                draftingView);
            pt30[0] = pt13[0];
            pt30[1] = pt13[1] + y;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30), GetFamilySymbolByName("VBA_PLAN_NOTES3"),
                draftingView);
            pt30[0] = pt13[0] - x;
            pt30[1] = pt13[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30), GetFamilySymbolByName("VBA_PLAN_NOTES4"),
                draftingView);
            pt30[0] = pt13[0] - x;
            pt30[1] = pt13[1] + y + 83.5;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30), GetFamilySymbolByName("VBA_PLAN_NOTES5"),
                draftingView);
            pt30[0] = pt13[0] - (x - 12) / 2;
            pt30[1] = pt13[1] + y + 83.5;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30), GetFamilySymbolByName("VBA_PLAN_NOTES6"),
                draftingView);
            pt30[0] = pt13[0] - x;
            pt30[1] = pt13[1] + 83.5;

            if (_ui.Picking.IsChecked == true)
            {
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30), GetFamilySymbolByName("VBA_PLAN_NOTES7_LOOP"),
                    draftingView);
            }
            else
            {
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30), GetFamilySymbolByName("VBA_PLAN_NOTES7"),
                    draftingView);
            }

            pt15[0] = pt13[0];
            pt15[1] = pt13[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15), GetFamilySymbolByName(clamp_block_op), draftingView);
            pt15[0] = pt13[0] - x;
            pt15[1] = pt13[1] + y;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15), GetFamilySymbolByName(clamp_block_bk), draftingView);
            pt15[0] = pt13[0];
            pt15[1] = pt13[1] + 83.5;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15), GetFamilySymbolByName(clamp_block_op),
                draftingView);
            if (window)
            {
                family.LookupParameter("Angle1")?.Set(swing_ang);
            }

            pt15[0] = pt13[0] - x;
            pt15[1] = pt13[1] + y + 83.5;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15), GetFamilySymbolByName(clamp_block_bk), draftingView);
            pt15[0] = pt13[0] - x;
            pt15[1] = pt13[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15), GetFamilySymbolByName(SQ_NAME), draftingView);
            pt15[0] = pt13[0];
            pt15[1] = pt13[1] + y;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15), GetFamilySymbolByName(SQ_NAME), draftingView);
            RotateFamily(family, 180);
            if (_ui.Regular.IsChecked == true)
            {
                pt15[0] = pt13[0] - x;
                pt15[1] = pt13[1] + 83.5;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15), GetFamilySymbolByName(SQ_NAME_SLINGS),
                    draftingView);
                pt15[0] = pt13[0];
                pt15[1] = pt13[1] + y + 83.5;
                if (window == false)
                {
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15),
                        GetFamilySymbolByName(SQ_NAME_SLINGS),
                        draftingView);
                    RotateFamily(family, 180);
                }

                if (x < 14 || y < 14 || window)
                {
                    pt15[0] = pt13[0];
                    pt15[1] = pt13[1] + 83.5 + y;
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15),
                        GetFamilySymbolByName("VBA_LIFTING_SLING_C"),
                        draftingView);
                    pt15[0] = pt13[0] - x;
                    pt15[1] = pt13[1] + 83.5;
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15),
                        GetFamilySymbolByName("VBA_LIFTING_SLING_C"),
                        draftingView);
                    RotateFamily(family, 180);
                }
            }
            else if (_ui.Picking.IsChecked == true)
            {
                pt15[0] = pt13[0] - x;
                pt15[1] = pt13[1] + 83.5;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15),
                    GetFamilySymbolByName("VBA_PICKING_CORNER"),
                    draftingView);
                pt15[0] = pt13[0];
                pt15[1] = pt13[1] + y + 83.5;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15),
                    GetFamilySymbolByName("VBA_PICKING_CORNER"),
                    draftingView);
                RotateFamily(family, 180);
            }

            if (x >= 14 && y >= 14)
            {
                pt15[0] = pt13[0];
                pt15[1] = pt13[1] + y + 83.5;
            }
            else
            {
                pt15[0] = pt13[0] + 6;
                pt15[1] = pt13[1] - 1 + 83.5;
            }

            if (window == false)
            {
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15), GetFamilySymbolByName("VBA_PLAN_NOTES8"),
                    draftingView);
            }

            pt12[0] = pt13[0] - x + chamf_thk;
            pt12[1] = pt13[1] - ply_thk;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt12[0] = pt12[0] - ply_thk - chamf_thk;
            pt12[1] = pt12[1] + ply_width_y + ply_thk - chamf_thk - 1.5;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                draftingView);
            RotateFamily(family, -90);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));

            pt12[0] = pt12[0] + ply_thk - chamf_thk + ply_width_x - 1.5;
            pt12[1] = pt12[1] + ply_thk + chamf_thk;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                draftingView);
            RotateFamily(family, 180);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_DECIMAL_INCHES));

            pt12[0] = pt12[0] + ply_thk + chamf_thk;
            pt12[1] = pt12[1] - ply_thk + chamf_thk - ply_width_y + 1.5;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                draftingView);
            RotateFamily(family, 90);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));

            pt12[0] = pt12[0] - ply_width_x + 1.5;
            pt12[1] = pt12[1] + 82;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt12[0] = pt12[0] - ply_thk - chamf_thk;
            pt12[1] = pt12[1] + ply_width_y - 1.5;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                draftingView);
            RotateFamily(family, -90);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt12[0] = pt12[0] + ply_width_x - 1.5;
            pt12[1] = pt12[1] + ply_thk + chamf_thk;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                draftingView);
            RotateFamily(family, 180);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt12[0] = pt12[0] + ply_thk + chamf_thk;
            pt12[1] = pt12[1] - ply_width_y + 1.5;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                draftingView);
            RotateFamily(family, 90);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));

            pt14[1] = pt13[1] - ply_thk - 1.5;
            for (int j = 0; j < n_studs_x; j++)
            {
                pt14[0] = pt13[0] - x + stud_spacing_x[j];
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block_spax),
                    draftingView);
            }

            pt14[0] = pt13[0] - x + stud_spacing_x[stud_spacing_x.Length - 1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block_bolt),
                draftingView);
            pt14[0] = pt13[0] - x - ply_thk - 1.5;
            for (int j = 0; j < n_studs_y; j++)
            {
                pt14[1] = pt13[1] + y - stud_spacing_y[j];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block_spax),
                    draftingView);
                RotateFamily(family, -90);
            }

            pt14[1] = pt13[1] + y + ply_thk + 1.5;
            for (int j = 0; j < n_studs_x; j++)
            {
                pt14[0] = pt13[0] - stud_spacing_x[j];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block_spax),
                    draftingView);
                RotateFamily(family, 180);
            }

            pt14[0] = pt13[0] + ply_thk + 1.5;
            for (int j = 0; j < n_studs_y; j++)
            {
                pt14[1] = pt13[1] + stud_spacing_y[j];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block_spax),
                    draftingView);
                RotateFamily(family, 90);
            }

            pt14[1] = pt13[1] + stud_spacing_x[1];
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block_bolt),
                draftingView);
            RotateFamily(family, 90);
            pt13[1] = pt13[1] + 83.5;
            pt14[1] = pt13[1] - ply_thk - 1.5;
            for (int j = 0; j < n_studs_x; j++)
            {
                pt14[0] = pt13[0] - x + stud_spacing_x[j];
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block_bolt),
                    draftingView);
            }

            pt14[0] = pt13[0] - x - ply_thk - 1.5;
            for (int j = 0; j < n_studs_y; j++)
            {
                pt14[1] = pt13[1] + ply_width_y - stud_spacing_y[j] - 1.5;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block_bolt),
                    draftingView);
                RotateFamily(family, -90);
            }

            pt14[1] = pt13[1] + y + ply_thk + 1.5;
            for (int j = 0; j < n_studs_x; j++)
            {
                pt14[0] = pt13[0] - stud_spacing_x[j];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block_bolt),
                    draftingView);
                RotateFamily(family, 180);
            }

            if (window)
            {
                stud_block_bolt = stud_block_bolt_hidden;
            }

            pt14[0] = pt13[0] + ply_thk + 1.5;
            for (int j = 0; j < n_studs_y; j++)
            {
                pt14[1] = pt13[1] + stud_spacing_y[j];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block_bolt),
                    draftingView);
                RotateFamily(family, 90);
            }

            stud_block_bolt = stud_block_bolt.Replace("_HIDDEN", "");
            if (window)
            {
                for (int j = 0; j < n_studs_y; j++)
                {
                    pt5[0] = pt13[0] + ply_thk + 1.5 + 1 + Math.Sin(swing_ang) +
                             (3.25 + y - stud_spacing_y[j]) * Math.Cos(swing_ang);
                    pt5[1] = pt13[1] + y + ply_thk + 2.5 - Math.Cos(swing_ang) +
                             (3.25 + y - stud_spacing_y[j]) * Math.Sin(swing_ang);
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt5), GetFamilySymbolByName(stud_block_bolt),
                        draftingView);
                    RotateFamily(family, (swing_ang + Math.PI) * 180 / Math.PI);
                    pt5[0] = pt13[0] + ply_thk + 1.5 + 1 + 2.5 * Math.Sin(swing_ang) +
                             (3.25 + y - chamf_thk) * Math.Cos(swing_ang);
                    pt5[1] = pt13[1] + y + ply_thk + 2.5 - 2.5 * Math.Cos(swing_ang) +
                             (3.25 + y - chamf_thk) * Math.Sin(swing_ang);
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt5),
                        GetFamilySymbolByName("VBA_PLY_WITH_CHAMFER"),
                        draftingView);
                    RotateFamily(family, (swing_ang + Math.PI) * 180 / Math.PI);
                    family.LookupParameter("Distance1")
                        ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));
                }
            }

            if (brace_clamp > n_top_clamps)
            {
                pt13[1] = pt13[1] - 83.5;
            }

            if (clamp_L - 12.5 - x >= 8.5)
            {
                pt16[0] = pt13[0] - x - 6.25 + clamp_L - 4.25;
                pt16[1] = pt13[1] + ply_thk + y;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt16), GetFamilySymbolByName("VBA_AFB_PLAN"),
                    draftingView);
            }
            else
            {
                if (stud_spacing_x[1] - stud_spacing_x[0] - 3.5 >= min_stud_gap)
                {
                    pt16[0] = pt13[0] - stud_spacing_x[0] - 3.5 - (stud_spacing_x[1] - stud_spacing_x[0] - 3.5) / 2;
                    pt16[1] = pt13[1] + ply_thk + y;
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt16), GetFamilySymbolByName("VBA_AFB_PLAN"),
                        draftingView);
                }
                else
                {
                    pt16[0] = pt13[0] - x - 6.25 + clamp_L - 2;
                    pt16[1] = pt13[1] + ply_thk + y;
                    _doc.Create.NewFamilyInstance(GetXYZByPoint(pt16), GetFamilySymbolByName("VBA_AFB_PLAN"),
                        draftingView);
                }
            }

            pt16[0] = pt13[0] - x + 3.8125;
            pt16[1] = pt13[1] + ply_thk + y;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt16), GetFamilySymbolByName("VBA_AFB_PLAN"),
                draftingView);

            pt16[0] = pt13[0] - x - ply_thk;
            pt16[1] = pt13[1] + 3.8125;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt16), GetFamilySymbolByName("VBA_AFB_PLAN"),
                draftingView);
            RotateFamily(family, 90);
            pt16[0] = pt13[0] - x;
            pt16[1] = pt13[1];
            if (y > 27)
            {
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt16), GetFamilySymbolByName("VBA_PLAN_NOTES9A"),
                    draftingView);
            }
            else
            {
                pt16[1] = pt13[1] + y;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt16), GetFamilySymbolByName("VBA_PLAN_NOTES9B"),
                    draftingView);
            }

            if (brace_clamp <= n_top_clamps)
            {
                pt13[1] = pt13[1] - 83.5;
            }

            var pl_pts = new double[10];
            pl_pts[0] = pt13[0];
            pl_pts[1] = pt13[1];
            pl_pts[2] = pt13[0] - x;
            pl_pts[3] = pt13[1];
            pl_pts[4] = pt13[0] - x;
            pl_pts[5] = pt13[1] + y;
            pl_pts[6] = pt13[0];
            pl_pts[7] = pt13[1] + y;
            pl_pts[8] = pt13[0];
            pl_pts[9] = pt13[1];

            var pl_pts1 = new double[2];
            pl_pts1[0] = pl_pts[0];
            pl_pts1[1] = pl_pts[1];
            var pl_pts2 = new double[2];
            pl_pts2[0] = pl_pts[2];
            pl_pts2[1] = pl_pts[3];
            var pl_pts3 = new double[2];
            pl_pts3[0] = pl_pts[4];
            pl_pts3[1] = pl_pts[5];
            var pl_pts4 = new double[2];
            pl_pts4[0] = pl_pts[6];
            pl_pts4[1] = pl_pts[7];
            var pl_pts5 = new double[2];
            pl_pts5[0] = pl_pts[8];
            pl_pts5[1] = pl_pts[9];

            CreateFilledRegion(draftingView, pl_pts1, pl_pts2, pl_pts3, pl_pts4, pl_pts5);

            pl_pts[1] = pl_pts[1] + 83.5;
            pl_pts[3] = pl_pts[3] + 83.5;
            pl_pts[5] = pl_pts[5] + 83.5;
            pl_pts[7] = pl_pts[7] + 83.5;
            pl_pts[9] = pl_pts[9] + 83.5;

            pl_pts1[1] = pl_pts[1];
            pl_pts2[1] = pl_pts[3];
            pl_pts3[1] = pl_pts[5];
            pl_pts4[1] = pl_pts[7];
            pl_pts5[1] = pl_pts[9];

            CreateFilledRegion(draftingView, pl_pts1, pl_pts2, pl_pts3, pl_pts4, pl_pts5);
            if (x <= 16 || y <= 16)
            {
                pt17[0] = pt13[0] - x;
                pt17[1] = pt13[1] + 0.75 + y / 2;
                ptTemp[0] = pt17[0] + x / 2;
                ptTemp[1] = pt17[1];
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Top,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "1.5 mm Bold").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "CLAMPS", textNoteOptions);

                pt17[0] = pt13[0] - x;
                pt17[1] = pt13[1] + 2.1 + y / 2 + 83.5;
                ptTemp[0] = pt17[0] + x / 2;
                ptTemp[1] = pt17[1];
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Top,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "1.5 mm Bold").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "Top\nCLAMPS", textNoteOptions);
            }
            else
            {
                pt17[0] = pt13[0] - x;
                pt17[1] = pt13[1] + 1.25 + y / 2;
                ptTemp[0] = pt17[0] + x / 2;
                ptTemp[1] = pt17[1];
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Top,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "2.5 mm Bold").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "CLAMPS", textNoteOptions);

                pt17[0] = pt13[0] - x;
                pt17[1] = pt13[1] + 3.35 + y / 2 + 83.5;
                ptTemp[0] = pt17[0] + x / 2;
                ptTemp[1] = pt17[1];
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Top,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "2.5 mm Bold").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "TOP\nCLAMPS", textNoteOptions);
            }

            for (var i = 0; i < 2; i++)
            {
                pt17[0] = pt13[0] - x;
                pt17[1] = pt13[1] + y - 0.5 + 83.5 * i;
                ptTemp[0] = pt17[0] + x / 2;
                ptTemp[1] = pt17[1];
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Top,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "1.75 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "SIDE A", textNoteOptions);
                if (DrawB)
                {
                    pt17[0] = pt13[0] - x + 0.5;
                    pt17[1] = pt13[1] + 83.5 * i;
                    ptTemp[0] = pt17[0];
                    ptTemp[1] = pt17[1] + y / 2;
                    textNoteOptions = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Top,
                        HorizontalAlignment = HorizontalTextAlignment.Center,
                        Rotation = Math.PI / 2,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "1.75 mm").Id
                    };
                    TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "SIDE B", textNoteOptions);
                }

                pt17[0] = pt13[0] - x;
                pt17[1] = pt13[1] + 2.25 + 83.5 * i;
                ptTemp[0] = pt17[0] + x / 2;
                ptTemp[1] = pt17[1];
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Top,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "1.75 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), $"{x}", textNoteOptions);
                pt17[0] = pt13[0] - 2.25;
                pt17[1] = pt13[1] + 83.5 * i;
                ptTemp[0] = pt17[0];
                ptTemp[1] = pt17[1] + y / 2;
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Top,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    Rotation = Math.PI / 2,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "1.75 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), $"{y}", textNoteOptions);
            }


            if (z > 190)
            {
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_o), GetFamilySymbolByName("VBA_COLUMN_BACKGROUND_HIGH"),
                    draftingView);
                var imageX = UnitUtils.ConvertToInternalUnits(385, DisplayUnitType.DUT_DECIMAL_INCHES);
                var imageY = UnitUtils.ConvertToInternalUnits(360, DisplayUnitType.DUT_DECIMAL_INCHES);
                var imageOption = new ImageImportOptions
                {
                    RefPoint = new XYZ(imageX, imageY, 0),
                    Resolution = 50,
                    Placement = BoxPlacement.Center
                };
                _doc.Import($"{GlobalNames.FilesLocationPrefix}POWERLAG SCREW.jpg", imageOption, draftingView,
                    out var image);
                image.get_Parameter(BuiltInParameter.RASTER_SHEETWIDTH)
                    .Set(UnitUtils.ConvertToInternalUnits(45, DisplayUnitType.DUT_DECIMAL_INCHES));

                imageX = UnitUtils.ConvertToInternalUnits(435, DisplayUnitType.DUT_DECIMAL_INCHES);
                imageY = UnitUtils.ConvertToInternalUnits(360, DisplayUnitType.DUT_DECIMAL_INCHES);
                imageOption = new ImageImportOptions
                {
                    RefPoint = new XYZ(imageX, imageY, 0),
                    Resolution = 50,
                    Placement = BoxPlacement.Center
                };
                _doc.Import($"{GlobalNames.FilesLocationPrefix}HEAD BOLT.jpg", imageOption, draftingView, out image);
                image.get_Parameter(BuiltInParameter.RASTER_SHEETWIDTH)
                    .Set(UnitUtils.ConvertToInternalUnits(45, DisplayUnitType.DUT_DECIMAL_INCHES));
            }
            else
            {
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_o),
                    GetFamilySymbolByName("VBA_COLUMN_BACKGROUND_MEDIUM"),
                    draftingView);
                var imageX = UnitUtils.ConvertToInternalUnits(385, DisplayUnitType.DUT_DECIMAL_INCHES);
                var imageY = UnitUtils.ConvertToInternalUnits(310, DisplayUnitType.DUT_DECIMAL_INCHES);
                var imageOption = new ImageImportOptions
                {
                    RefPoint = new XYZ(imageX, imageY, 0),
                    Resolution = 50,
                    Placement = BoxPlacement.Center
                };
                _doc.Import($"{GlobalNames.FilesLocationPrefix}POWERLAG SCREW.jpg", imageOption, draftingView,
                    out var image);
                image.get_Parameter(BuiltInParameter.RASTER_SHEETWIDTH)
                    .Set(UnitUtils.ConvertToInternalUnits(45, DisplayUnitType.DUT_DECIMAL_INCHES));

                imageX = UnitUtils.ConvertToInternalUnits(435, DisplayUnitType.DUT_DECIMAL_INCHES);
                imageY = UnitUtils.ConvertToInternalUnits(310, DisplayUnitType.DUT_DECIMAL_INCHES);
                imageOption = new ImageImportOptions
                {
                    RefPoint = new XYZ(imageX, imageY, 0),
                    Resolution = 50,
                    Placement = BoxPlacement.Center
                };
                _doc.Import($"{GlobalNames.FilesLocationPrefix}HEAD BOLT.jpg", imageOption, draftingView, out image);
                image.get_Parameter(BuiltInParameter.RASTER_SHEETWIDTH)
                    .Set(UnitUtils.ConvertToInternalUnits(45, DisplayUnitType.DUT_DECIMAL_INCHES));
            }

            if (window == false)
            {
                pt_blk[0] = pt_o[0] + 402;
                pt_blk[1] = pt_o[1] + 261 - 36 * (z <= 190 ? 1 : 0);
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk),
                    GetFamilySymbolByName("VBA_COLUMN_ALT_PICKING_DETAILS"),
                    draftingView);
            }

            pt_blk[0] = ptA[0] + ply_width_x / 2;
            pt_blk[1] = pt_o[1] + 9;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk), GetFamilySymbolByName("VBA_DETAIL_REF"), draftingView);
            ptTemp[0] = pt_blk[0] - 20.5;
            ptTemp[1] = pt_blk[1];
            textNoteOptions = new TextNoteOptions
            {
                VerticalAlignment = VerticalTextAlignment.Middle,
                HorizontalAlignment = HorizontalTextAlignment.Center,
                TypeId = new FilteredElementCollector(_doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .First(q => q.Name == "4.5 mm").Id
            };
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "A", textNoteOptions);
            ptTemp[0] = pt_blk[0] + 6;
            ptTemp[1] = pt_blk[1] + 2;
            textNoteOptions = new TextNoteOptions
            {
                VerticalAlignment = VerticalTextAlignment.Middle,
                HorizontalAlignment = HorizontalTextAlignment.Center,
                TypeId = new FilteredElementCollector(_doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .First(q => q.Name == "3.0 mm").Id
            };
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "SIDE \"A\" PANEL", textNoteOptions);
            ptTemp[0] = pt_blk[0] + 6;
            ptTemp[1] = pt_blk[1] - 2;
            textNoteOptions = new TextNoteOptions
            {
                VerticalAlignment = VerticalTextAlignment.Middle,
                HorizontalAlignment = HorizontalTextAlignment.Center,
                TypeId = new FilteredElementCollector(_doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .First(q => q.Name == "2.25 mm").Id
            };
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "VIEWED FROM PLYWOOD FACE", textNoteOptions);
            if (DrawB)
            {
                pt_blk[0] = ptB[0] + ply_width_y / 2;
                pt_blk[1] = pt_o[1] + 9;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk), GetFamilySymbolByName("VBA_DETAIL_REF"),
                    draftingView);
                ptTemp[0] = pt_blk[0] - 20.5;
                ptTemp[1] = pt_blk[1];
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Middle,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "4.5 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "B", textNoteOptions);
                ptTemp[0] = pt_blk[0] + 6;
                ptTemp[1] = pt_blk[1] + 2;
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Middle,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "3.0 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "SIDE \"B\" PANEL", textNoteOptions);
                ptTemp[0] = pt_blk[0] + 6;
                ptTemp[1] = pt_blk[1] - 2;
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Middle,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "2.25 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "VIEWED FROM PLYWOOD FACE",
                    textNoteOptions);
            }

            if (DrawW)
            {
                pt_blk[0] = ptW[0] + ply_width_w / 2;
                pt_blk[1] = pt_o[1] + 9;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk), GetFamilySymbolByName("VBA_DETAIL_REF"),
                    draftingView);
                ptTemp[0] = pt_blk[0] - 20.5;
                ptTemp[1] = pt_blk[1];
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Middle,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "4.5 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "W", textNoteOptions);
                ptTemp[0] = pt_blk[0] + 6;
                ptTemp[1] = pt_blk[1] + 2;
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Middle,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "3.0 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "POUR WINDOW PANEL", textNoteOptions);
                ptTemp[0] = pt_blk[0] + 6;
                ptTemp[1] = pt_blk[1] - 2;
                textNoteOptions = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Middle,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "2.25 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "VIEWED FROM PLYWOOD FACE",
                    textNoteOptions);
            }

            pt_blk[0] = ptE[0] + ply_width_e / 2;
            pt_blk[1] = pt_o[1] + 9;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk), GetFamilySymbolByName("VBA_DETAIL_REF"), draftingView);
            ptTemp[0] = pt_blk[0] - 20.5;
            ptTemp[1] = pt_blk[1];
            textNoteOptions = new TextNoteOptions
            {
                VerticalAlignment = VerticalTextAlignment.Middle,
                HorizontalAlignment = HorizontalTextAlignment.Center,
                TypeId = new FilteredElementCollector(_doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .First(q => q.Name == "4.5 mm").Id
            };
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "E", textNoteOptions);
            ptTemp[0] = pt_blk[0] + 6;
            ptTemp[1] = pt_blk[1] + 2;
            textNoteOptions = new TextNoteOptions
            {
                VerticalAlignment = VerticalTextAlignment.Middle,
                HorizontalAlignment = HorizontalTextAlignment.Center,
                TypeId = new FilteredElementCollector(_doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .First(q => q.Name == "3.0 mm").Id
            };
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "COLUMN FORM ELEVATION", textNoteOptions);

            var count_str = $"FAB {n_col}-EA";
            pt32[0] = ptE[0] + ply_width_e / 2;
            pt32[1] = pt_o[1] + 7;
            textNoteOptions = new TextNoteOptions
            {
                VerticalAlignment = VerticalTextAlignment.Top,
                HorizontalAlignment = HorizontalTextAlignment.Center,
                TypeId = new FilteredElementCollector(_doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .First(q => q.Name == "2.25 mm").Id
            };
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt32),
                UnitUtils.ConvertToInternalUnits(30, DisplayUnitType.DUT_MILLIMETERS), count_str, textNoteOptions);
        }

        public static XYZ GetXYZByPoint(double[] point)
        {
            var x = UnitUtils.ConvertToInternalUnits(point[0], DisplayUnitType.DUT_DECIMAL_INCHES);
            var y = UnitUtils.ConvertToInternalUnits(point[1], DisplayUnitType.DUT_DECIMAL_INCHES);
            return new XYZ(x, y, 0);
        }

        public static void RotateFamily(FamilyInstance family, double angle)
        {
            _doc.Regenerate();
            if (family.Location is not LocationPoint lp) return;
            var p1 = new XYZ(lp.Point.X, lp.Point.Y, 0);
            var p2 = new XYZ(p1.X, p1.Y, p1.Z + 10);
            var axis = Line.CreateBound(p1, p2);
            lp.Rotate(axis, (Math.PI / 180) * (angle));
        }


        public static FamilySymbol GetFamilySymbolByName(string name)
        {
            FamilySymbol symbol;
            try
            {
                symbol = new FilteredElementCollector(_doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .First(s => s.Name.Equals(name));
            }
            catch (Exception _)
            {
                throw new Exception($"Not found the family: {name}");
            }

            if (!symbol.IsActive) symbol.Activate();
            return symbol;
        }

        private static DetailCurve CalculateDraftingLocation(Document doc, View view)
        {
            var x = UnitUtils.ConvertToInternalUnits(0, DisplayUnitType.DUT_DECIMAL_INCHES);
            var y = UnitUtils.ConvertToInternalUnits(3.71, DisplayUnitType.DUT_DECIMAL_INCHES);
            var position = new XYZ(x, y, 0);
            var arc = Arc.Create(position, 0.1, 0.0, 180 * Math.PI, XYZ.BasisX, XYZ.BasisY);
            return doc.Create.NewDetailCurve(view, arc);
        }

        private static ViewDrafting CreateDraftingView(Document doc, string sheetName, string sheetNum)
        {
            var vd = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(q => q.ViewFamily == ViewFamily.Drafting);

            if (vd == null) return null;
            var draftView = ViewDrafting.Create(doc, vd.Id);
            draftView.Name = $"{sheetNum} - {sheetName}_{DateTime.Now:MM-dd-yyyy-HH-mm-ss}";
            draftView.Scale = 25;
            return draftView;
        }

        private static ViewSheet CreateSheet(Document doc, ColumnCreatorView ui, DrawingTypes type)
        {
            var titleBlock = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                .First(element => element.Name.Contains("TitleBlock"));

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

        public static void CreateFilledRegion(View view, params double[][] points)
        {
            var patterns = new FilteredElementCollector(_doc)
                .OfClass(typeof(FilledRegionType))
                .Cast<FilledRegionType>()
                .First(pattern => pattern.Name.Equals("Concrete"));

            var boundaries = new List<CurveLoop>();
            var curveLoop = new CurveLoop();

            for (var i = 0; i < points.Length - 1; i++)
            {
                var line = Line.CreateBound(GetXYZByPoint(points[i]), GetXYZByPoint(points[i + 1]));
                curveLoop.Append(line);
            }

            boundaries.Add(curveLoop);

            var activeViewId = view.Id;
            FilledRegion.Create(_doc, patterns.Id, activeViewId, boundaries);
        }

        public static void LoadFamilies(Document doc)
        {
            try
            {
                var block = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_DetailComponents)
                    .First(element => element.Name.Contains("VBA_PLY"));
            }
            catch (Exception e)
            {
                TaskDialog.Show("Message", "Wait for the families to be loaded into the project");
                var docHasFamily =
                    Application.RApplication.Application.OpenDocumentFile(
                        $"{GlobalNames.FamiliesLocationPrefix}BasicProject.rvt");
                var textTypesProject = new FilteredElementCollector(docHasFamily).OfClass(typeof(TextNoteType));
                var fillTypesProject = new FilteredElementCollector(docHasFamily).OfClass(typeof(FilledRegionType));
                var dimTypesProject = new FilteredElementCollector(docHasFamily).OfClass(typeof(DimensionType));
                var linesCategoryFilter = new ElementCategoryFilter(BuiltInCategory.OST_Lines);
                var lineTypesProject = new FilteredElementCollector(docHasFamily).OfClass(typeof(GraphicsStyle)).WherePasses(linesCategoryFilter);
                foreach (var wtcur in textTypesProject)
                {
                    var copyIds = new Collection<ElementId> {wtcur.Id};
                    var option = new CopyPasteOptions();
                    ElementTransformUtils.CopyElements(docHasFamily, copyIds, doc, Transform.Identity, option);
                }

                foreach (var wtcur in fillTypesProject)
                {
                    var copyIds = new Collection<ElementId> {wtcur.Id};
                    var option = new CopyPasteOptions();
                    ElementTransformUtils.CopyElements(docHasFamily, copyIds, doc, Transform.Identity, option);
                }

                foreach (var wtcur in dimTypesProject)
                {
                    var copyIds = new Collection<ElementId> {wtcur.Id};
                    var option = new CopyPasteOptions();
                    ElementTransformUtils.CopyElements(docHasFamily, copyIds, doc, Transform.Identity, option);
                }

                foreach (var wtcur in lineTypesProject)
                {
                    var copyIds = new Collection<ElementId> {wtcur.Id};
                    var option = new CopyPasteOptions();
                    ElementTransformUtils.CopyElements(docHasFamily, copyIds, doc, Transform.Identity, option);
                }

                var familyList = new List<string>()
                {
                    "TitleBlock.rfa",
                    "VBA_2X2_NOTES.rfa",
                    "VBA_2X4.rfa",
                    "VBA_2X4_BOLT.rfa",
                    "VBA_2X4_BOLT_HIDDEN.rfa",
                    "VBA_2X4_FACE.rfa",
                    "VBA_2X4_SIDE.rfa",
                    "VBA_2X4_SPAX.rfa",
                    "VBA_3-4_CHAMFER.rfa",
                    "VBA_8-24_CLAMP_PLAN_BACK.rfa",
                    "VBA_8-24_CLAMP_PLAN_OP.rfa",
                    "VBA_8-24_CLAMP_PLAN_OP_WIN.rfa",
                    "VBA_8-24_CLAMP_PROFILE.rfa",
                    "VBA_8-24_CLAMP_PROFILE_FLIPPED.rfa",
                    "VBA_12-36_CLAMP_PLAN_BACK.rfa",
                    "VBA_12-36_CLAMP_PLAN_OP.rfa",
                    "VBA_12-36_CLAMP_PLAN_OP_WIN.rfa",
                    "VBA_12-36_CLAMP_PROFILE.rfa",
                    "VBA_12-36_CLAMP_PROFILE_FLIPPED.rfa",
                    "VBA_24-48_CLAMP_PLAN_BACK.rfa",
                    "VBA_24-48_CLAMP_PLAN_OP.rfa",
                    "VBA_24-48_CLAMP_PLAN_OP_WIN.rfa",
                    "VBA_24-48_CLAMP_PROFILE.rfa",
                    "VBA_24-48_CLAMP_PROFILE_FLIPPED.rfa",
                    "VBA_36_SCISSOR_CLAMP_PLAN.rfa",
                    "VBA_36_SCISSOR_CLAMP_PROFILE.rfa",
                    "VBA_36_SCISSOR_CLAMP_PROFILE_FLIP.rfa",
                    "VBA_48_SCISSOR_CLAMP_PLAN.rfa",
                    "VBA_48_SCISSOR_CLAMP_PROFILE.rfa",
                    "VBA_48_SCISSOR_CLAMP_PROFILE_FLIP.rfa",
                    "VBA_60_SCISSOR_CLAMP_PLAN.rfa",
                    "VBA_60_SCISSOR_CLAMP_PROFILE.rfa",
                    "VBA_60_SCISSOR_CLAMP_PROFILE_FLIP.rfa",
                    "VBA_AFB_7-11.rfa",
                    "VBA_AFB_11-19.rfa",
                    "VBA_AFB_FACE.rfa",
                    "VBA_AFB_PLAN.rfa",
                    "VBA_AFB_SIDE.rfa",
                    "VBA_BRACE_ANGLE.rfa",
                    "VBA_BRACE_NOTES.rfa",
                    "VBA_CHAIN.rfa",
                    "VBA_CHAMF.rfa",
                    "VBA_COIL_ROD.rfa",
                    "VBA_COIL_ROD_14.rfa",
                    "VBA_COIL_ROD_NOTES.rfa",
                    "VBA_COIL_ROD_NOTES_14.rfa",
                    "VBA_COIL_ROD_NOTES_B.rfa",
                    "VBA_COIL_ROD_NOTES_C.rfa",
                    "VBA_COLUMN_ALT_PICKING_DETAILS.rfa",
                    "VBA_COLUMN_BACKGROUND_HIGH.rfa",
                    "VBA_COLUMN_BACKGROUND_MEDIUM.rfa",
                    "VBA_DETAIL_REF.rfa",
                    "VBA_DOWN_PLATE_NOTES.rfa",
                    "VBA_DOWN_PLATE_NOTES_SCISSOR.rfa",
                    "VBA_ELEV_A.rfa",
                    "VBA_ELEVATION_NOTES.rfa",
                    "VBA_ELEVATION_NOTES_MIRRORED.rfa",
                    "VBA_ELEVATION_NOTES_SCREWS.rfa",
                    "VBA_ELEVATION_NOTES_SCREWS_MIRRORED.rfa",
                    "VBA_GATES_SQUARING_CORNER.rfa",
                    "VBA_GATES_SQUARING_CORNER_INV.rfa",
                    "VBA_GATES_SQUARING_CORNER_SLINGS.rfa",
                    "VBA_LIFTING_SLING_A.rfa",
                    "VBA_LIFTING_SLING_B.rfa",
                    "VBA_LIFTING_SLING_C.rfa",
                    "VBA_LIFTING_CORNER_A.rfa",
                    "VBA_LIFTING_CORNER_B.rfa",
                    "VBA_LINE.rfa",
                    "VBA_LVL.rfa",
                    "VBA_LVL_BOLT.rfa",
                    "VBA_LVL_BOLT_HIDDEN.rfa",
                    "VBA_LVL_FACE.rfa",
                    "VBA_LVL_SPAX.rfa",
                    "VBA_NAILING_NOTES_SCISSOR.rfa",
                    "VBA_NUT_5-8.rfa",
                    "VBA_PICKING_CORNER.rfa",
                    "VBA_PLAN_NOTES1.rfa",
                    "VBA_PLAN_NOTES2.rfa",
                    "VBA_PLAN_NOTES3.rfa",
                    "VBA_PLAN_NOTES4.rfa",
                    "VBA_PLAN_NOTES5.rfa",
                    "VBA_PLAN_NOTES6.rfa",
                    "VBA_PLAN_NOTES7.rfa",
                    "VBA_PLAN_NOTES7_LOOP.rfa",
                    "VBA_PLAN_NOTES8.rfa",
                    "VBA_PLAN_NOTES9A.rfa",
                    "VBA_PLAN_NOTES9B.rfa",
                    "VBA_PLY.rfa",
                    "VBA_PLY_SHEET.rfa",
                    "VBA_PLY_WITH_CHAMFER.rfa",
                    "VBA_RECTANGLE.rfa",
                    "VBA_REINFORCING_ANGLE.rfa",
                    "VBA_SCISSOR_CLAMP_NOTES_FRAME.rfa",
                    "VBA_SCREW_HEAD.rfa",
                    "VBA_STUD_HOLDBACK_CALLOUT.rfa",
                    "VBA_STUD_HOLDBACK_DETAIL.rfa",
                    "VBA_TOP_SECTION_DETAILS1.rfa"
                };

                foreach (var family in familyList)
                {
                    doc.LoadFamily($"{GlobalNames.FamiliesLocationPrefix}{family}", out _);
                }
            }
        }
    }
}