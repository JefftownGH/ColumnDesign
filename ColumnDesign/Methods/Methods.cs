using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ColumnDesign.Modules;
using ColumnDesign.UI;
using static ColumnDesign.Modules.ConvertFeetInchesToNumber;
using static ColumnDesign.Modules.ImportMatrixFunction;
using static ColumnDesign.Modules.ReadSizesFunction;
using static ColumnDesign.Modules.ImportMatrixFunction;

namespace ColumnDesign.Methods
{
    public static class Methods
    {
        private static Document _doc;
        private static ColumnCreatorView _ui;
        private static double win_clamp_top_max = 5;
        private static int win_clamp_bot_max = 4;
        private const int bot_clamp_gap = 8;
        private static double WinPos;
        private const double WinGap = 6;
        private const double WinStudOff = 1;

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
            double WinPos;
            double WinGap;
            double WinStudOff;

            WinX = _ui.WindowX.IsChecked == true;
            WinY = _ui.WindowY.IsChecked == true;
            if (WinX || WinY) window = true;
            WinPos = ConvertToNum(_ui.WinDim2.Text);
            WinGap = 6;
            WinStudOff = 1;
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
            // ply_name = ColumnCreator.PlyNameBox
            if (WinX)
            {
                x = ConvertToNum(_ui.LengthY.Text);
                y = ConvertToNum(_ui.WidthX.Text);
                WinX = false;
                WinY = true;
            }

            if (z < 16 * 12)
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

            string stud_block_bolt_hidden;
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
            //                                                                                           S T U D S
            // #####################################################################################
            int row_num = 0;
            int col_num_x = 0;
            int col_num_y = 0;
            int n_studs_total;
            stud_matrix = ImportMatrix(@$"{GlobalNames.WtFileLocationPrefix}Columns\n_stud_matrix.csv");
            for (var i = 0; i < stud_matrix.GetLength(0); i++)
            {
                if (stud_matrix[i, 0] > z)
                {
                    row_num = i;
                    break;
                }
            }

            for (var i = 0; i < stud_matrix.GetLength(1); i++)
            {
                if (stud_matrix[0, i] == x)
                {
                    col_num_x = i;
                    break;
                }
            }

            for (var i = 0; i < stud_matrix.GetLength(1); i++)
            {
                if (stud_matrix[0, i] == y)
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

            double[] stud_spacing_x = new double[n_studs_x - 1];
            double[] stud_spacing_y = new double[n_studs_y - 1];
            double available_2x2_x;
            double available_2x2_y;
            double min_2x2_gap = 1.625;
            int AFB_x2;
            for (var i = 0; i < n_studs_x - 1; i++)
            {
                stud_spacing_x[i] = stud_start_offset + i * (3.5 + avg_gap_x);
            }

            for (var i = 0; i < n_studs_y - 1; i++)
            {
                stud_spacing_y[i] = stud_start_offset + i * (3.5 + avg_gap_y);
            }

            if (push_studs_x == 1)
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

            for (var i = 0; i < n_studs_x - 2; i++)
            {
                if (stud_spacing_x[i + 1] - stud_spacing_x[i] < 3.5)
                {
                    stud_spacing_x[i] = stud_spacing_x[i + 1] - 3.5;
                }
            }

            for (var i = 0; i < n_studs_y - 2; i++)
            {
                if (stud_spacing_y[i + 1] - stud_spacing_y[i] < 3.5)
                {
                    stud_spacing_y[i] = stud_spacing_y[i + 1] - 3.5;
                }
            }

            for (var i = 0; i < n_studs_x - 2; i++)
            {
                if (stud_spacing_x[i + 1] - stud_spacing_x[i] < 3.5)
                {
                    TaskDialog.Show("Error", "Error: Studs overlap, check and manually correct drawing.");
                    goto EndOfStudChecks;
                }
            }

            for (var i = 0; i < n_studs_y - 2; i++)
            {
                if (stud_spacing_y[i + 1] - stud_spacing_y[i] < 3.5)
                {
                    TaskDialog.Show("Error", "Error: Studs overlap, check and manually correct drawing.");
                    goto EndOfStudChecks;
                }
            }

            if (window)
            {
                for (var i = 0; i < n_studs_x - 2; i++)
                {
                    if (stud_spacing_x[i + 1] - stud_spacing_x[i] < 5)
                    {
                        TaskDialog.Show("Error",
                            "Warning: Insufficient clearance between studs for a 2x2 window lock. Check and manually correct drawing if necessary.");
                        goto EndOfStudChecks;
                    }
                }

                for (var i = 0; i < n_studs_y - 2; i++)
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
            //                                                                                  P L Y W O O D
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
            if (UpdatePly.ValidatePlySeams(_ui, ply_seams, x, y, z) == 0)
            {
                throw new Exception("Plywood layout invalid. You should never see this message...How did you do this?");
            }

            var temp_1 = 0d;

            for (var i = 0; i < ply_seams.Length; i++)
            {
                temp_1 += ply_seams[i];
                if (Math.Round(temp_1, 3) == Math.Round(WinPos, 3))
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
                        for (var j = 0; j < ply_seams.Length; j++)
                        {
                            ply_seams_win[j + 1] = ply_seams[j];
                        }
                    }

                    break;
                }
            }

            ply_width_x = x + 1.5;
            ply_width_y = y + 1.5;
            int n_studs_w;
            double ply_width_w = 0;
            double[] stud_spacing_w = new double[0];
            int n_studs_e;
            double ply_width_e = 0;
            double[] stud_spacing_e = new double[0];
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

            ply_bot_n = 0;
            for (var i = 0; i < ply_seams.Length; i++)
            {
                if (Math.Round(ply_seams[i], 3) == Math.Round(max_ply_ht, 3))
                {
                    ply_bot_n++;
                }
            }

            int unique_plys;
            double[] ply_widths = new double[1];
            double[,] ply_cuts = new double[2, 0];
            unique_plys = 0;
            ply_widths[0] = ply_width_x;
            ply_widths[1] = ply_width_y;
            for (var i = 0; i < ply_seams.Length; i++)
            {
                for (var j = 0; j < ply_widths.Length; j++)
                {
                    for (var l = 0; l < ply_cuts.GetLength(1); l++)
                    {
                        if (Math.Round(ply_cuts[0, l], 3) == Math.Round(ply_widths[j], 3) &&
                            Math.Round(ply_cuts[1, l], 3) == Math.Round(ply_seams[i], 3))
                        {
                            ply_cuts[2, l] += 2 - (WinX == true ? 1 : 0) * (j == 0 ? 1 : 0) -
                                              (WinY == true ? 1 : 0) * (j == 1 ? 1 : 0);
                            break;
                        }
                        else if (l == ply_cuts.GetLength(1))
                        {
                            unique_plys++;
                            ply_cuts = ResizeArray<double>(ply_cuts, 3, unique_plys);
                            ply_cuts[0, unique_plys] = ply_widths[j];
                            ply_cuts[1, unique_plys] = ply_seams[i];
                            ply_cuts[2, unique_plys] = 2 - (WinX == true ? 1 : 0) * (j == 0 ? 1 : 0) -
                                                       (WinY == true ? 1 : 0) * (j == 1 ? 1 : 0);
                            break;
                        }
                    }
                }
            }

            if (Math.Round(ply_cuts[0, ply_cuts.GetLength(1)], 3) == 0)
            {
                ply_cuts = ResizeArray<double>(ply_cuts, 3, ply_cuts.GetLength(1) - 1);
                unique_plys--;
            }

            for (var i = 0; i < ply_seams_win.Length; i++)
            {
                for (var j = 0; j < ply_cuts.GetLength(1); j++)
                {
                    if (Math.Round(ply_cuts[1, j], 3) == Math.Round(ply_width_w, 3) &&
                        Math.Round(ply_cuts[2, j], 3) == Math.Round(ply_seams_win[i], 3))
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

            if (Math.Round(ply_cuts[0, ply_cuts.GetLength(1) - 1], 3) == 0)
            {
                ply_cuts = ResizeArray(ply_cuts, 3, ply_cuts.GetLength(1) - 1);
                unique_plys--;
            }

            var msg = "";
            for (var i = 0; i < ply_cuts.GetLength(1); i++)
            {
                for (var j = 0; j < ply_cuts.GetLength(0); j++)
                {
                    msg += $"{ply_cuts[j, i]}\t";
                }

                msg += "\n";
            }

            TaskDialog.Show("Message", msg);

            // #####################################################################################
            //                                                                              C L A M P S
            // #####################################################################################
            var clamp_spacing = new int[] { };
            double swing_ang;
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
                clamp_block_bk = GlobalNames.WtClampPlanBack824;
                clamp_block_op = GlobalNames.WtClampPlanOp824;
                clamp_block_pr = GlobalNames.WtClampProfile824;
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

            var clamp_matrix = ImportMatrix(@$"{GlobalNames.WtFileLocationPrefix}Columns\clamp_matrix.csv");
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

            Array.Resize(ref clamp_spacing, clamp_matrix.GetLength(1));
            var k = 0;
            for (var i = 0; i < clamp_matrix.GetLength(1); i++)
            {
                if (clamp_matrix[row_num, i] == 0)
                {
                    Array.Resize(ref clamp_spacing, k);
                    break;
                }
                else
                {
                    clamp_spacing[i] = clamp_matrix[row_num, i];
                    k++;
                }
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
            for (var i = 0; i < clamp_spacing.Length; i++)
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
                            Array.Resize(ref clamp_spacing_con, n_clamps);
                            for (var j = i; j < n_clamps; j++)
                            {
                                clamp_spacing_con[n_clamps - j + i] = clamp_spacing_con[n_clamps - j + i - 1];
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
                for (var i = 0; i < clamp_spacing_con.Length; i++)
                {
                    io = clamp_spacing_con.Length - 1 - i;
                    if (clamp_spacing_con[io] > WinPos)
                    {
                        if ((clamp_spacing_con[io] - WinPos) > win_clamp_top_max)
                        {
                            n_clamps++;
                            Array.Resize(ref clamp_spacing_con, n_clamps);
                            clamp_spacing_con[clamp_spacing_con.Length - 1] = (int) WinPos + (int) win_clamp_top_max;
                            for (var j = 0; j < clamp_spacing_con.Length; j++)
                            {
                                for (var l = 0; l < clamp_spacing_con.Length - 1; l++)
                                {
                                    if (clamp_spacing_con[l + 1] > clamp_spacing_con[l])
                                    {
                                        temp_value = clamp_spacing_con[l];
                                        clamp_spacing_con[l] = clamp_spacing_con[l + 1];
                                        clamp_spacing_con[l + 1] = temp_value;
                                    }
                                }
                            }

                            if (clamp_spacing_con[io] - clamp_spacing_con[io - 1] < 8)
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

                                goto UpperWinClampSet;
                            }

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
                Array.Resize(ref clamp_spacing, n_clamps + 1);
                clamp_spacing_con[0] = (int) z - clamp_spacing_con[0];
                for (var i = 1; i < n_clamps; i++)
                {
                    clamp_spacing_con[i] = clamp_spacing_con[i - 1] - clamp_spacing_con[i];
                }

                clamp_spacing[n_clamps] = clamp_spacing_con[clamp_spacing_con.Length - 2];
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
            //                                                               M I S C E L L A N E O U S
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

            int n = 0;
            do
            {
                k++;
            } while (clamp_spacing_con[n] > z * 0.7);

            int brace_clamp;
            if (Math.Abs(clamp_spacing_con[n] - z * 0.7) < Math.Abs(clamp_spacing_con[n - 1] - z * 0.7))
            {
                brace_clamp = n;
            }
            else
            {
                brace_clamp = n - 1;
            }

            do
            {
                brace_clamp--;
            } while (clamp_spacing_con[brace_clamp] + 0.9289 < brace_L_stored + 4);

            int chain_clamp = 1;
            do
            {
                chain_clamp++;
            } while (clamp_spacing_con[chain_clamp] > clamp_spacing_con[brace_clamp] + 0.9289 - brace_L_stored + 1);

            chain_clamp--;
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
            //                                                                                   T E X T
            // #####################################################################################
            double[] pt_click;
            string qty_text;
            var pt_o = new double[2];
            var pt1 = new double[2];
            var pt2 = new double[2];
            var pt3 = new double[2];
            pt_o[0] = 20.23568507;
            pt_o[1] = 5.17943146;
            pt_o[2] = 0;
            pt1[0] = pt_o[0] + 448.4;
            pt1[1] = pt_o[1] + 360;
            pt2[0] = pt1[0] + 0;
            pt2[1] = pt1[1] + 7;
            pt3[0] = pt1[0] + 0;
            pt3[1] = pt1[1] - 52;
            //TODO TEXT

            // #####################################################################################
            //                                                                                        D R A W I N G
            // #####################################################################################
            var ptA = new double[3];
            var ptB = new double[3];
            var ptW = new double[3];
            var ptE = new double[3];
            var pt4 = new double[3];
            var pt5 = new double[3];
            var pt6 = new double[3];
            var pt7 = new double[3];
            var pt8 = new double[3];
            var pt9 = new double[3];
            var pt10 = new double[3];
            var pt11 = new double[3];
            var pt12 = new double[3];
            var pt13 = new double[3];
            var pt14 = new double[3];
            var pt15 = new double[3];
            var pt16 = new double[3];
            var pt17 = new double[3];
            var pt18 = new double[3];
            var pt20 = new double[3];
            var pt21 = new double[3];
            var pt22 = new double[3];
            var pt23 = new double[3];
            var pt24 = new double[3];
            var pt25 = new double[3];
            var pt26 = new double[3];
            var pt27 = new double[3];
            var pt28 = new double[3];
            var pt29 = new double[3];
            var pt30 = new double[3];
            var pt31 = new double[3];
            var pt32 = new double[3];
            var pt33 = new double[3];
            var pt34 = new double[3];
            var pt_blk = new double[3];
            bool DrawB;
            bool DrawW;
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

            // _doc.Create.NewFamilyInstance(clampPlanBack824Location,
            //     GetFamilySymbolByName(GlobalNames.WtClampPlanBack824), draftingView);
            // _doc.Create.NewFamilyInstance(clampPlanOp824Location,
            //     GetFamilySymbolByName(GlobalNames.WtClampPlanOp824),
            //     draftingView);
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