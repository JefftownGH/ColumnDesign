using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ColumnDesign.UI;
using static ColumnDesign.Methods.Methods;
using static ColumnDesign.Modules.ConvertFeetInchesToNumber;
using static ColumnDesign.Modules.ConvertNumberToFeetInches;
using static ColumnDesign.Modules.ImportMatrixFunction;
using static ColumnDesign.Modules.ReadSizesFunction;

namespace ColumnDesign.Modules
{
    public class zScissorClampModule
    {
        private static Document _doc;
        private static UIDocument _uiDoc;

        public static void CreateScissorClamp(ColumnCreatorView _ui, UIDocument uiDoc, View draftingView)
        {
            _uiDoc = uiDoc;
            _doc = uiDoc.Document;
            var filepath = @$"{GlobalNames.FilesLocationPrefix}Columns\scissor_clamp_matrix.csv";
            //     Dim version As String
//     Dim col_ver_str As String
//     On Error Resume Next
//         version = ThisDrawing.Application.version
//         version = Mid(version, 1, 2) 'Records the first 2 characters in the string as the version number
//         col_ver_str = "AutoCAD.AcCmColor." & version
//         Set color = AcadApplication.GetInterfaceObject(col_ver_str) 'Use version number to set color object
//
//         color.ColorIndex = 154:     Set LayerObj = ThisDrawing.Layers.Add("DIM"):       LayerObj.TrueColor = color:    LayerObj.Linetype = "Continuous"
//         color.ColorIndex = acRed:   Set LayerObj = ThisDrawing.Layers.Add("ELEV"):      LayerObj.TrueColor = color:    LayerObj.Linetype = "Continuous"
//         color.ColorIndex = acCyan:  Set LayerObj = ThisDrawing.Layers.Add("FMTEXTA"):   LayerObj.TrueColor = color:    LayerObj.Linetype = "Continuous"
//         color.ColorIndex = acBlue:  Set LayerObj = ThisDrawing.Layers.Add("FORM"):      LayerObj.TrueColor = color:    LayerObj.Linetype = "Continuous"
//         color.ColorIndex = acBlue:  Set LayerObj = ThisDrawing.Layers.Add("NOTES"):     LayerObj.TrueColor = color:    LayerObj.Linetype = "Continuous"
//         color.ColorIndex = acCyan:  Set LayerObj = ThisDrawing.Layers.Add("PLYA"):      LayerObj.TrueColor = color:    LayerObj.Linetype = "Continuous"
//         color.ColorIndex = 23:      Set LayerObj = ThisDrawing.Layers.Add("SHADE"):     LayerObj.TrueColor = color:    LayerObj.Linetype = "Continuous"
//         color.ColorIndex = acBlue:  Set LayerObj = ThisDrawing.Layers.Add("STEEL"):     LayerObj.TrueColor = color:    LayerObj.Linetype = "Continuous"
//     UsedLayerList(1) = "DIM": UsedLayerList(2) = "ELEV": UsedLayerList(3) = "FMTEXTA": UsedLayerList(4) = "FORM": UsedLayerList(5) = "NOTES":
//     UsedLayerList(6) = "PLYA": UsedLayerList(7) = "SHADE": UsedLayerList(8) = "STEEL":
            double x;
            double y;
            double z;
            int n_col;
            int n_studs_x = 0;
            int n_studs_y = 0;
            int n_clamps = 0;
            int n_top_clamps;
            int stud_type;
            int clamp_size;
            int col_wt;
            string ply_name;
            string stud_name;
            string stud_name_full;
            string stud_block;
            string stud_block_spax;
            string stud_block_bolt;
            string stud_face_block;
            var ply_thk = 0.75;
            var stud_base_gap = 0.25;
            int bot_clamp_gap = 6;
            var stud_start_offset_x = 0.125;
            var stud_start_offset_y = 2.375;
            x = ConvertToNum(_ui.SWidthX.Text);
            y = ConvertToNum(_ui.SLengthY.Text);
            z = ConvertToNum(_ui.SHeightZ.Text);
            n_col = (int) ConvertToNum(_ui.SQuantity.Text);
            ply_name = _ui.SPlywoodType.Text;
            stud_type = 1;
            stud_name = "2X4";
            stud_name_full = "2X4";
            stud_block = "VBA_2X4";
            stud_block_spax = "VBA_2X4_SPAX";
            stud_block_bolt = "VBA_2X4_BOLT";
            stud_face_block = "VBA_2X4_FACE";
            double ply_width_x = x;
            double ply_width_y = y + 4.5;
            double[] ptTemp = new double[2];
            // '#####################################################################################
            // '                                     S T U D S
            // '#####################################################################################
            var row_num = 0;
            var col_num_x = 0;
            var col_num_y = 0;
            int n_studs_total;
            var stud_matrix = new int[20, 2];
            stud_matrix[0, 0] = 8;
            stud_matrix[0, 1] = 2;
            stud_matrix[1, 0] = 10;
            stud_matrix[1, 1] = 2;
            stud_matrix[2, 0] = 12;
            stud_matrix[2, 1] = 2;
            stud_matrix[3, 0] = 14;
            stud_matrix[3, 1] = 3;
            stud_matrix[4, 0] = 16;
            stud_matrix[4, 1] = 3;
            stud_matrix[5, 0] = 18;
            stud_matrix[5, 1] = 3;
            stud_matrix[6, 0] = 20;
            stud_matrix[6, 1] = 3;
            stud_matrix[7, 0] = 22;
            stud_matrix[7, 1] = 4;
            stud_matrix[8, 0] = 24;
            stud_matrix[8, 1] = 4;
            stud_matrix[9, 0] = 26;
            stud_matrix[9, 1] = 4;
            stud_matrix[10, 0] = 28;
            stud_matrix[10, 1] = 4;
            stud_matrix[11, 0] = 30;
            stud_matrix[11, 1] = 5;
            stud_matrix[12, 0] = 32;
            stud_matrix[12, 1] = 5;
            stud_matrix[13, 0] = 34;
            stud_matrix[13, 1] = 5;
            stud_matrix[14, 0] = 36;
            stud_matrix[14, 1] = 5;
            stud_matrix[15, 0] = 38;
            stud_matrix[15, 1] = 5;
            stud_matrix[16, 0] = 40;
            stud_matrix[16, 1] = 6;
            stud_matrix[17, 0] = 42;
            stud_matrix[17, 1] = 6;
            stud_matrix[18, 0] = 44;
            stud_matrix[18, 1] = 6;
            stud_matrix[19, 0] = 46;
            stud_matrix[19, 1] = 6;
            for (int i = 0; i < stud_matrix.GetLength(0); i++)
            {
                if (stud_matrix[i, 0] >= x)
                {
                    n_studs_x = stud_matrix[i, 1];
                    break;
                }
            }

            for (int i = 0; i < stud_matrix.GetLength(0); i++)
            {
                if (stud_matrix[i, 0] >= y)
                {
                    n_studs_y = stud_matrix[i, 1];
                    break;
                }
            }

            n_studs_total = n_studs_x * 2 + n_studs_y * 2;
            var stud_spacing_x = new double[n_studs_x];
            var stud_spacing_y = new double[n_studs_y];
            for (var i = 0; i < n_studs_x; i++)
            {
                stud_spacing_x[i] = stud_start_offset_x +
                                    (i) * (ply_width_x - 2 * stud_start_offset_x - 3.5) / (n_studs_x - 1);
            }

            for (var i = 0; i < n_studs_y; i++)
            {
                stud_spacing_y[i] = stud_start_offset_y +
                                    (i) * (ply_width_y - 2 * stud_start_offset_y - 3.5) / (n_studs_y - 1);
            }

            // '#####################################################################################
            // '                                   P L Y W O O D
            // '#####################################################################################
            double ply_top_ht_min = 6;
            double max_ply_ht;
            var ply_seams = new double[1];
            ply_seams = ReadSizes(_ui.SBoxPlySeams.Text);
            if (sUpdatePly_Function.sValidatePlySeams(_ui, ply_seams, x, y, z) == 0)
            {
                throw new Exception("Plywood layout invalid. You should never see this message...How did you do this?");
            }

            int n_studs_e;
            double ply_width_e = 0;
            double[] stud_spacing_e = new double[stud_spacing_x.Length];
            n_studs_e = n_studs_x;
            ply_width_e = ply_width_x;
            for (int i = 0; i < stud_spacing_x.Length; i++)
            {
                stud_spacing_e[i] = stud_spacing_x[i];
            }

            if (ply_width_x > 48 || ply_width_y > 48)
            {
                max_ply_ht = 48;
            }
            else
            {
                max_ply_ht = 96;
            }

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
                            ply_cuts[2, l] += 2;
                            break;
                        }
                        else if (l == ply_cuts.GetLength(1) - 1)
                        {
                            unique_plys++;
                            ply_cuts = ResizeArray<double>(ply_cuts, 3, unique_plys);
                            ply_cuts[0, unique_plys - 1] = ply_widths[j];
                            ply_cuts[1, unique_plys - 1] = ply_seams[i];
                            ply_cuts[2, unique_plys - 1] += 2;
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

            // '#####################################################################################
            // '                                   C L A M P S
            // '#####################################################################################
            var clamp_spacing = new double[1];
            var bigSide = new[] {x, y}.Max();
            int n_pours;
            double clamp_L;
            string clamp_name;
            double clamp_width;
            string clamp_block;
            string clamp_block_pr;
            string clamp_block_pr_flip;

            if (_ui.Clampsize1.IsChecked == true || _ui.ClampSizeChoose.IsChecked == true && bigSide < 23.5)
            {
                clamp_size = 4;
                clamp_L = 36;
                clamp_width = 2.5;
                clamp_name = "36\" SCISSOR CLAMP";
                clamp_block = "VBA_36_SCISSOR_CLAMP_PLAN";
                clamp_block_pr = "VBA_36_SCISSOR_CLAMP_PROFILE";
                clamp_block_pr_flip = "VBA_36_SCISSOR_CLAMP_PROFILE_FLIP";
            }
            else if (_ui.Clampsize2.IsChecked == true || _ui.ClampSizeChoose.IsChecked == true && bigSide <= 35.5)
            {
                clamp_size = 5;
                clamp_L = 48;
                clamp_width = 2.5;
                clamp_name = "48\" SCISSOR CLAMP";
                clamp_block = "VBA_48_SCISSOR_CLAMP_PLAN";
                clamp_block_pr = "VBA_48_SCISSOR_CLAMP_PROFILE";
                clamp_block_pr_flip = "VBA_48_SCISSOR_CLAMP_PROFILE_FLIP";
            }
            else if (_ui.Clampsize3.IsChecked == true || _ui.ClampSizeChoose.IsChecked == true && bigSide <= 46)
            {
                clamp_size = 6;
                clamp_L = 60;
                clamp_width = 3;
                clamp_name = "60\" SCISSOR CLAMP";
                clamp_block = "VBA_60_SCISSOR_CLAMP_PLAN";
                clamp_block_pr = "VBA_60_SCISSOR_CLAMP_PROFILE";
                clamp_block_pr_flip = "VBA_60_SCISSOR_CLAMP_PROFILE_FLIP";
            }
            else throw new Exception("Error: Column is too wide (>46\") in one dimension");

            var clamp_matrix = ImportMatrix(filepath);
            var long_side = x <= y ? y : x;
            row_num = 0;
            for (var i = 0; i < clamp_matrix.GetLength(0); i++)
            {
                if (clamp_matrix[i, 1] >= long_side)
                {
                    row_num = i;
                    break;
                }
            }

            n_pours = (int) (z / clamp_matrix[row_num, 0]);
            if (n_pours * clamp_matrix[row_num, 0] < z)
            {
                n_pours++;
            }

            var z_original = z;
            var pour_overlap = 0d;
            var multi_clamp_spacing = new double[100, n_pours];
            var total_clamp_spacings = new int[n_pours];
            for (var i = 0; i < n_pours; i++)
            {
                if (n_pours >= 2)
                {
                    if (z_original - (i + 1) * clamp_matrix[row_num, 0] >= 0)
                    {
                        z = clamp_matrix[row_num, 0];
                    }
                    else
                    {
                        z = z_original - i * clamp_matrix[row_num, 0];
                        if (z < 12)
                        {
                            pour_overlap = 12 - z;
                            z = 12;
                        }
                    }
                }

                Array.Resize(ref clamp_spacing, clamp_matrix.GetLength(1) - 1);
                var k = 0;
                for (int j = 1; j < clamp_matrix.GetLength(1) - 1; j++)
                {
                    if (clamp_matrix[row_num, j + 1] == 0)
                    {
                        Array.Resize(ref clamp_spacing, k);
                        break;
                    }
                    else
                    {
                        clamp_spacing[j - 1] = clamp_matrix[row_num, j + 1];
                        k++;
                    }
                }

                var ht_rem = z - bot_clamp_gap;
                var n_clamps_temp = 0;
                for (int j = 0; j < clamp_spacing.Length - 1; j++)
                {
                    if (clamp_spacing[j] <= ht_rem)
                    {
                        n_clamps_temp++;
                        ht_rem -= clamp_spacing[j];
                    }
                    else
                    {
                        n_clamps_temp++;
                        break;
                    }
                }

                Array.Resize(ref clamp_spacing, n_clamps_temp + 1);
                clamp_spacing[clamp_spacing.Length - 1] = bot_clamp_gap;
                var tot_temp = 0;
                for (int j = 0; j < clamp_spacing.Length; j++)
                {
                    tot_temp += (int) clamp_spacing[j];
                }

                clamp_spacing[clamp_spacing.Length - 2] -= tot_temp - z;
                TestClampSpacing:
                var infExit = 0;
                for (var j = 1; j < clamp_spacing.Length; j++)
                {
                    if (clamp_spacing[j] < 5)
                    {
                        clamp_spacing[j - 1] -= 5 - clamp_spacing[j];
                        clamp_spacing[j] = 5;


                        infExit++;
                        if (infExit > 100)
                        {
                            throw new TimeoutException(
                                "Error: Infinite loop encountered while computing clamp spacings.");
                        }
                        else
                        {
                            goto TestClampSpacing;
                        }
                    }
                }

                for (int j = 0; j < clamp_spacing.Length; j++)
                {
                    multi_clamp_spacing[j, i] = clamp_spacing[j];
                }

                total_clamp_spacings[i] = clamp_spacing.Length;
                n_clamps += n_clamps_temp;
                if (i >= 1) n_clamps++;
            }

            var clamp_spacing_r = new int[100];
            clamp_spacing = new double[100];
            var jCount = 0;
            for (int p = n_pours - 1; p >= 0; p--)
            {
                for (int i = 0; i < total_clamp_spacings[p]; i++)
                {
                    if (multi_clamp_spacing[i, p] != 0)
                    {
                        if ((p != (n_pours - 1)) && i == 0)
                        {
                            clamp_spacing[jCount - 1] = multi_clamp_spacing[i, p];
                            n_clamps--;
                        }
                        else
                        {
                            clamp_spacing[jCount] = multi_clamp_spacing[i, p];
                            jCount++;
                        }
                    }
                }
            }

            if (pour_overlap != 0)
            {
                clamp_spacing[total_clamp_spacings[n_pours - 1] - 1] += pour_overlap;
                clamp_spacing[total_clamp_spacings[n_pours - 1]] -= pour_overlap;
            }

            for (int i = 0; i < clamp_spacing.Length; i++)
            {
                if (clamp_spacing[i] == 0 && clamp_spacing[i + 1] == 0)
                {
                    Array.Resize(ref clamp_spacing, i - 1);
                    break;
                }
            }

            if (clamp_spacing[0] < 2 && clamp_spacing[0] >= 0)
            {
                if (clamp_spacing[1] >= (2 - clamp_spacing[0]) + 5)
                {
                    clamp_spacing[1] -= 2 - clamp_spacing[0];
                    clamp_spacing[0] = 2;
                }
                else if (clamp_spacing[2] >= (2 - clamp_spacing[0]) + 5)
                {
                    clamp_spacing[2] -= 2 - clamp_spacing[0];
                    clamp_spacing[0] = 2;
                }
            }

            z = z_original;
            var clamp_spacing_con = new double[clamp_spacing.Length];
            clamp_spacing_con[0] = z - clamp_spacing[0];
            for (int i = 1; i < clamp_spacing.Length; i++)
            {
                clamp_spacing_con[i] = clamp_spacing_con[i - 1] - clamp_spacing[i];
            }

            // '#####################################################################################
            // '                                     T E X T
            // '#####################################################################################
            var qty_text = "";
            var pt_o = new double[2];
            var pt1 = new double[2];
            var pt2 = new double[2];
            var pt3 = new double[2];
            double panel_wt;
            pt_o[0] = 20.23568507;
            pt_o[1] = 5.17943146;
            pt1[0] = pt_o[0] + 548.37926872;
            pt1[1] = pt_o[1] + 369.60966303;

            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt1), GetFamilySymbolByName("VBA_SCISSOR_CLAMP_NOTES_FRAME"),
                draftingView);
            col_wt = CalcWeightFunction.wt_total(x, y, z, n_studs_x, n_studs_y, stud_type, n_clamps, clamp_size, 0);
            panel_wt = CalcWeightFunction.wt_panel(x, y, z, n_studs_x, n_studs_y, stud_type);
            pt1[0] = pt_o[0] + 448.4;
            pt1[1] = pt_o[1] + 360;
            pt2[0] = pt1[0] + 0;
            pt2[1] = pt1[1] + 7;
            pt3[0] = pt1[0] + 0;
            pt3[1] = pt1[1] - 68;
            var strFabNotes = $"• COLUMN SIZE = {x}\" X + {y}\"\n\n" +
                              $"• NUMBER OF COLUMN FORMS = {n_col}-EA\n\n" +
                              $"• COLUMN FORM WEIGHT (APPROXIMATE) = {col_wt}-LBS\n" +
                              $"• COLUMN PANEL WEIGHT (SINGLE PANEL) = {panel_wt}-LBS\n\n" +
                              $"• PLYWOOD = 3/4'' PLYFORM (\"{ply_name}\"), CLASS-1 (MIN)\n\n" +
                              "• COLUMN FORMS AND CLAMP SPACING LAYOUTS FOR SCISSOR CLAMPS ARE DESIGNED FOR A POUR RATE = ";
            if (z <= clamp_matrix[row_num, 0])
            {
                strFabNotes += "FULL LIQUID HEAD U.N.O.\n\n";
            }
            else
            {
                strFabNotes += $"{ConvertFtIn(clamp_matrix[row_num, 0])}\n\n";
            }

            strFabNotes +=
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
                UnitUtils.ConvertToInternalUnits(95, DisplayUnitType.DUT_MILLIMETERS), strFabNotes, textNoteOptions);
            for (int i = 0; i < ply_cuts.GetLength(1); i++)
            {
                qty_text +=
                    $"• ({ply_cuts[2, i] * n_col}-EA) = ({n_col}-COL) X ({ply_cuts[2, i]}-EA/COL) @ {ConvertFtIn(ply_cuts[0, i])} WIDE X {ConvertFtIn(ply_cuts[1, i])} LONG 3/4'' PLYWOOD\n";
            }

            qty_text = $"PLYWOOD\n{qty_text}\n";
            qty_text +=
                $"STUDS\n• ({n_studs_total * n_col}-EA) = ({n_col}-COL) X ({n_studs_total}-EA/COL) @ {ConvertFtIn(z - 0.25)} {stud_name_full}";
            qty_text += "\n\nSCISSOR CLAMP SETS (2 CLAMPS PER SET)\n";
            qty_text +=
                $"• ({n_clamps * n_col}-EA) = ({n_col}-COL) X ({n_clamps}-EA/COL) @ GATES {clamp_name} LOK-FAST CLAMP ASSEMBLIES (SETS).";
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

//     '#####################################################################################
//     '                                 D R A W I N G
//     '#####################################################################################
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
            var pt19 = new double[2];
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
            bool DrawB = true;

            ptA[0] = pt_o[0] + 300;
            ptA[1] = pt_o[1] + 25;
            ptB[0] = pt_o[0] + 185;
            ptB[1] = pt_o[1] + 25;
            ptE[0] = pt_o[0] + 54.5;
            ptE[1] = pt_o[1] + 25;
            FamilyInstance family;
            family = _doc.Create.NewFamilyInstance(
                GetXYZByPoint(new double[] {ptE[0] - ply_thk - 1.5, ptE[1] + stud_base_gap}),
                GetFamilySymbolByName("VBA_RECTANGLE"),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("xB")
                .Set(UnitUtils.ConvertToInternalUnits(1.5, DisplayUnitType.DUT_DECIMAL_INCHES));
            family.LookupParameter("yB").Set(UnitUtils.ConvertToInternalUnits(z, DisplayUnitType.DUT_DECIMAL_INCHES));

            family = _doc.Create.NewFamilyInstance(
                GetXYZByPoint(new double[] {ptE[0] + ply_width_e + ply_thk, ptE[1] + stud_base_gap}),
                GetFamilySymbolByName("VBA_RECTANGLE"),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("xB")
                .Set(UnitUtils.ConvertToInternalUnits(1.5, DisplayUnitType.DUT_DECIMAL_INCHES));
            family.LookupParameter("yB").Set(UnitUtils.ConvertToInternalUnits(z, DisplayUnitType.DUT_DECIMAL_INCHES));
            for (var i = 0; i < n_studs_x; i++)
            {
                pt6[0] = ptA[0] + ply_width_x - 3.5 - stud_spacing_x[i];
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
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
            }

            for (var i = 0; i < n_studs_y; i++)
            {
                pt6[0] = ptB[0] + ply_width_y - 3.5 - stud_spacing_y[i];
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
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
            }

            pt4[1] = ptA[1];
            foreach (var t in ply_seams)
            {
                pt4[0] = ptA[0];
                pt4[1] = pt4[1] + t;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt4), GetFamilySymbolByName("VBA_PLY_SHEET"),
                    draftingView);
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_DECIMAL_INCHES));
                family.LookupParameter("Distance2")
                    ?.Set(UnitUtils.ConvertToInternalUnits(t, DisplayUnitType.DUT_DECIMAL_INCHES));

                pt4[0] = ptB[0];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt4), GetFamilySymbolByName("VBA_PLY_SHEET"),
                    draftingView);
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));
                family.LookupParameter("Distance2")
                    ?.Set(UnitUtils.ConvertToInternalUnits(t, DisplayUnitType.DUT_DECIMAL_INCHES));
            }

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
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
            }

            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(new double[] {ptE[0], ptE[1]}),
                GetFamilySymbolByName("VBA_RECTANGLE"),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("xB")
                .Set(UnitUtils.ConvertToInternalUnits(ply_width_e, DisplayUnitType.DUT_DECIMAL_INCHES));
            family.LookupParameter("yB").Set(UnitUtils.ConvertToInternalUnits(z, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt21[0] = ptE[0];
            pt21[1] = ptE[1];
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt21), GetFamilySymbolByName("VBA_PLY"),
                draftingView);
            RotateFamily(family, 90);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(z, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt21[0] = ptE[0] + ply_width_e + ply_thk;
            pt21[1] = ptE[1];
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt21), GetFamilySymbolByName("VBA_PLY"),
                draftingView);
            RotateFamily(family, 90);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(z, DisplayUnitType.DUT_DECIMAL_INCHES));

            pt21[0] = ptE[0] - 2.25;
            pt21[1] = ptE[1] + z;
            CreateDimension(pt21[0] + clamp_L + 2, pt21[1],
                pt21[0] + clamp_L + 2, pt21[1] - clamp_spacing[0], 0, 0, draftingView, _doc);
            for (var q = 0; q < n_clamps; q++)
            {
                pt21[1] -= clamp_spacing[q];
                if (q % 2 != 0)
                {
                    pt19[0] = pt21[0];
                    pt19[1] = pt21[1];
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt19), GetFamilySymbolByName(clamp_block_pr),
                        draftingView);
                }
                else
                {
                    pt19[0] = pt21[0] + 4.5 + x;
                    pt19[1] = pt21[1];
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt19),
                        GetFamilySymbolByName(clamp_block_pr_flip),
                        draftingView);
                }

                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(4.5 + ply_width_e, DisplayUnitType.DUT_DECIMAL_INCHES));

                if (q != 0)
                {
                    if (q % 2 != 0)
                    {
                        CreateDimension(pt21[0] + clamp_L + 2, pt21[1],
                            pt21[0] + clamp_L + 2, pt21[1] + clamp_spacing[q], 0, 0, draftingView, _doc);
                    }
                    else
                    {
                        CreateDimension(pt21[0] + clamp_L + 2, pt21[1],
                            pt21[0] + clamp_L + 2, pt21[1] + clamp_spacing[q], 0, 0, draftingView, _doc);
                    }
                }

                pt22[0] = pt21[0] + clamp_L + 10;
                pt22[1] = ptE[1] + clamp_spacing_con[q] + 1.25;
                var clamp_str = $"{clamp_spacing_con[q]}\"";
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
                if (q == n_clamps - 1)
                {
                    if (q % 2 != 0)
                    {
                        CreateDimension(pt21[0] + clamp_L + 2, ptE[1],
                            pt21[0] + clamp_L + 2, ptE[1] + bot_clamp_gap, 0, 0, draftingView, _doc);
                    }
                    else
                    {
                        CreateDimension(pt21[0] + clamp_L + 2, ptE[1],
                            pt21[0] + clamp_L + 2, ptE[1] + bot_clamp_gap, 0, 0, draftingView, _doc);
                    }
                }
            }


            string strOrdinal;
            if (n_pours >= 2)
            {
                string iStr;
                for (int i = 0; i < n_pours - 1; i++)
                {
                    pt20[0] = ptE[0] - 24;
                    pt20[1] = ptE[1] + clamp_matrix[row_num, 0] * (i + 1);
                    pt21[0] = ptE[1] + x + 58;
                    pt21[1] = ptE[1] + clamp_matrix[row_num, 0] * (i + 1);
                    _doc.Create.NewDetailCurve(draftingView,
                        Line.CreateBound(GetXYZByPoint(pt20), GetXYZByPoint(pt21)));
                    iStr = i.ToString();
                    if (i >= 10 && i <= 18)
                    {
                        strOrdinal = "TH";
                    }
                    else if (iStr.Substring(iStr.Length - 1, 1).Equals(0))
                    {
                        strOrdinal = "ST";
                    }
                    else if (iStr.Substring(iStr.Length - 1, 1).Equals(1))
                    {
                        strOrdinal = "ND";
                    }
                    else if (iStr.Substring(iStr.Length - 1, 1).Equals(2))
                    {
                        strOrdinal = "RD";
                    }
                    else
                    {
                        strOrdinal = "TH";
                    }

                    pt20[0] = ptE[0] + x + 40;
                    pt20[1] = ptE[1] + clamp_matrix[row_num, 0] * (i + 1) + 3.25;
                    textNoteOptions = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Top,
                        HorizontalAlignment = HorizontalTextAlignment.Left,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "2.5 mm").Id
                    };
                    TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt20),
                        $"MAX HT OF \n{(i + 1).ToString()}{strOrdinal} POUR", textNoteOptions);
                    pt20[0] = ptE[0] - 32;
                    pt20[1] = ptE[1] + clamp_matrix[row_num, 0] * (i + 1) + 1.25;
                    textNoteOptions = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Top,
                        HorizontalAlignment = HorizontalTextAlignment.Left,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "2.5 mm").Id
                    };
                    TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt20),
                        $"{(i + 1) * clamp_matrix[row_num, 0]}\"", textNoteOptions);
                }
            }

            pt20[0] = ptE[0] - ply_thk - 1.5 - 3.5;
            pt20[1] = ptE[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName("VBA_2X4"),
                draftingView);
            pt20[0] = ptE[0] + ply_width_e;
            pt20[1] = ptE[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName("VBA_2X4"),
                draftingView);
            pt20[0] = ptE[0] - ply_thk - 1.5 - 6;
            pt20[1] = ptE[1] + 3;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName("VBA_2X4_SIDE"),
                draftingView);
            family.LookupParameter("Distance1")?.Set(UnitUtils.ConvertToInternalUnits(ply_width_e + ply_thk + 1.5 + 12,
                DisplayUnitType.DUT_DECIMAL_INCHES));
            pt30[0] = ptE[0] + ply_width_e + 6;
            pt30[1] = ptE[1] + 2.25;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30),
                GetFamilySymbolByName("VBA_DOWN_PLATE_NOTES_SCISSOR"),
                draftingView);
            pt8[0] = ptA[0] + ply_width_x;
            pt8[1] = ptA[1] + z + 18;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt8), GetFamilySymbolByName("VBA_PLY"),
                draftingView);
            RotateFamily(family, 180);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_DECIMAL_INCHES));
            for (var i = 0; i < n_studs_x; i++)
            {
                pt11[0] = ptA[0] + ply_width_x - 3.5 - stud_spacing_x[i];
                pt11[1] = ptA[1] + z + 18;
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt11), GetFamilySymbolByName(stud_block),
                    draftingView);
            }

            pt9[0] = ptB[0] + ply_width_y;
            pt9[1] = ptB[1] + z + 18;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt9), GetFamilySymbolByName("VBA_PLY"),
                draftingView);
            RotateFamily(family, 180);
            family.LookupParameter("Distance1")
                ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));
            for (var i = 0; i < n_studs_y; i++)
            {
                pt11[0] = ptB[0] + ply_width_y - 3.5 - stud_spacing_y[i];
                pt11[1] = ptB[1] + z + 18;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt11), GetFamilySymbolByName(stud_block),
                    draftingView);
            }

            CreateDimension(pt8[0], pt8[1] - ply_thk, pt8[0] - ply_width_x, pt8[1] - ply_thk, 0, -5, draftingView,
                _doc);
            for (var i = 0; i < n_studs_x; i++)
            {
                CreateDimension(pt8[0] - stud_spacing_x[stud_spacing_x.Length - i - 1], pt8[1],
                    pt8[0] - ply_width_x, pt8[1], 0, (i + 1) * 5, draftingView, _doc);
            }

            //TODO Line so small
            // CreateDimention(pt8[0] - stud_spacing_x[stud_spacing_x.Length - 1] - 3.5, pt8[1] + 5,
            //     pt8[0] - ply_width_x, pt8[1]+5, draftingView, _doc);
            // CreateDimention(pt8[0] - stud_start_offset_x, pt8[1] + 1.5, pt8[0], pt8[1], draftingView, _doc);

            CreateDimension(pt9[0], pt9[1] - ply_thk, pt9[0] - ply_width_y, pt9[1] - ply_thk, 0, -5, draftingView,
                _doc);
            for (var i = 0; i < n_studs_y; i++)
            {
                CreateDimension(pt9[0] - stud_spacing_x[stud_spacing_y.Length - i - 1], pt9[1],
                    pt9[0] - ply_width_y, pt9[1], 0, (i + 2) * 5, draftingView, _doc);
            }

            CreateDimension(pt9[0] - stud_spacing_y[stud_spacing_x.Length - 1] - 3.5, pt9[1],
                pt9[0] - ply_width_y, pt9[1], 0, 5, draftingView, _doc);
            CreateDimension(pt9[0] - stud_start_offset_y, pt9[1], pt9[0], pt9[1], 0, 5, draftingView, _doc);
            // family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt9),
            //     GetFamilySymbolByName("VBA_TOP_SECTION_DETAILS1"),
            //     draftingView);
            var PlyTemp = ptA[1] + ply_seams[0];
            var dimention = CreateDimension(ptA[0] + ply_width_x, ptA[1], ptA[0] + ply_width_x, PlyTemp, 7, 0,
                draftingView, _doc);
            dimention.Suffix = "PLYWOOD";
            if (ply_seams.Length >= 2)
            {
                for (int i = 1; i < ply_seams.Length; i++)
                {
                    PlyTemp += ply_seams[i];
                    dimention = CreateDimension(ptA[0] + ply_width_x, PlyTemp - ply_seams[i],
                        ptA[0] + ply_width_x, PlyTemp, 7, 0, draftingView, _doc);
                    dimention.Suffix = "PLYWOOD";
                }
            }

            dimention = CreateDimension(ptA[0], ptA[1],
                ptA[0], ptA[1] + z, -7, 0, draftingView, _doc);
            dimention.Suffix = "OVERALL HEIGHT";
            dimention = CreateDimension(ptA[0], ptA[1] + z,
                ptA[0], ptA[1] + stud_base_gap, -2, 0, draftingView, _doc);
            dimention.Suffix = "STUD";
            dimention = CreateDimension(ptA[0] + ply_width_x, ptA[1] + z,
                ptA[0], ptA[1] + z, 0, 2, draftingView, _doc);
            dimention.Suffix = "PLYWOOD";
            PlyTemp = ptB[1] + ply_seams[0];
            dimention = CreateDimension(ptB[0] + ply_width_y, ptB[1], ptB[0] + ply_width_y, PlyTemp,
                7, 0, draftingView, _doc);
            dimention.Suffix = "PLYWOOD";
            if (ply_seams.Length >= 2)
            {
                for (int i = 1; i < ply_seams.Length; i++)
                {
                    PlyTemp += ply_seams[i];
                    dimention = CreateDimension(ptB[0] + ply_width_y, PlyTemp - ply_seams[i],
                        ptB[0] + ply_width_y, PlyTemp, 7, 0, draftingView, _doc);
                    dimention.Suffix = "PLYWOOD";
                }
            }

            dimention = CreateDimension(ptB[0], ptB[1],
                ptB[0], ptB[1] + z, -7, 0, draftingView, _doc);
            dimention.Suffix = "OVERALL HEIGHT";
            dimention = CreateDimension(ptB[0], ptB[1] + z,
                ptB[0], ptB[1] + stud_base_gap, -2, 0, draftingView, _doc);
            dimention.Suffix = "STUD";
            dimention = CreateDimension(ptB[0] + ply_width_y, ptB[1] + z,
                ptB[0], ptB[1] + z, 0, 2, draftingView, _doc);
            dimention.Suffix = "PLYWOOD";
            pt13[0] = pt_o[0] + 427 + x;
            pt13[1] = pt_o[1] + 42 + ply_thk;

            pt15[0] = pt13[0] - x - 2.25;
            pt15[1] = pt13[1] - 2.25;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt15), GetFamilySymbolByName(clamp_block),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("Distance1")
                .Set(UnitUtils.ConvertToInternalUnits(x + 4.5, DisplayUnitType.DUT_DECIMAL_INCHES));
            family.LookupParameter("Distance2")
                .Set(UnitUtils.ConvertToInternalUnits(y + 4.5, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt12[0] = pt13[0] - x;
            pt12[1] = pt13[1] - ply_thk;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY"), draftingView);
            _doc.Regenerate();
            family.LookupParameter("Distance1")
                .Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt12[0] = pt12[0] - ply_thk;
            pt12[1] = pt12[1] + ply_width_y + ply_thk - 2.25;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY"), draftingView);
            RotateFamily(family, -90);
            family.LookupParameter("Distance1")
                .Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt12[0] = pt12[0] + ply_thk + ply_width_x;
            pt12[1] = pt12[1] + ply_thk - 2.25;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY"), draftingView);
            RotateFamily(family, 180);
            family.LookupParameter("Distance1")
                .Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt12[0] = pt12[0] + ply_thk;
            pt12[1] = pt12[1] - ply_thk - ply_width_y + 2.25;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt12), GetFamilySymbolByName("VBA_PLY"), draftingView);
            RotateFamily(family, 90);
            family.LookupParameter("Distance1")
                .Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_DECIMAL_INCHES));
            pt14[1] = pt13[1] - ply_thk - 1.5;
            for (int j = 0; j < n_studs_x; j++)
            {
                pt14[0] = pt13[0] - x + stud_spacing_x[j];
                _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block),
                    draftingView);
            }

            pt14[0] = pt13[0] - x + stud_spacing_x[stud_spacing_x.Length - 1];
            pt14[0] = pt13[0] - x - ply_thk - 1.5;
            for (int j = 0; j < n_studs_y; j++)
            {
                pt14[1] = pt13[1] + y + 2.25 - stud_spacing_y[j];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block),
                    draftingView);
                RotateFamily(family, -90);
            }

            pt14[1] = pt13[1] + y + ply_thk + 1.5;
            for (int j = 0; j < n_studs_x; j++)
            {
                pt14[0] = pt13[0] - stud_spacing_x[j];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block),
                    draftingView);
                RotateFamily(family, 180);
            }

            pt14[0] = pt13[0] + ply_thk + 1.5;
            for (int j = 0; j < n_studs_y; j++)
            {
                pt14[1] = pt13[1] + stud_spacing_y[j];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName(stud_block),
                    draftingView);
                RotateFamily(family, 90);
            }

            pt14[0] = pt13[0] - x / 2;
            pt14[1] = pt13[1] - 12;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName("VBA_ELEV_A"),
                draftingView);
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

            pt14[0] = pt13[0];
            pt14[1] = pt13[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName("VBA_3-4_CHAMFER"),
                draftingView);
            pt14[0] = pt13[0] - x;
            pt14[1] = pt13[1];
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName("VBA_3-4_CHAMFER"),
                draftingView);
            RotateFamily(family, -90);
            pt14[0] = pt13[0] - x;
            pt14[1] = pt13[1] + y;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName("VBA_3-4_CHAMFER"),
                draftingView);
            RotateFamily(family, 180);
            pt14[0] = pt13[0];
            pt14[1] = pt13[1] + y;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt14), GetFamilySymbolByName("VBA_3-4_CHAMFER"),
                draftingView);
            RotateFamily(family, 90);
            pt17[0] = pt13[0] - x;
            pt17[1] = pt13[1] + y - 0.5;
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
            pt17[0] = pt13[0] - x + 0.5;
            pt17[1] = pt13[1];
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
            pt17[0] = pt13[0] - x;
            pt17[1] = pt13[1] + 2.25;
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
            pt17[1] = pt13[1];
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
            pt_blk[0] = ptB[0] + (x + 4.5) / 2 - 14;
            pt_blk[1] = pt_o[1] + 315;
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk), GetFamilySymbolByName("VBA_STUD_HOLDBACK_DETAIL"),
                draftingView);
            pt_blk[0] = ptB[0] + x + 4.5;
            pt_blk[1] = ptB[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk), GetFamilySymbolByName("VBA_STUD_HOLDBACK_CALLOUT"),
                draftingView);
            pt_blk[0] = ptB[0] + x + 0.375;
            pt_blk[1] = ptB[1] + 0.95 * ply_seams[0];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt_blk), GetFamilySymbolByName("VBA_NAILING_NOTES_SCISSOR"),
                draftingView);
            pt_blk[0] = pt13[0] - ply_width_x / 2;
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
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "D", textNoteOptions);
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
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "PLAN VIEW", textNoteOptions);
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
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "C", textNoteOptions);
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
            pt_blk[0] = ptB[0] + ply_width_y / 2;
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
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "VIEWED FROM PLYWOOD FACE", textNoteOptions);
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
// ' / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / / /
// '                          C L E A N   U P
// ' \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \ \
//     ThisDrawing.ActiveDimStyle = OldDimStyle
//     ThisDrawing.ActiveTextStyle = OldTextStyle
//     ThisDrawing.ActiveLayer = OldLayer
//     
//     'Clean up objects that may be causing a lock violation
//     Set MTextObject1 = Nothing
//     Set MTextObject2 = Nothing
//     Set objTextStyle = Nothing
//     Set dimObj = Nothing
//     Set LineObj = Nothing
//     Set plineObj = Nothing
//     Set BlockRefObj = Nothing
//     Set ObjDynBlockProp = Nothing
//     Set OldDimStyle = Nothing
//     Set OldTextStyle = Nothing
//     Set OldLayer = Nothing
//     Set NewDimStyle = Nothing
//     Set NewTextStyle = Nothing
//     Set NewLayer = Nothing
//     
//     GoTo SkipErrorHandler
// ErrorHandler;
//     If Err.Number = -2147352567 Then 'This error occurs when the user presses the ESC key while inside the GetPoint utility
//         ColumnCreator.show
//         Exit Function
//     ElseIf Err.Number = -2145386422 Then 'This error occurs when the path to Egnyte isn't found
//         MsgBox ["Error; VBA block library not found at; " & FileToInsert & vbNewLine & vbNewLine & "Check Egnyte [Z;\] connection. Try closing and relaunching Egnyte Connect."]
//         Exit Function
//     ElseIf Err.Number = -2145386445 Then 'This error occurs when a block definition that doesn't exist is attempted to be inserted.
//         If BlockExists = True Then
//             MsgBox ["This drawing seems to be missing blocks for this VBA script. This is normal if you are using new button features on an old drawing. The block library will be re-downloaded and this script terminated." & vbNewLine & vbNewLine & "Please run the button again."]
//             'Insert the block library then delete it
//             FileToInsert = FileLocationPrefix & "VBA_Block_Library.dwg"
//             Set BlockDwg = ThisDrawing.ModelSpace.InsertBlock[insertionPnt, FileToInsert, 1#, 1#, 1#, 0]
//             BlockDwg.Delete
//             Exit Function
//         Else
//             MsgBox ["Error; This drawing seems to be missing blocks for this VBA script despite the block library having been downloaded. The library may be missing the necessary blocks, or something mysterious has gone horribly wrong." & vbNewLine & vbNewLine & "Library location; " & FileToInsert]
//             Exit Function
//         End If
//     Else 'Generic error message
//         MsgBox ["An error occured." & vbNewLine & "error " & Err.Number & vbNewLine & Err.Description]
//         Exit Function
//     End If
//     
// SkipErrorHandler;
//
// End Function
//
//
//
// Public Function ChangeProp[BlockRefObj, PropName, PropValue]
//
//     Props = BlockRefObj.GetDynamicBlockProperties
//     Dim oprop As AcadDynamicBlockReferenceProperty
//     
//     For i = LBound[Props] To UBound[Props]
//         Set oprop = Props[i] 'Get object properties from block one at a time
//         If oprop.PropertyName = PropName Then 'compare to property name
//             oprop.Value = PropValue 'if name matches, assign new PropValue
//         End If
//     'Exit For
//     Next i
// End Function
// Public Function DrawDim[x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double]
//
//     Dim dimObj As AcadDimAligned
//     Dim point1[0 To 2] As Double 'first point
//     Dim point2[0 To 2] As Double 'second point
//     Dim location[0 To 2] As Double 'text location
//     
//     ' Define the dimension
//     point1[0] = x1#;    point1[1] = y1#;    point1[2] = 0#
//     point2[0] = x2#;    point2[1] = y2#;    point2[2] = 0#
//     location[0] = x3#;  location[1] = y3#;  location[2] = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimAligned[point1, point2, location]
// End Function
// Public Function DrawDimSuffix[x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, txt As String]
//
//     Dim dimObj As AcadDimAligned
//     Dim point1[0 To 2] As Double
//     Dim point2[0 To 2] As Double
//     Dim location[0 To 2] As Double
//     
//     ' Define the dimension
//     point1[0] = x1#;    point1[1] = y1#;    point1[2] = 0#
//     point2[0] = x2#;    point2[1] = y2#;    point2[2] = 0#
//     location[0] = x3#;  location[1] = y3#;  location[2] = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimAligned[point1, point2, location]
//     dimObj.TextSuffix = txt
//     
// End Function
// Public Function DrawDimLin[x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, angle As Double]
//
//     Dim dimObj As AcadDimRotated
//     Dim point1[0 To 2] As Double 'first point
//     Dim point2[0 To 2] As Double 'second point
//     Dim location[0 To 2] As Double 'text location
//     
//     ' Define the dimension
//     point1[0] = x1#;    point1[1] = y1#;    point1[2] = 0#
//     point2[0] = x2#;    point2[1] = y2#;    point2[2] = 0#
//     location[0] = x3#;  location[1] = y3#;  location[2] = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimRotated[point1, point2, location, angle]
// End Function
// Public Function DrawDimLinSuffixLeader[x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, txt As String, angle As Double]
//
//     Dim dimObj As AcadDimRotated
//     Dim point1[0 To 2] As Double
//     Dim point2[0 To 2] As Double
//     Dim location[0 To 2] As Double
//     Dim txtLocation[0 To 2] As Double
//     
//     ' Define the dimension
//     point1[0] = x1#; point1[1] = y1#; point1[2] = 0#
//     point2[0] = x2#; point2[1] = y2#; point2[2] = 0#
//     location[0] = x3#; location[1] = y3#; location[2] = 0#
//     txtLocation[0] = x4#; txtLocation[1] = y4#; txtLocation[2] = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimRotated[point1, point2, location, angle]
//     dimObj.TextMovement = acMoveTextAddLeader
//     dimObj.TextPosition = txtLocation
//     dimObj.TextSuffix = txt
// End Function
// Public Function DrawDimLinLeader[x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, angle As Double]
//
//     Dim dimObj As AcadDimRotated
//     Dim point1[0 To 2] As Double
//     Dim point2[0 To 2] As Double
//     Dim location[0 To 2] As Double
//     Dim txtLocation[0 To 2] As Double
//     
//     ' Define the dimension
//     point1[0] = x1#; point1[1] = y1#; point1[2] = 0#
//     point2[0] = x2#; point2[1] = y2#; point2[2] = 0#
//     location[0] = x3#; location[1] = y3#; location[2] = 0#
//     txtLocation[0] = x4#; txtLocation[1] = y4#; txtLocation[2] = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimRotated[point1, point2, location, angle]
//     dimObj.TextMovement = acMoveTextAddLeader
//     dimObj.TextPosition = txtLocation
// End Function
// 'draws a rectangle with lower left corner at xoff, yoff with COORDINATES for far point specified
// '    4_______3B
// '    |       |
// '    |       |
// '    |_______|
// '   A1       2
//
// Public Function DrawRectangle[xA As Double, yA As Double, xB As Double, yB As Double]
//     Dim plineObj As AcadLWPolyline
//     Dim pts[0 To 7] As Double
//     
//     pts[0] = xA#;   pts[1] = yA
//     pts[2] = xB#;   pts[3] = yA
//     pts[4] = xB#;   pts[5] = yB
//     pts[6] = xA#;   pts[7] = yB
//
//     Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline[pts]
//     plineObj.Closed = True
//
// End Function
// Public Function FindMax[Arr As Variant]
//     Dim RetMax As Double
//     RetMax = Arr[LBound[Arr]] 'initialize with first value
//     
//     For i = LBound[Arr] To UBound[Arr]
//     If Arr[i] > RetMax Then
//         RetMax = Arr[i]
//     End If
//     Next i
//     FindMax = RetMax
// End Function
        }

        public static Dimension CreateDimension(double point1, double point2, double point3, double point4,
            double xOffset, double yOffset, View view, Document document)
        {
            var xyz1 = new XYZ(UnitUtils.ConvertToInternalUnits(point1, DisplayUnitType.DUT_DECIMAL_INCHES),
                UnitUtils.ConvertToInternalUnits(point2, DisplayUnitType.DUT_DECIMAL_INCHES), 0);
            var xyz2 =
                new XYZ(UnitUtils.ConvertToInternalUnits(point3, DisplayUnitType.DUT_DECIMAL_INCHES),
                    UnitUtils.ConvertToInternalUnits(point4, DisplayUnitType.DUT_DECIMAL_INCHES), 0);
            var line = Line.CreateBound(xyz1, xyz2);
            var curve = document.Create.NewDetailCurve(view, line);
            var references = new ReferenceArray();
            references.Append(curve.GeometryCurve.GetEndPointReference(0));
            references.Append(curve.GeometryCurve.GetEndPointReference(1));
            var line2 = Line.CreateBound(
                new XYZ(xyz1.X + UnitUtils.ConvertToInternalUnits(xOffset, DisplayUnitType.DUT_DECIMAL_INCHES),
                    xyz1.Y + UnitUtils.ConvertToInternalUnits(yOffset, DisplayUnitType.DUT_DECIMAL_INCHES), xyz1.Z),
                new XYZ(xyz2.X + UnitUtils.ConvertToInternalUnits(xOffset, DisplayUnitType.DUT_DECIMAL_INCHES),
                    xyz2.Y + UnitUtils.ConvertToInternalUnits(yOffset, DisplayUnitType.DUT_DECIMAL_INCHES), xyz2.Z));

            var dimensionType = new FilteredElementCollector(document)
                .OfClass(typeof(DimensionType))
                .Cast<DimensionType>()
                .First(n => n.Name.Equals("Inches 3-32"));
            return document.Create.NewDimension(view, line2, references, dimensionType);
        }

        // public static Dimension CreateDimension(double point1, double point2, double point3, double point4, View view, Document document)
        // {
        //     var xyz1 =
        //         new XYZ(UnitUtils.ConvertToInternalUnits(point1, DisplayUnitType.DUT_DECIMAL_INCHES),
        //             UnitUtils.ConvertToInternalUnits(point2, DisplayUnitType.DUT_DECIMAL_INCHES), 0);
        //     var xyz2 =
        //         new XYZ(UnitUtils.ConvertToInternalUnits(point3, DisplayUnitType.DUT_DECIMAL_INCHES),
        //             UnitUtils.ConvertToInternalUnits(point4, DisplayUnitType.DUT_DECIMAL_INCHES), 0);
        //     var line = Line.CreateBound(xyz1, xyz2);
        //     var curve= document.Create.NewDetailCurve(view, line);
        //     var references = new ReferenceArray();
        //     references.Append(curve.GeometryCurve.GetEndPointReference(0));
        //     references.Append(curve.GeometryCurve.GetEndPointReference(1));
        //     var dimensionType = new FilteredElementCollector(document)
        //         .OfClass(typeof(DimensionType))
        //         .Cast<DimensionType>()
        //         .First(n=>n.Name.Equals("Inches 3-32"));
        //     return document.Create.NewDimension(view, line, references, dimensionType);
        // }
    }
}