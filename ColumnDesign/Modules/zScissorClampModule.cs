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
            var filepath = @$"{GlobalNames.WtFileLocationPrefix}Columns\scissor_clamp_matrix.csv";
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
            int n_studs_x=0;
            int n_studs_y=0;
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
            x = ConvertToNum(_ui.WidthX.Text);
            y = ConvertToNum(_ui.LengthY.Text);
            z = ConvertToNum(_ui.HeightZ.Text);
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
                if (stud_matrix[i,0]>=x)
                {
                    n_studs_x = stud_matrix[i, 1];
                    break;
                }
            }
            for (int i = 0; i < stud_matrix.GetLength(0); i++)
            {
                if (stud_matrix[i,0]>=y)
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
                stud_spacing_x[i] = stud_start_offset_x + (i) * (ply_width_x - 2 * stud_start_offset_x - 3.5) / (n_studs_x - 1);
            }

            for (var i = 0; i < n_studs_y; i++)
            {
                stud_spacing_y[i] = stud_start_offset_y + (i) * (ply_width_y - 2 * stud_start_offset_y - 3.5) / (n_studs_y - 1);
            }
            // '#####################################################################################
            // '                                   P L Y W O O D
            // '#####################################################################################
            double ply_top_ht_min = 6;
            double max_ply_ht;
            var ply_seams = new double[1];
            ply_seams = ReadSizes(_ui.SPlywoodType.Text);
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
                        else if (l == ply_cuts.GetLength(1))
                        {
                            unique_plys++;
                            ply_cuts = ResizeArray<double>(ply_cuts, 3, unique_plys);
                            ply_cuts[0, unique_plys] = ply_widths[j];
                            ply_cuts[1, unique_plys] = ply_seams[i];
                            ply_cuts[2, unique_plys] += 2;
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
                if ((clamp_matrix[i, 0] >= long_side))
                {
                    row_num = i;
                    break;
                }
            }

            n_pours = (int) (z / clamp_matrix[row_num, 0]);
            if (n_pours*clamp_matrix[row_num, 0]<z)
            {
                n_pours++;
            }

            var z_original = z;
            var pour_overlap = 0d;
            var multi_clamp_spacing = new double[100, n_pours];
            var total_clamp_spacings = new int[n_pours];
            for (int i = 0; i < n_pours; i++)
            {
                if (n_pours>=2)
                {
                    if (z_original-(i+1)*clamp_matrix[row_num,0]>=0)
                    {
                        z = clamp_matrix[row_num, 0];
                    }
                    else
                    {
                        z = z_original - i * clamp_matrix[row_num, 0];
                        if (z<12)
                        {
                            pour_overlap = 12 - z;
                            z = 12;
                        }
                    }
                }
                Array.Resize(ref clamp_spacing, clamp_matrix.GetLength(1));
                var k = 0;
                for (int j = 1; j < clamp_matrix.GetLength(1); j++)
                {
                    if (clamp_matrix[row_num, j]==0)
                    {
                        Array.Resize(ref clamp_spacing, k);
                        break;
                    }
                    else
                    {
                        clamp_spacing[j - 1] = clamp_matrix[row_num, j];
                        k++;
                    }
                }

                var ht_rem = z - bot_clamp_gap;
                var n_clamps_temp = 0;
                for (int j = 0; j < clamp_spacing.Length-1; j++)
                {
                    if (clamp_spacing[j]<=ht_rem)
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
                    tot_temp += (int)clamp_spacing[j];
                }

                clamp_spacing[clamp_spacing.Length - 2] -= tot_temp - z;
                TestClampSpacing:
                var infExit = 0;
                for (var j = 0; i < clamp_spacing.Length; i++)
                {
                    if (clamp_spacing[j] < 8)
                    {
                        clamp_spacing[j - 1] -= 8 - clamp_spacing[j];
                        clamp_spacing[j] = 8;


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

                for (int j = 0; j < clamp_spacing.Length; j++)
                {
                    multi_clamp_spacing[j, i] = clamp_spacing[j];
                }

                total_clamp_spacings[i] = clamp_spacing.Length;
                n_clamps += n_clamps_temp;
                if (i>=2)
                {
                    n_clamps++;
                }
            }

            var clamp_spacing_r = new int[100];
            Array.Resize(ref clamp_spacing, 100);
            var jCount = 1;
            for (int p = n_pours-1; p >=0; p--)
            {
                for (int i = 0; i < total_clamp_spacings[p]; i++)
                {
                    if (multi_clamp_spacing[i,p] !=0)       
                    {
                        if ((p!=(n_pours-1)) && i==0)
                        {
                            clamp_spacing[jCount - 1] = multi_clamp_spacing[i, p];
                            n_clamps++;
                        }
                        else
                        {
                            clamp_spacing[i] = multi_clamp_spacing[i, p];
                            jCount++;
                        }
                    }
                }
            }

            if (pour_overlap!=0)
            {
                clamp_spacing[total_clamp_spacings[n_pours - 1]] += pour_overlap;
                clamp_spacing[total_clamp_spacings[n_pours - 1]+1] -= pour_overlap;
            }

            for (int i = 0; i < clamp_spacing.Length; i++)
            {
                if (clamp_spacing[i]==0 && clamp_spacing[i+1]==0)
                {
                    Array.Resize(ref clamp_spacing, i-1);
                    break;  
                }
            }

            if (clamp_spacing[0]<2 && clamp_spacing[0]>=0)
            {
                if (clamp_spacing[1] >= (2-clamp_spacing[0]) + 5)
                {
                    clamp_spacing[1] -= 2 - clamp_spacing[0];
                    clamp_spacing[0] = 2;
                }
                else if (clamp_spacing[2] >= (2-clamp_spacing[0]) + 5)
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
            var strFabNotes = $"• COLUMN SIZE = {x}' X + {y}'\n\n" +
                              $"• NUMBER OF COLUMN FORMS = {n_col}-EA\n\n" +
                              $"• COLUMN FORM WEIGHT (APPROXIMATE) = {col_wt}-LBS\n" +
                              $"• COLUMN PANEL WEIGHT (SINGLE PANEL) = {panel_wt}-LBS\n\n" +
                              $"• PLYWOOD = 3/4'' PLYFORM (\"{ply_name}\"), CLASS-1 (MIN)\n\n" +
                              "• COLUMN FORMS AND CLAMP SPACING LAYOUTS FOR SCISSOR CLAMPS ARE DESIGNED FOR A POUR RATE = ";
            if (z<=clamp_matrix[row_num,0])
            {
                strFabNotes += "FULL LIQUID HEAD U.N.O.\n\n";
            }
            else
            {
                strFabNotes += $"{ConvertFtIn(clamp_matrix[row_num,0])}\n\n";
            }
            strFabNotes += "• CONTACT THE MCC ENGINEER PRIOR TO ANY CHANGES OR MODIFICATIONS TO THE DETAILS ON THIS SHEET.";
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
                UnitUtils.ConvertToInternalUnits(100, DisplayUnitType.DUT_MILLIMETERS), strFabNotes, textNoteOptions);
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
            // TODO n_chamf value ??
            // qty_text += "3/4'' GATES PLASTIC CHAMFER (BASED ON 12' CHAMFER LENGTHS)\n" +
            //             $"•   ({n_col * n_chamf}-EA) = ({n_col}-COL) X ({n_chamf}-EA/COL) @ {chamf_length}'-0'' LONG PIECES";
            textNoteOptions = new TextNoteOptions
            {
                VerticalAlignment = VerticalTextAlignment.Top,
                HorizontalAlignment = HorizontalTextAlignment.Left,
                TypeId = new FilteredElementCollector(_doc)
                    .OfClass(typeof(TextNoteType))
                    .Cast<TextNoteType>()
                    .First(q => q.Name == "2.0 mm").Id
            };
            TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt1),
                UnitUtils.ConvertToInternalUnits(100, DisplayUnitType.DUT_MILLIMETERS), qty_text, textNoteOptions);

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
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(new double[] {ptE[0] - ply_thk-1.5, ptE[1]+stud_base_gap}),
                GetFamilySymbolByName("VBA_RECTANGLE"),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("xB")
                .Set(UnitUtils.ConvertToInternalUnits(1.5, DisplayUnitType.DUT_MILLIMETERS));
            family.LookupParameter("yB").Set(UnitUtils.ConvertToInternalUnits(z, DisplayUnitType.DUT_MILLIMETERS));

            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(new double[] {ptE[0] +ply_width_e+ply_thk, ptE[1]+stud_base_gap}),
                GetFamilySymbolByName("VBA_RECTANGLE"),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("xB")
                .Set(UnitUtils.ConvertToInternalUnits(1.5, DisplayUnitType.DUT_MILLIMETERS));
            family.LookupParameter("yB").Set(UnitUtils.ConvertToInternalUnits(z, DisplayUnitType.DUT_MILLIMETERS));
            for (var i = 0; i < n_studs_x; i++)
            {
                pt6[0] = ptA[0] + ply_width_x  - 3.5 - stud_spacing_x[i];
                pt6[1] = ptA[1] + stud_base_gap;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt6), GetFamilySymbolByName(stud_face_block),
                    draftingView);
                RotateFamily(family, 90);
                family.LookupParameter("Distance1")?.Set(UnitUtils.ConvertToInternalUnits(z - stud_base_gap, DisplayUnitType.DUT_MILLIMETERS));
                textNoteOptions   = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Middle,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    Rotation = Math.PI/2,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "2.0 mm").Id
                };
                var ptTemp = new double[2];
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
                family.LookupParameter("Distance1")?.Set(UnitUtils.ConvertToInternalUnits(z - stud_base_gap, DisplayUnitType.DUT_MILLIMETERS));
                textNoteOptions   = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Middle,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    Rotation = Math.PI/2,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "2.0 mm").Id
                };
                var ptTemp = new double[2];
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
                    ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_x, DisplayUnitType.DUT_MILLIMETERS));
                family.LookupParameter("Distance2")
                    ?.Set(UnitUtils.ConvertToInternalUnits(t, DisplayUnitType.DUT_MILLIMETERS));
                
                pt4[0] = ptB[0];
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt4), GetFamilySymbolByName("VBA_PLY_SHEET"),
                    draftingView);
                family.LookupParameter("Distance1")
                    ?.Set(UnitUtils.ConvertToInternalUnits(ply_width_y, DisplayUnitType.DUT_MILLIMETERS));
                family.LookupParameter("Distance2")
                    ?.Set(UnitUtils.ConvertToInternalUnits(t, DisplayUnitType.DUT_MILLIMETERS));
            }
            for (var i = 0; i < n_studs_e; i++)
            {
                pt20[0] = ptE[0] + stud_spacing_e[i];
                pt20[1] = ptE[1] + stud_base_gap;
                family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName(stud_face_block),
                    draftingView);
                RotateFamily(family, 90);
                family.LookupParameter("Distance1")?.Set(UnitUtils.ConvertToInternalUnits( z - stud_base_gap, DisplayUnitType.DUT_MILLIMETERS));
                family.LookupParameter("Solid")?.Set(1);
                textNoteOptions   = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Middle,
                    HorizontalAlignment = HorizontalTextAlignment.Center,
                    Rotation = Math.PI/2,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "2.0 mm").Id
                };
                var ptTemp = new double[2];
                ptTemp[0] = pt20[0] + 1.5;
                ptTemp[1] = pt20[1] + (z - stud_base_gap) / 2;
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(ptTemp), "2x4", textNoteOptions);
            }
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(new double[] {ptE[0], ptE[1]}),
                GetFamilySymbolByName("VBA_RECTANGLE"),
                draftingView);
            _doc.Regenerate();
            family.LookupParameter("xB")
                .Set(UnitUtils.ConvertToInternalUnits(ply_width_e, DisplayUnitType.DUT_MILLIMETERS));
            family.LookupParameter("yB").Set(UnitUtils.ConvertToInternalUnits(z, DisplayUnitType.DUT_MILLIMETERS));
            pt21[0] = ptE[1];
            pt21[1] = ptE[1];
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt21), GetFamilySymbolByName("VBA_PLY"),
                draftingView);
            RotateFamily(family, 90);
            family.LookupParameter("Distance1")?.Set(UnitUtils.ConvertToInternalUnits( z, DisplayUnitType.DUT_MILLIMETERS));
            pt21[0] = ptE[0] + ply_width_e + ply_thk;
            pt21[1] = ptE[1];
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt21), GetFamilySymbolByName("VBA_PLY"),
                draftingView);
            RotateFamily(family, 90);
            family.LookupParameter("Distance1")?.Set(UnitUtils.ConvertToInternalUnits( z, DisplayUnitType.DUT_MILLIMETERS));

            pt21[0] = ptE[0] -2.25;
            pt21[1] = ptE[1] + z;
            // TODO    Call DrawDimLin(pt21(0) + x + 4.5 + clamp_width, pt21(1) - clamp_spacing(1), pt21(0) + ply_width_e + 3.75 + ply_thk, pt21(1), pt21(0) + clamp_L + 5, (pt21(1) * 2 - clamp_spacing(1)) / 2, pi / 2)
            for (int i = 0; i < n_clamps; i++)
            {
                pt21[1] -= clamp_spacing[i];
                if (i % 2 == 0)
                {
                    pt19[0] = pt21[0];
                    pt19[1] = pt21[1];
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt19), GetFamilySymbolByName(clamp_block_pr),
                        draftingView);
                }
                else
                {
                    pt19[0] = pt21[0] +4.5+x;
                    pt19[1] = pt21[1]; 
                    family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt19), GetFamilySymbolByName(clamp_block_pr_flip),
                        draftingView);
                }
                family.LookupParameter("Distance1")?.Set(UnitUtils.ConvertToInternalUnits(4.5 + ply_width_e, DisplayUnitType.DUT_MILLIMETERS));
            }

            var fakeI = 0;
            for (fakeI = 0; fakeI < n_clamps; fakeI++)
            {
                pt21[1] -= clamp_spacing[fakeI];
                if (fakeI !=1)
                {
                    if (fakeI %2 ==0)
                    {
                        // TODO                Call DrawDimLin(pt21(0) + clamp_L, pt21(1), pt21(0) + x + 4.5 + clamp_width, pt21(1) + clamp_spacing(i), pt21(0) + clamp_L + 5, (pt21(1) * 2 + clamp_spacing(i)) / 2, -pi / 2)
                    }
                    else
                    {
                        //  TODO               Call DrawDimLin(pt21(0) + x + 4.5 + clamp_width, pt21(1), pt21(0) + clamp_L, pt21(1) + clamp_spacing(i), pt21(0) + clamp_L + 5, (pt21(1) * 2 + clamp_spacing(i)) / 2, -pi / 2)
                    }
                }
                pt22[0] = pt21[0] - clamp_L + 7;
                pt22[1] = ptE[1] + clamp_spacing_con[fakeI] + 1.25;
                var clamp_str = $"{clamp_spacing_con[fakeI]}\" {ConvertFtIn(clamp_spacing_con[fakeI])}";
                textNoteOptions   = new TextNoteOptions
                {
                    VerticalAlignment = VerticalTextAlignment.Top,
                    HorizontalAlignment = HorizontalTextAlignment.Left,
                    TypeId = new FilteredElementCollector(_doc)
                        .OfClass(typeof(TextNoteType))
                        .Cast<TextNoteType>()
                        .First(q => q.Name == "2.5 mm").Id
                };
                TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt22), clamp_str, textNoteOptions);
            }

            if (fakeI % 2 ==0)
            {
                //     TODO    Call DrawDimLin(pt21(0) + x + 4.5 + clamp_width, ptE(1) + bot_clamp_gap, pt21(0) + ply_width_e + 2.25 + ply_thk + 2.75, ptE(1), pt21(0) + clamp_L + 5, (2 * ptE(1) + bot_clamp_gap) / 2, pi / 2)
            }
            else
            {
                //  TODO       Call DrawDimLin(pt21(0) + clamp_L, ptE(1) + bot_clamp_gap, pt21(0) + ply_width_e + 2.25 + ply_thk + 2.75, ptE(1), pt21(0) + clamp_L + 5, (2 * ptE(1) + bot_clamp_gap) / 2, pi / 2)
            }

            string strOrdinal;
            if (n_pours>=2)
            {
                string iStr;
                for (int i = 0; i < n_pours-1; i++)
                {
                    pt20[0] = ptE[0] - 24;
                    pt20[1] = ptE[1] - clamp_matrix[row_num,0]*(i+1);
                    pt21[1] = ptE[1] + x +58;
                    pt21[1] = ptE[1] - clamp_matrix[row_num,0]*(i+1);
                    _doc.Create.NewDetailCurve(draftingView, Line.CreateBound(GetXYZByPoint(pt20), GetXYZByPoint(pt21)));
                    iStr = i.ToString();
                    if (i>=10 && i<=18)
                    {
                        strOrdinal = "TH";
                    }
                    else if (iStr.Substring(iStr.Length-1, 1).Equals(0))
                    {
                        strOrdinal = "ST";
                    }
                    else if (iStr.Substring(iStr.Length-1, 1).Equals(1))
                    {
                        strOrdinal = "ND";
                    }
                    else if (iStr.Substring(iStr.Length-1, 1).Equals(2))
                    {
                        strOrdinal = "RD";
                    }
                    else
                    {
                        strOrdinal = "TH";
                    }
                    pt20[0] = ptE[0] +x+40;
                    pt20[1] = ptE[1] + clamp_matrix[row_num,0]*(i+1)+3.25;
                    textNoteOptions   = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Top,
                        HorizontalAlignment = HorizontalTextAlignment.Left,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "2.5 mm").Id
                    };
                    TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt20), $"MAX HT OF \n{(i+1).ToString()}{strOrdinal} POUR", textNoteOptions);
                    pt20[0] = ptE[0] -32;
                    pt20[1] = ptE[1] + clamp_matrix[row_num,0]*(i+1)+1.25;
                    textNoteOptions   = new TextNoteOptions
                    {
                        VerticalAlignment = VerticalTextAlignment.Top,
                        HorizontalAlignment = HorizontalTextAlignment.Left,
                        TypeId = new FilteredElementCollector(_doc)
                            .OfClass(typeof(TextNoteType))
                            .Cast<TextNoteType>()
                            .First(q => q.Name == "2.5 mm").Id
                    };
                    TextNote.Create(_doc, draftingView.Id, GetXYZByPoint(pt20), $"{(i+1)*clamp_matrix[row_num,0]}\"", textNoteOptions);
                }
            }
            pt20[0] = ptE[0] - ply_thk - 1.5-3.5;
            pt20[1] = ptE[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName("VBA_2X4"),
                draftingView);
            pt20[0] = ptE[0] + ply_width_e;
            pt20[1] = ptE[1];
            _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName("VBA_2X4"),
                draftingView);
            pt20[0] = ptE[0] -ply_thk-1.5-6;
            pt20[1] = ptE[1]+3;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt20), GetFamilySymbolByName("VBA_2X4_SIDE"),
                draftingView);
            family.LookupParameter("Distance1")?.Set(UnitUtils.ConvertToInternalUnits(ply_width_e + ply_thk + 1.5 + 12, DisplayUnitType.DUT_MILLIMETERS));
            pt30[0] = ptE[0] + ply_width_e+6;
            pt30[1] = ptE[1]+2.25;
            family = _doc.Create.NewFamilyInstance(GetXYZByPoint(pt30), GetFamilySymbolByName("VBA_DOWN_PLATE_NOTES_SCISSOR"),
                draftingView);
        }
    }
}
//     pt8(0) = ptA(0) + ply_width_x: pt8(1) = ptA(1) + z + 18
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt8, "VBA_PLY", 1#, 1#, 1#, pi) 'Insert ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
//     For i = 1 To n_studs_x
//         pt11(0) = ptA(0) + ply_width_x - 3.5 - stud_spacing_x(i)
//         pt11(1) = ptA(1) + z + 18
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt11, stud_block, 1#, 1#, 1#, 0) 'Insert studs
//     Next i
//     
//     'Side B
//     pt9(0) = ptB(0) + ply_width_y: pt9(1) = ptB(1) + z + 18
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt9, "VBA_PLY", 1#, 1#, 1#, pi) 'Insert ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
//     For i = 1 To n_studs_y
//         pt11(0) = ptB(0) + ply_width_y - 3.5 - stud_spacing_y(i)
//         pt11(1) = ptB(1) + z + 18
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt11, stud_block, 1#, 1#, 1#, 0) 'Insert studs
//     Next i
//
//     'Dim sections
//         'Side A
//         Call DrawDim(pt8(0), pt8(1) - ply_thk, pt8(0) - ply_width_x, pt8(1) - ply_thk, (pt8(0) - ply_width_x - pt8(0)) / 2, pt8(1) - ply_thk - 5) 'Dimension plywood
//         For i = 1 To n_studs_x
//             Call DrawDimLin(pt8(0) - stud_spacing_x(UBound(stud_spacing_x) - i + 1), pt8(1) + 1.5, pt8(0) - ply_width_x, pt8(1), ((pt8(0) - ply_width_x) - (pt8(0) - stud_spacing_x(UBound(stud_spacing_x) - i + 1))) / 2, pt8(1) + 4 + i * 4, 0)   'Dimension studs from rightmost (in plan) stud
//         Next i
//         Call DrawDimLin(pt8(0) - stud_spacing_x(UBound(stud_spacing_x)) - 3.5, pt8(1) + 1.5, pt8(0) - ply_width_x, pt8(1), ((pt8(0) - ply_width_x) - (pt8(0) - stud_spacing_x(UBound(stud_spacing_x)) - 3.5)) / 2, pt8(1) + 4, 0)   'Dimension leftmost (in plan) stud to face of ply
//         Call DrawDimLin(pt8(0) - stud_start_offset_x, pt8(1) + 1.5, pt8(0), pt8(1), ((pt8(0)) - (pt8(0) - stud_start_offset_x)) / 2, pt8(1) + 4 * (n_studs_x + 1), 0)    'dimension stud_start_gap
//         
//         'Side B
//         Call DrawDim(pt9(0), pt9(1) - ply_thk, pt9(0) - ply_width_y, pt9(1) - ply_thk, (pt9(0) - ply_width_y - pt9(0)) / 2, pt9(1) - ply_thk - 5) 'Dimension plywood
//         For i = 1 To n_studs_y
//             Call DrawDimLin(pt9(0) - stud_spacing_y(UBound(stud_spacing_y) - i + 1), pt9(1) + 1.5, pt9(0) - ply_width_y, pt9(1), ((pt9(0) - ply_width_y) - (pt9(0) - stud_spacing_y(UBound(stud_spacing_y) - i + 1))) / 2, pt9(1) + 4 + i * 4, 0)   'Dimension studs from rightmost (in plan) stud
//         Next i
//         Call DrawDimLin(pt9(0) - stud_spacing_y(UBound(stud_spacing_y)) - 3.5, pt9(1) + 1.5, pt9(0) - ply_width_y, pt9(1), ((pt9(0) - ply_width_y) - (pt9(0) - stud_spacing_y(UBound(stud_spacing_y)) - 3.5)) / 2, pt9(1) + 4, 0)   'Dimension leftmost (in plan) stud to face of ply
//         Call DrawDimLin(pt9(0) - stud_start_offset_y, pt9(1) + 1.5, pt9(0), pt9(1), ((pt9(0)) - (pt9(0) - stud_start_offset_y)) / 2, pt9(1) + 4 * (n_studs_y + 1), 0)    'dimension stud_start_gap
//
//         'Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt9, "VBA_TOP_SECTION_DETAILS1", 1#, 1#, 1#, 0) 'Insert text notes, blow up detail at top section
//     
//     'Dim plywood
//         'Side A
//         PlyTemp = ptA(1) + ply_seams(1)
//         Call DrawDimSuffix(ptA(0) + ply_width_x, ptA(1), ptA(0) + ply_width_x, PlyTemp, ptA(0) + ply_width_x + 6, (PlyTemp - ptA(1)) / 2, " PLYWOOD")
//         If UBound(ply_seams) >= 2 Then 'If more than one ply panel, draw the rest, bottom up
//             For i = 2 To UBound(ply_seams)
//                 PlyTemp = PlyTemp + ply_seams(i)
//                 Call DrawDimSuffix(ptA(0) + ply_width_x, PlyTemp - ply_seams(i), ptA(0) + ply_width_x, PlyTemp, ptA(0) + ply_width_x + 6, (PlyTemp - (PlyTemp - ply_seams(i))) / 2, " PLYWOOD")
//             Next i
//         End If
//         Call DrawDimSuffix(ptA(0), ptA(1), ptA(0), ptA(1) + z, ptA(0) - 9, (2 * ptA(1) + z) / 2, " OVERALL HEIGHT")
//         Call DrawDimSuffix(ptA(0), ptA(1) + stud_base_gap, ptA(0), ptA(1) + z, ptA(0) - 4.5, ((ptA(1) + z) + (ptA(1) + stud_base_gap)) / 2, " STUD")
//         Call DrawDimSuffix(ptA(0) + ply_width_x, ptA(1) + z, ptA(0), ptA(1) + z, (2 * ptA(0) + ply_width_x) / 2, ptA(1) + z + 4, " PLYWOOD")
//         
//         'Side B
//         PlyTemp = ptB(1) + ply_seams(1)
//         Call DrawDimSuffix(ptB(0) + ply_width_y, ptB(1), ptB(0) + ply_width_y, PlyTemp, ptB(0) + ply_width_y + 6, (PlyTemp - ptB(1)) / 2, " PLYWOOD")
//         If UBound(ply_seams) >= 2 Then 'If more than one ply panel, draw the rest, bottom up
//             For i = 2 To UBound(ply_seams)
//                 PlyTemp = PlyTemp + ply_seams(i)
//                 Call DrawDimSuffix(ptB(0) + ply_width_y, PlyTemp - ply_seams(i), ptB(0) + ply_width_y, PlyTemp, ptB(0) + ply_width_y + 6, (PlyTemp - (PlyTemp - ply_seams(i))) / 2, " PLYWOOD")
//             Next i
//         End If
//         Call DrawDimSuffix(ptB(0), ptB(1), ptB(0), ptB(1) + z, ptB(0) - 9, (2 * ptB(1) + z) / 2, " OVERALL HEIGHT")
//         Call DrawDimSuffix(ptB(0), ptB(1) + stud_base_gap, ptB(0), ptB(1) + z, ptB(0) - 4.5, ((ptB(1) + z) + (ptB(1) + stud_base_gap)) / 2, " STUD")
//         Call DrawDimSuffix(ptB(0) + ply_width_y, ptB(1) + z, ptB(0), ptB(1) + z, (2 * ptB(0) + ply_width_y) / 2, ptB(1) + z + 4, " PLYWOOD")
//
//
//
//             
//     'PLAN VIEWS
//     pt13(0) = pt_o(0) + 427 + x - chamf_thk: pt13(1) = pt_o(1) + 42 + ply_thk 'origin point for plan views (lower right corner)
//
//     'Draw clamps
//     pt15(0) = pt13(0) - x - 2.25: pt15(1) = pt13(1) - 2.25
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, clamp_block, 1#, 1#, 1#, 0) 'Insert clamp (operator typical clamps)
//     Call ChangeProp(BlockRefObj, "Distance1", x + 4.5)
//     Call ChangeProp(BlockRefObj, "Distance2", y + 4.5)
//     
//     'Draw plywood
//     pt12(0) = pt13(0) - x: pt12(1) = pt13(1) - ply_thk
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY", 1#, 1#, 1#, 0) 'Bottom ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
//     
//     pt12(0) = pt12(0) - ply_thk: pt12(1) = pt12(1) + ply_width_y + ply_thk - 2.25
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY", 1#, 1#, 1#, -pi / 2) 'Left ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
//     
//     pt12(0) = pt12(0) + ply_thk + ply_width_x: pt12(1) = pt12(1) + ply_thk - 2.25
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY", 1#, 1#, 1#, pi) 'Top ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
//     
//     pt12(0) = pt12(0) + ply_thk: pt12(1) = pt12(1) - ply_thk - ply_width_y + 2.25
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY", 1#, 1#, 1#, pi / 2) 'Right ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
//
//
//     'Draw studs
//     pt14(1) = pt13(1) - ply_thk - 1.5
//     For i = 1 To n_studs_x
//         pt14(0) = pt13(0) - x + stud_spacing_x(i)
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block, 1#, 1#, 1#, 0) 'Insert bottom studs
//     Next i
//     pt14(0) = pt13(0) - x + stud_spacing_x(UBound(stud_spacing_x))
//
//     pt14(0) = pt13(0) - x - ply_thk - 1.5
//     For i = 1 To n_studs_y
//         pt14(1) = pt13(1) + y + 2.25 - stud_spacing_y(i)
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block, 1#, 1#, 1#, -pi / 2) 'Insert left studs
//     Next i
//     
//     pt14(1) = pt13(1) + y + ply_thk + 1.5
//     For i = 1 To n_studs_x
//         pt14(0) = pt13(0) - stud_spacing_x(i)
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block, 1#, 1#, 1#, pi) 'Insert top studs
//     Next i
//     
//     pt14(0) = pt13(0) + ply_thk + 1.5
//     For i = 1 To n_studs_y
//         pt14(1) = pt13(1) - 2.25 + stud_spacing_y(i)
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block, 1#, 1#, 1#, pi / 2) 'Insert right studs
//     Next i
//
//     'Draw elevation tag
//     pt14(0) = pt13(0) - x / 2: pt14(1) = pt13(1) - 12
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, "VBA_ELEV_A", 1#, 1#, 1#, 0)
//     
//     'Draw hatch
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("CONC")
//     Dim hatchObj As AcadHatch
//     Dim OuterLoop(0) As AcadEntity
//     Dim pl_pts(0 To 9) As Double
//     
//     'Set boundry of hatch (must close hatch by returning to initial point)
//     pl_pts(0) = pt13(0):     pl_pts(1) = pt13(1)
//     pl_pts(2) = pt13(0) - x: pl_pts(3) = pt13(1)
//     pl_pts(4) = pt13(0) - x: pl_pts(5) = pt13(1) + y
//     pl_pts(6) = pt13(0):     pl_pts(7) = pt13(1) + y
//     pl_pts(8) = pt13(0):     pl_pts(9) = pt13(1)
//     
//     Set OuterLoop(0) = ThisDrawing.ModelSpace.AddLightWeightPolyline(pl_pts) 'Assign the pline object as the hatch's outer loop (inner loop optional and skipped here)
//     Set hatchObj = ThisDrawing.ModelSpace.AddHatch(0, "AR-CONC", True) 'Set hatch
//     hatchObj.AppendOuterLoop (OuterLoop)
//     OuterLoop(0).Delete
//     
//     'Draw chamfer in each corner
//     pt14(0) = pt13(0): pt14(1) = pt13(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, "VBA_3-4_CHAMFER", 1#, 1#, 1#, 0) 'Lower right
//
//     pt14(0) = pt13(0) - x: pt14(1) = pt13(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, "VBA_3-4_CHAMFER", 1#, 1#, 1#, -pi / 2) 'Lower left
//     
//     pt14(0) = pt13(0) - x: pt14(1) = pt13(1) + y
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, "VBA_3-4_CHAMFER", 1#, 1#, 1#, pi) 'Upper left
//     
//     pt14(0) = pt13(0): pt14(1) = pt13(1) + y
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, "VBA_3-4_CHAMFER", 1#, 1#, 1#, pi / 2) 'Upper right
//     
//     
//     'Draw text in plan view
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
//     pt17(0) = pt13(0) - x: pt17(1) = pt13(1) + y - 0.5
//     Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, x, "SIDE A")
//     MTextObject1.StyleName = "Arial Narrow": MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75 * Abs(Round(x, 3) >= 12) + (1.75 - 0.125 * (12 - x)) * Abs(Round(x, 3) < 12) 'Uses a text size of 1.75 when the x side of the column is >= 12", text size is reduced by 0.125" for every 1" the x side is below 12"
//
//     pt17(0) = pt13(0) - x + 0.5: pt17(1) = pt13(1)
//     Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, y, "SIDE B")
//     MTextObject1.StyleName = "Arial Narrow": MTextObject1.Rotation = pi / 2: MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75 * Abs(Round(y, 3) >= 12) + (1.75 - 0.125 * (12 - y)) * Abs(Round(y, 3) < 12) 'Uses a text size of 1.75 when the y side of the column is >= 12", text size is reduced by 0.125" for every 1" the y side is below 12"
//     
//     pt17(0) = pt13(0) - x: pt17(1) = pt13(1) + 2.25
//     Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, x, x)
//     MTextObject1.StyleName = "Arial Narrow": MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75 * Abs(Round(x, 3) >= 12) + (1.75 - 0.125 * (12 - x)) * Abs(Round(x, 3) < 12) 'Uses a text size of 1.75 when the x side of the column is >= 12", text size is reduced by 0.125" for every 1" the x side is below 12"
//     
//     pt17(0) = pt13(0) - 2.25: pt17(1) = pt13(1)
//     Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, y, y)
//     MTextObject1.StyleName = "Arial Narrow": MTextObject1.Rotation = pi / 2: MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75 * Abs(Round(y, 3) >= 12) + (1.75 - 0.125 * (12 - y)) * Abs(Round(y, 3) < 12) 'Uses a text size of 1.75 when the y side of the column is >= 12", text size is reduced by 0.125" for every 1" the y side is below 12"
//
//
//     'Add stud holdback detail and callout
//     'Detail on top of sheet
//     pt_blk(0) = ptB(0) + (x + 4.5) / 2 - 14: pt_blk(1) = pt_o(1) + 315
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_STUD_HOLDBACK_DETAIL", 1, 1, 1, 0)
//     
//     'Callout on panel
//     pt_blk(0) = ptB(0) + x + 4.5: pt_blk(1) = ptB(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_STUD_HOLDBACK_CALLOUT", 1, 1, 1, 0)
//     
//     'Add nailing note
//     pt_blk(0) = ptB(0) + x + 0.375: pt_blk(1) = ptB(1) + 0.95 * ply_seams(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_NAILING_NOTES_SCISSOR", 1, 1, 1, 0)
//     
//     'Add detail references to each elevation view
//     'Plan view detail reference
//     pt_blk(0) = pt13(0) - ply_width_x / 2: pt_blk(1) = pt_o(1) + 9
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
//     AttList = BlockRefObj.GetAttributes
//     For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
//         If AttList(i).TextString = "X" Then
//             AttList(i).TextString = "D"
//         End If
//         If AttList(i).TextString = "DESCRIPTION" Then
//             AttList(i).TextString = "PLAN VIEW"
//         End If
//         If AttList(i).TextString = "REFERENCE" Then
//             AttList(i).TextString = ""
//         End If
//     Next i
//     
//     'Side A detail reference
//     pt_blk(0) = ptA(0) + ply_width_x / 2: pt_blk(1) = pt_o(1) + 9
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
//     AttList = BlockRefObj.GetAttributes
//     For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
//         If AttList(i).TextString = "X" Then
//             AttList(i).TextString = "C"
//         End If
//         If AttList(i).TextString = "DESCRIPTION" Then
//             AttList(i).TextString = "SIDE ""A"" PANEL"
//         End If
//         If AttList(i).TextString = "REFERENCE" Then
//             AttList(i).TextString = "VIEWED FROM PLYWOOD FACE"
//         End If
//     Next i
//     
//     'Side B detail reference
//     pt_blk(0) = ptB(0) + ply_width_y / 2: pt_blk(1) = pt_o(1) + 9
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
//     AttList = BlockRefObj.GetAttributes
//     For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
//         If AttList(i).TextString = "X" Then
//             AttList(i).TextString = "B"
//         End If
//         If AttList(i).TextString = "DESCRIPTION" Then
//             AttList(i).TextString = "SIDE ""B"" PANEL"
//         End If
//         If AttList(i).TextString = "REFERENCE" Then
//             AttList(i).TextString = "VIEWED FROM PLYWOOD FACE"
//         End If
//     Next i
//         
//     'Elevation view detail reference
//     pt_blk(0) = ptE(0) + ply_width_e / 2: pt_blk(1) = pt_o(1) + 9
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
//     AttList = BlockRefObj.GetAttributes
//     For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
//         If AttList(i).TextString = "X" Then
//             AttList(i).TextString = "A"
//         End If
//         If AttList(i).TextString = "DESCRIPTION" Then
//             AttList(i).TextString = "COLUMN FORM ELEVATION"
//         End If
//         If AttList(i).TextString = "REFERENCE" Then
//             AttList(i).TextString = ""
//         End If
//     Next i
//     
//     'Add lower left fab count below "column form elevation"
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
//     count_str = "FAB " & n_col & "-EA"
//     pt32(0) = ptE(0) + ply_width_e / 2: pt32(1) = pt_o(1) + 7
//     Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt32, 30, count_str)
//     MTextObject2.Height = 2.5
//
//
//     
// '#####################################################################################
// '                                S H E E T I N G
// '#####################################################################################
//    
//     Dim ptVP1(0 To 2) As Double
//     Dim ptVP2(0 To 2) As Double
//     Dim pt35(0 To 2) As Double
//     Dim AreaName As String
//     Dim oprop As AcadDynamicBlockReferenceProperty
//     
//     'Insert the McClone area block
//     AreaName = CStr(ColumnCreator.sAreaBox.Value)
//     pt35(0) = pt_o(0) + 548.37926872: pt35(1) = pt_o(1) + 186: pt35(2) = 0
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt35, "VBA_OFFICE_LOCATION", (2 / 3), (2 / 3), (2 / 3), 0)
//     
//     Props = BlockRefObj.GetDynamicBlockProperties
//     For i = LBound(Props) To UBound(Props)
//         Set oprop = Props(i) 'Get object properties from border block
//         If oprop.PropertyName = "Visibility1" Then 'Finds visibility state property
//             oprop.Value = ColumnCreator.sAreaBox.Value
//         End If
//     Next i
//     
//     'Insert border
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_click, "VBA_24X36_LAYOUT", 1#, 1#, 1#, 0) 'Insert 3/4" = 1'-0" sheet boundry
//     BlockRefObj.ScaleEntity pt_click, 12 / 0.75
//     
//     'Sets Sheet issued for: "" text (defined as a block reference property)
//     Props = BlockRefObj.GetDynamicBlockProperties
//     For i = LBound(Props) To UBound(Props)
//         Set oprop = Props(i) 'Get object properties from border block
//         If oprop.PropertyName = "Visibility1" Then 'Finds visibility state property
//             oprop.Value = ColumnCreator.sSheetIssuedForBox.Value
//         End If
//     Next i
//     
//     'Place timestamp in lower left corner, outside of sheet boundry
//     pt34(0) = pt_o(0) - 1.65: pt34(1) = pt_o(1): pt34(2) = pt_o(2)
//     ThisDrawing.ActiveTextStyle = ThisDrawing.TextStyles("VBA_txt_narrow")
//     Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt34, 18, Now())
//     MTextObject2.Height = 1.25
//     MTextObject2.Rotate pt34, pi / 2
//     
//     'Places other text (defined as attributes)
//     AttList = BlockRefObj.GetAttributes
//     For i = LBound(AttList) To UBound(AttList)
//         ' Check for the correct attribute tag.
//         If AttList(i).TextString = "SHEET NAME" Then
//             AttList(i).TextString = "" & ColumnCreator.sProjectNameBox.Value
//         End If
//         If AttList(i).TextString = "PROJECT TITLE" Then
//             AttList(i).TextString = "" & ColumnCreator.sProjectTitleBox.Value
//         End If
//         If AttList(i).TextString = "PROJECT ADDRESS" Then
//             AttList(i).TextString = "" & ColumnCreator.sProjectAddressBox.Value
//         End If
//         If AttList(i).TextString = "DATE" Then
//             AttList(i).TextString = "" & ColumnCreator.sDateBox.Value
//         End If
//         If AttList(i).TextString = "SCALE" Then
//             AttList(i).TextString = "3/4'' = 1'-0''"
//         End If
//         If AttList(i).TextString = "DRAWN BY" Then
//             AttList(i).TextString = "" & ColumnCreator.sDrawnByBox.Value
//         End If
//         If AttList(i).TextString = "JOB NUMBER" Then
//             AttList(i).TextString = "" & ColumnCreator.sJobNoBox.Value
//         End If
//         If AttList(i).TextString = "SHEET_#" Then
//             AttList(i).TextString = "" & ColumnCreator.sSheetBox.Value
//         End If
//         If AttList(i).TextString = "SUFFIX" Then
//             AttList(i).TextString = "" & ColumnCreator.sSuffixBox.Value
//         End If
//     Next i
//
//
// last_line:
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
// ErrorHandler:
//     If Err.Number = -2147352567 Then 'This error occurs when the user presses the ESC key while inside the GetPoint utility
//         ColumnCreator.show
//         Exit Function
//     ElseIf Err.Number = -2145386422 Then 'This error occurs when the path to Egnyte isn't found
//         MsgBox ("Error: VBA block library not found at: " & FileToInsert & vbNewLine & vbNewLine & "Check Egnyte (Z:\) connection. Try closing and relaunching Egnyte Connect.")
//         Exit Function
//     ElseIf Err.Number = -2145386445 Then 'This error occurs when a block definition that doesn't exist is attempted to be inserted.
//         If BlockExists = True Then
//             MsgBox ("This drawing seems to be missing blocks for this VBA script. This is normal if you are using new button features on an old drawing. The block library will be re-downloaded and this script terminated." & vbNewLine & vbNewLine & "Please run the button again.")
//             'Insert the block library then delete it
//             FileToInsert = FileLocationPrefix & "VBA_Block_Library.dwg"
//             Set BlockDwg = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, FileToInsert, 1#, 1#, 1#, 0)
//             BlockDwg.Delete
//             Exit Function
//         Else
//             MsgBox ("Error: This drawing seems to be missing blocks for this VBA script despite the block library having been downloaded. The library may be missing the necessary blocks, or something mysterious has gone horribly wrong." & vbNewLine & vbNewLine & "Library location: " & FileToInsert)
//             Exit Function
//         End If
//     Else 'Generic error message
//         MsgBox ("An error occured." & vbNewLine & "error " & Err.Number & vbNewLine & Err.Description)
//         Exit Function
//     End If
//     
// SkipErrorHandler:
//
// End Function
//
//
//
// Public Function ChangeProp(BlockRefObj, PropName, PropValue)
//
//     Props = BlockRefObj.GetDynamicBlockProperties
//     Dim oprop As AcadDynamicBlockReferenceProperty
//     
//     For i = LBound(Props) To UBound(Props)
//         Set oprop = Props(i) 'Get object properties from block one at a time
//         If oprop.PropertyName = PropName Then 'compare to property name
//             oprop.Value = PropValue 'if name matches, assign new PropValue
//         End If
//     'Exit For
//     Next i
// End Function
// Public Function DrawDim(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double)
//
//     Dim dimObj As AcadDimAligned
//     Dim point1(0 To 2) As Double 'first point
//     Dim point2(0 To 2) As Double 'second point
//     Dim location(0 To 2) As Double 'text location
//     
//     ' Define the dimension
//     point1(0) = x1#:    point1(1) = y1#:    point1(2) = 0#
//     point2(0) = x2#:    point2(1) = y2#:    point2(2) = 0#
//     location(0) = x3#:  location(1) = y3#:  location(2) = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimAligned(point1, point2, location)
// End Function
// Public Function DrawDimSuffix(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, txt As String)
//
//     Dim dimObj As AcadDimAligned
//     Dim point1(0 To 2) As Double
//     Dim point2(0 To 2) As Double
//     Dim location(0 To 2) As Double
//     
//     ' Define the dimension
//     point1(0) = x1#:    point1(1) = y1#:    point1(2) = 0#
//     point2(0) = x2#:    point2(1) = y2#:    point2(2) = 0#
//     location(0) = x3#:  location(1) = y3#:  location(2) = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimAligned(point1, point2, location)
//     dimObj.TextSuffix = txt
//     
// End Function
// Public Function DrawDimLin(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, angle As Double)
//
//     Dim dimObj As AcadDimRotated
//     Dim point1(0 To 2) As Double 'first point
//     Dim point2(0 To 2) As Double 'second point
//     Dim location(0 To 2) As Double 'text location
//     
//     ' Define the dimension
//     point1(0) = x1#:    point1(1) = y1#:    point1(2) = 0#
//     point2(0) = x2#:    point2(1) = y2#:    point2(2) = 0#
//     location(0) = x3#:  location(1) = y3#:  location(2) = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimRotated(point1, point2, location, angle)
// End Function
// Public Function DrawDimLinSuffixLeader(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, txt As String, angle As Double)
//
//     Dim dimObj As AcadDimRotated
//     Dim point1(0 To 2) As Double
//     Dim point2(0 To 2) As Double
//     Dim location(0 To 2) As Double
//     Dim txtLocation(0 To 2) As Double
//     
//     ' Define the dimension
//     point1(0) = x1#: point1(1) = y1#: point1(2) = 0#
//     point2(0) = x2#: point2(1) = y2#: point2(2) = 0#
//     location(0) = x3#: location(1) = y3#: location(2) = 0#
//     txtLocation(0) = x4#: txtLocation(1) = y4#: txtLocation(2) = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimRotated(point1, point2, location, angle)
//     dimObj.TextMovement = acMoveTextAddLeader
//     dimObj.TextPosition = txtLocation
//     dimObj.TextSuffix = txt
// End Function
// Public Function DrawDimLinLeader(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, angle As Double)
//
//     Dim dimObj As AcadDimRotated
//     Dim point1(0 To 2) As Double
//     Dim point2(0 To 2) As Double
//     Dim location(0 To 2) As Double
//     Dim txtLocation(0 To 2) As Double
//     
//     ' Define the dimension
//     point1(0) = x1#: point1(1) = y1#: point1(2) = 0#
//     point2(0) = x2#: point2(1) = y2#: point2(2) = 0#
//     location(0) = x3#: location(1) = y3#: location(2) = 0#
//     txtLocation(0) = x4#: txtLocation(1) = y4#: txtLocation(2) = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimRotated(point1, point2, location, angle)
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
// Public Function DrawRectangle(xA As Double, yA As Double, xB As Double, yB As Double)
//     Dim plineObj As AcadLWPolyline
//     Dim pts(0 To 7) As Double
//     
//     pts(0) = xA#:   pts(1) = yA
//     pts(2) = xB#:   pts(3) = yA
//     pts(4) = xB#:   pts(5) = yB
//     pts(6) = xA#:   pts(7) = yB
//
//     Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pts)
//     plineObj.Closed = True
//
// End Function
// Public Function FindMax(Arr As Variant)
//     Dim RetMax As Double
//     RetMax = Arr(LBound(Arr)) 'initialize with first value
//     
//     For i = LBound(Arr) To UBound(Arr)
//     If Arr(i) > RetMax Then
//         RetMax = Arr(i)
//     End If
//     Next i
//     FindMax = RetMax
// End Function
//
// Public Function CreateScissorClamp()
//     'Use error handler
//     On Error GoTo ErrorHandler
//     
//     ColumnCreator.Hide 'Hide GUI form after clicking this command button
//     
//     'Referenced files are found within this folder, easier to change this partial path
//     Dim FileLocationPrefix As String
//     Dim filepath As String
//     
//     'FileLocationPrefix = "Z:\Shared\NCA\Users\Quincy\VBA\"
//     FileLocationPrefix = "D:\Rider projects\ColumnDesign\ColumnDesign\Source\"
//     filepath = FileLocationPrefix & "Columns\scissor_clamp_matrix.csv"
//     
//     'Define layers, assign colors, assign linetypes
//     Dim LayerObj As AcadLayer 'Creates layer object
//     Dim color As AcadAcCmColor
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
//     On Error GoTo ErrorHandler
//     
//     'Check if any of the used layers are locked, remember which are locked, then unlock them
//     Dim UsedLayerList(1 To 8) As String ' fill with used layers
//     Dim LockedLayerList(1 To 8) As Integer 'leave at default values of 0, change to 1 for locked layers
//     Dim FrozenLayerList(1 To 8) As Integer 'leave at default values of 0, change to 1 for frozen layers
//     Dim InvisLayerList(1 To 8) As Integer 'leave at default values of 0, change to 1 for invisible layers
//     UsedLayerList(1) = "DIM": UsedLayerList(2) = "ELEV": UsedLayerList(3) = "FMTEXTA": UsedLayerList(4) = "FORM": UsedLayerList(5) = "NOTES":
//     UsedLayerList(6) = "PLYA": UsedLayerList(7) = "SHADE": UsedLayerList(8) = "STEEL":
//
//     For i = 1 To UBound(UsedLayerList)
//         Set LayerObj = ThisDrawing.Layers(UsedLayerList(i))
//         If LayerObj.Lock = True Then
//             LayerObj.Lock = False
//             LockedLayerList(i) = 1
//         End If
//         If LayerObj.Freeze = True Then
//             LayerObj.Freeze = False
//             FrozenLayerList(i) = 1
//         End If
//         If LayerObj.LayerOn = False Then
//             InvisLayerList(i) = 1
//         End If
//     Next i
//     
//     'Make new dim style, text style, and layers
//     Dim OldDimStyle As AcadDimStyle
//     Dim OldTextStyle As AcadTextStyle
//     Dim OldLayer As AcadLayer
//     
//     Set OldDimStyle = ThisDrawing.ActiveDimStyle
//     Set OldTextStyle = ThisDrawing.ActiveTextStyle
//     Set OldLayer = ThisDrawing.ActiveLayer
//     
//     Call BuildColumnStyles
//
//     'Define various column parameters
//     Dim x As Double
//     Dim y As Double
//     Dim z As Double
//     Dim n_col As Integer
//     Dim n_studs_x As Integer 'Number of studs on the "x" side
//     Dim n_studs_y As Integer 'Number of studs on the "y" side
//     Dim n_clamps As Integer: n_clamps = 0 'Total number of clamps on column
//     Dim n_clamps_temp As Integer: n_clamps_temp = 0 'Number of clamps for one pour lift
//     Dim stud_type As Integer 'Stud types: 1 = 2x4,  2 = LVL
//     Dim clamp_size As Integer 'Clamp sizes: 4 = 36",  5 = 48",  6 = 60"
//     Dim col_wt As Integer
//     Dim ply_name As String
//     Dim stud_name As String
//     Dim stud_name_full As String 'A more elaborate name to include LVL dimensions
//     Dim stud_block As String
//     Dim stud_block_spax As String
//     Dim stud_block_bolt As String
//     Dim stud_face_block As String
//
//     'Other definitions
//     Dim ply_thk As Double: ply_thk = 0.75
//     Dim stud_base_gap As Double: stud_base_gap = 0.25 'Gap between bottom of stud and ground
//     Dim bot_clamp_gap As Double: bot_clamp_gap = 6 'space between the bottom clamp and the ground.
//     Dim stud_start_offset_x As Double: stud_start_offset_x = 0.125 'Distance from edge of ply to outermost stud on short panel
//     Dim stud_start_offset_y As Double: stud_start_offset_y = 2.375 'Distance from edge of ply to outermost stud on long panel
//     pi = 4 * Atn(1)
//     
//     'Assign input box values to variables
//     x = ConvertToNum(ColumnCreator.sWidthBox)
//     y = ConvertToNum(ColumnCreator.sLengthBox)
//     z = ConvertToNum(ColumnCreator.sHeightBox)
//     n_col = ConvertToNum(ColumnCreator.sQuantityBox)
//     ply_name = ColumnCreator.sPlyNameBox
//
//     'Assign stud type (always use 2x4s)
//     stud_type = 1
//     stud_name = "2X4"
//     stud_name_full = "2X4"
//     stud_block = "VBA_2X4": stud_block_spax = "VBA_2X4_SPAX": stud_block_bolt = "VBA_2X4_BOLT"
//     stud_face_block = "VBA_2X4_FACE"
//     
//     'Assign x and y values to array
//     Dim Arr(2) As Double: Arr(0) = x: Arr(1) = y
//     
//     'x: Add 0 inches to nominal width for ply sheet width on the short side
//     'y: Add 4.5 inches to nominal width for ply sheet width on the long side
//     Dim ply_width_x As Double
//     Dim ply_width_y As Double
//     ply_width_x = x
//     ply_width_y = y + 4.5
//         
//         
//     'If library block is missing: load VBA_Block_Library .dwg file as a block to load all needed blocks, then delete that file block
//     Dim BlockExists As Boolean
//     BlockExists = False
//
//     'Check if library was loaded:
//     For Each entry In ActiveDocument.Blocks
//         If entry.Name = "VBA_SHIBBOLETH" Then
//             BlockExists = True
//             Exit For
//         End If
//     Next
//     
//     Dim insertionPnt(0 To 2) As Double
//     insertionPnt(0) = 0#: insertionPnt(1) = 0#: insertionPnt(2) = 0#
//     FileToInsert = FileLocationPrefix & "VBA_Block_Library.dwg"
//     If BlockExists = False Then
//         Dim BlockDwg As AcadBlockReference
//         Set BlockDwg = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, FileToInsert, 1#, 1#, 1#, 0)
//         BlockDwg.Delete
//     End If
//
//
// '#####################################################################################
// '                                     S T U D S
// '#####################################################################################
//     'Compute number of studs required for each side
//     Dim row_num As Integer: row_num = 0 'height
//     Dim col_num_x As Integer: col_num_x = 0 'x
//     Dim col_num_y As Integer: col_num_y = 0 'y
//     Dim n_studs_total As Integer 'Total studs on all sides
//     
//     'Create stud_matrix, number of studs used for each column size (column height doesn't matter)
//     Dim stud_matrix(1 To 20, 1 To 2) As Integer
//     stud_matrix(1, 1) = 8:     stud_matrix(1, 2) = 2
//     stud_matrix(2, 1) = 10:     stud_matrix(2, 2) = 2
//     stud_matrix(3, 1) = 12:     stud_matrix(3, 2) = 2
//     stud_matrix(4, 1) = 14:     stud_matrix(4, 2) = 3
//     stud_matrix(5, 1) = 16:     stud_matrix(5, 2) = 3
//     stud_matrix(6, 1) = 18:     stud_matrix(6, 2) = 3
//     stud_matrix(7, 1) = 20:     stud_matrix(7, 2) = 3
//     stud_matrix(8, 1) = 22:     stud_matrix(8, 2) = 4
//     stud_matrix(9, 1) = 24:     stud_matrix(9, 2) = 4
//     stud_matrix(10, 1) = 26:     stud_matrix(10, 2) = 4
//     stud_matrix(11, 1) = 28:     stud_matrix(11, 2) = 4
//     stud_matrix(12, 1) = 30:     stud_matrix(12, 2) = 5
//     stud_matrix(13, 1) = 32:     stud_matrix(13, 2) = 5
//     stud_matrix(14, 1) = 34:    stud_matrix(14, 2) = 5
//     stud_matrix(15, 1) = 36:    stud_matrix(15, 2) = 5
//     stud_matrix(16, 1) = 38:    stud_matrix(16, 2) = 5
//     stud_matrix(17, 1) = 40:    stud_matrix(17, 2) = 6
//     stud_matrix(18, 1) = 42:    stud_matrix(18, 2) = 6
//     stud_matrix(19, 1) = 44:    stud_matrix(19, 2) = 6
//     stud_matrix(20, 1) = 46:    stud_matrix(20, 2) = 6
//
//     'Find number of studs for each side
//     'Number of x studs
//     For i = 1 To UBound(stud_matrix, 1)
//         If Round(stud_matrix(i, 1), 3) >= Round(x, 3) Then
//             n_studs_x = stud_matrix(i, 2)
//             Exit For
//         End If
//     Next i
//     
//     'Number of y studs
//     For i = 1 To UBound(stud_matrix, 1)
//         If Round(stud_matrix(i, 1), 3) >= Round(y, 3) Then
//             n_studs_y = stud_matrix(i, 2)
//             Exit For
//         End If
//     Next i
//     
//     'Calculate total number of studs
//     n_studs_total = n_studs_x * 2 + n_studs_y * 2
//     
//     'Compute stud spacing/locations assuming equal spacing
//     Dim stud_spacing_x() As Double
//     Dim stud_spacing_y() As Double
//     ReDim stud_spacing_x(1 To n_studs_x)
//     ReDim stud_spacing_y(1 To n_studs_y)
//
//     For i = 1 To n_studs_x
//         stud_spacing_x(i) = stud_start_offset_x + (i - 1) * (ply_width_x - 2 * stud_start_offset_x - 3.5) / (n_studs_x - 1)
//     Next i
//     For i = 1 To n_studs_y
//         stud_spacing_y(i) = stud_start_offset_y + (i - 1) * (ply_width_y - 2 * stud_start_offset_y - 3.5) / (n_studs_y - 1)
//     Next i
//     
//
//
// '#####################################################################################
// '                                   P L Y W O O D
// '#####################################################################################
//     Dim ply_top_ht_min As Double: ply_top_ht_min = 6 'Smallest allowable strip of plywood at top of column
//     Dim max_ply_ht As Double
//     Dim ply_seams() As Double '1 dimensional matrix that stores vertical length of each ply piece. First entry is bottom plywood.
//
//     'Read plywood seams from the userform
//     ply_seams = ReadSizes(ColumnCreator.sboxPlySeams.Value)
//
//     'Re-validate the ply seams if somehow they have an issue
//     If sUpdatePly.sValidatePlySeams(ply_seams, x, y, z) = 0 Then
//         MsgBox ("Error: Plywood layout invalid. You should never see this message...How did you do this?")
//         GoTo last_line
//     End If
//     
//     'Assign x or y information to the elevation view so it can easily be referenced
//     Dim n_studs_e As Integer
//     Dim ply_width_e As Double
//     Dim stud_spacing_e() As Double
//     
//     'Aassume elevation view will display side x
//     n_studs_e = n_studs_x
//     ply_width_e = ply_width_x
//     ReDim stud_spacing_e(1 To UBound(stud_spacing_x))
//     For i = 1 To UBound(stud_spacing_x)
//         stud_spacing_e(i) = stud_spacing_x(i)
//     Next i
//     
//     'If column is very wide, use rotated sheets and thus 48" max height
//     If Round(ply_width_x, 3) > 48 Or Round(ply_width_y, 3) > 48 Then
//         max_ply_ht = 48
//     Else
//         max_ply_ht = 96
//     End If
//
//     'Create 2D array for plywood cut counts. Columns: (width / height / quantity) 'ply_cuts(k, 1) = ply_width_x 'TRANSPOSE THIS STUPID MATRIX
//     'First find number of unique sheet sizes
//     Dim unique_plys As Integer
//     Dim ply_widths(1 To 2) As Double
//     Dim ply_cuts() As Double:   ReDim ply_cuts(1 To 3, 1 To 1) As Double    'Array for plywood sheet sizes for counts
//     
//     unique_plys = 0 '1 size of plywood is the bare minimum
//     
//     'Create array of plywood widths
//     ply_widths(1) = ply_width_x
//     ply_widths(2) = ply_width_y
//
//     'For each entry in ply_seams
//     For i = 1 To UBound(ply_seams)
//         'For each side of the panel (A and B)
//         For j = 1 To UBound(ply_widths)
//             'Check if this width and height combination has been entered before, then add it to the existing entry or create an new entry
//             For k = 1 To UBound(ply_cuts, 2)
//                 If Round(ply_cuts(1, k), 3) = Round(ply_widths(j), 3) And Round(ply_cuts(2, k), 3) = Round(ply_seams(i), 3) Then
//                     ply_cuts(3, k) = ply_cuts(3, k) + 2 'Add 2 to the existing count
//                     GoTo ExitPlywoodCheckLoop
//                 ElseIf k = UBound(ply_cuts, 2) Then
//                     unique_plys = unique_plys + 1
//                     ReDim Preserve ply_cuts(1 To 3, 1 To unique_plys)
//                     ply_cuts(1, unique_plys) = ply_widths(j)
//                     ply_cuts(2, unique_plys) = ply_seams(i)
//                     ply_cuts(3, unique_plys) = ply_cuts(3, unique_plys) + 2 'Start a new count at 2
//                     GoTo ExitPlywoodCheckLoop
//                 End If
//             Next k
// ExitPlywoodCheckLoop:
//         Next j
//     Next i
//
//     'Remove the last line if it's 0
//     If Round(ply_cuts(1, UBound(ply_cuts, 2)), 3) = 0 Then
//         ReDim Preserve ply_cuts(1 To 3, 1 To UBound(ply_cuts, 2) - 1)
//         unique_plys = unique_plys - 1
//     End If
//
//     Dim msg As String
//     For i = 1 To UBound(ply_cuts, 2)
//         For ii = 1 To UBound(ply_cuts, 1)
//             msg = msg & ply_cuts(ii, i) & vbTab 'replace with space(1) if array too big
//         Next ii
//         msg = msg & vbCrLf
//     Next i
//     'MsgBox msg
//
//     
// '#####################################################################################
// '                                   C L A M P S
// '#####################################################################################
//     'Compute clamp size. Use maximum dimension to pick clamp.
//     '1 = 36",  2 = 48",  3 = 60"
//     Dim clamp_spacing() As Double
//     Dim n_pours As Integer 'Minimum number of pours needed to fill scissor form given full liquid head limit from clamp spacing matrix.
//     
//     If ColumnCreator.sClampSizeButton1.Value = True Then
//         GoTo Clampsize1
//     ElseIf ColumnCreator.sClampSizeButton2.Value = True Then
//         GoTo Clampsize2
//     ElseIf ColumnCreator.sClampSizeButton3.Value = True Then
//         GoTo Clampsize3
//     ElseIf ColumnCreator.sClampSizeButton4.Value = True Then
//         GoTo ClampSizeChoose
//     End If
//     
// ClampSizeChoose:
//     If Round(FindMax(Arr), 3) < 23.5 Then
// Clampsize1:
//         clamp_size = 4
//         clamp_L = 36
//         clamp_width = 2.5
//         clamp_name = "36"" SCISSOR CLAMP"
//         clamp_block = "VBA_36_SCISSOR_CLAMP_PLAN"
//         clamp_block_pr = "VBA_36_SCISSOR_CLAMP_PROFILE"
//         clamp_block_pr_flip = "VBA_36_SCISSOR_CLAMP_PROFILE_FLIP"
//     ElseIf Round(FindMax(Arr), 3) <= 35.5 Then
// Clampsize2:
//         clamp_size = 5
//         clamp_L = 48
//         clamp_width = 2.5
//         clamp_name = "48"" SCISSOR CLAMP"
//         clamp_block = "VBA_48_SCISSOR_CLAMP_PLAN"
//         clamp_block_pr = "VBA_48_SCISSOR_CLAMP_PROFILE"
//         clamp_block_pr_flip = "VBA_48_SCISSOR_CLAMP_PROFILE_FLIP"
//     ElseIf Round(FindMax(Arr), 3) <= 46 Then
// Clampsize3:
//         clamp_size = 6
//         clamp_L = 60
//         clamp_width = 3
//         clamp_name = "60"" SCISSOR CLAMP"
//         clamp_block = "VBA_60_SCISSOR_CLAMP_PLAN"
//         clamp_block_pr = "VBA_60_SCISSOR_CLAMP_PROFILE"
//         clamp_block_pr_flip = "VBA_60_SCISSOR_CLAMP_PROFILE_FLIP"
//     Else
//         MsgBox ("Error: Column is too wide (>46"") in one dimension")
//         ColumnCreator.show
//         Exit Function
//     End If
//     
//     'Determine clamp spacing
//     clamp_matrix = ImportMatrix(filepath)
//     
//     'Find row number corresponding to longer side of column
//     If x <= y Then
//         long_side = y
//     Else
//         long_side = x
//     End If
//     
//     For i = 0 To UBound(clamp_matrix, 1) 'Cycle through each row
//         If clamp_matrix(i, 1) >= Round(long_side, 3) Then 'Look at second entry of each row and compare to the width of the long side.
//             row_num = i 'If value equals or exceds that width then that's the row we want.
//             Exit For
//         End If
//     Next i
//
//     'Calculate the number of pours needed, given available full liquid head
//     n_pours = Int(z / clamp_matrix(row_num, 0)) 'Rounds down to the nearest integer
//     If Round(n_pours * clamp_matrix(row_num, 0), 3) < Round(z, 3) Then n_pours = n_pours + 1 'Adds 1 if there's any remainder
//     
//     'Save the original column height
//     Dim z_original As Double
//     z_original = z
//
//     'Create a variable for the overlap that occurs when the remainder after a pour is < 12 inches. The next pour will be set to 12 inches and the overlap of that pour with the previous pour is pour_overlap
//     Dim pour_overlap As Double: pour_overlap = 0
//     
//     'Define a matrix to contain all the clamp spacings for the "sub columns" if multiple pours are required. Use an arbitrary large value for the number of clamps, which is at this point unknown.
//     Dim multi_clamp_spacing() As Variant
//     ReDim multi_clamp_spacing(1 To 100, 1 To n_pours) '1st dimension: clamp spacing, 2nd dimension: the pour number
//     Dim total_clamp_spacings() As Integer 'Tracks the size of clamp_spacing for each lift
//     ReDim total_clamp_spacings(1 To n_pours)
//     
//     'Run through the clamp spacing procedure for each pour designing a "virtual column" for each
//     For p = 1 To n_pours
//         'Redefine z for the "virtual column", it must be at least 12" tall
//         If n_pours >= 2 Then
//             If Round(z_original - p * clamp_matrix(row_num, 0), 3) >= 0 Then
//                 z = clamp_matrix(row_num, 0) 'Pour goes to max allowable pour height
//             Else 'Pour "tops off" column
//                 z = z_original - (p - 1) * clamp_matrix(row_num, 0)
//                 If z < 12 Then 'Ensure z is at least 12 inches, if less than 12 make it 12 and record the overlap with the previous pour
//                     pour_overlap = 12 - z
//                     z = 12
//                 End If
//             End If
//         End If
//         
//         'Build clamp spacing matrix
//         ReDim clamp_spacing(1 To UBound(clamp_matrix, 2))
//         k = 0
//
//         For i = 2 To UBound(clamp_matrix, 2) 'First column (index 0) is column heights, second column (index 1) is column side dim, clamp spacings start on 3rd column (index 2)
//             If clamp_matrix(row_num, i) = 0 Then 'Store all non-zero values. When a 0 is encountered, exit the loop and redim to truncate off the trailing 0s
//                 ReDim Preserve clamp_spacing(1 To k)
//                 Exit For
//             Else
//                 clamp_spacing(i - 1) = clamp_matrix(row_num, i)
//                 k = k + 1
//             End If
//         Next i
//         
//         'Reduce clamp_spacing to fit actual column size. Use number of clamps for the larger size (a 9'-8" column uses clamps designed for a 10'-0" column)
//         'First cut off all clamps below actual height
//         Dim ht_rem As Double
//         ht_rem = z - bot_clamp_gap
//         n_clamps_temp = 0 'Reset number of clamps
//         
//         For i = 1 To UBound(clamp_spacing) - 1 'Look at the length of column above the bottom clamp
//             If Round(clamp_spacing(i), 3) <= Round(ht_rem, 3) Then 'Starting at the top, subtract clamp spacings until the clamp spacing is larger than whatever distance remains
//                 n_clamps_temp = n_clamps_temp + 1
//                 ht_rem = ht_rem - clamp_spacing(i)
//             Else
//                 n_clamps_temp = n_clamps_temp + 1 'When the clamp no longer fits add back the bottom clamp to the count and exit the loop
//                 Exit For
//             End If
//         Next i
//
//         'Cut off all clamps below what fits onto this column and add back the bottom spacing
//         ReDim Preserve clamp_spacing(1 To n_clamps_temp + 1)
//         clamp_spacing(UBound(clamp_spacing)) = bot_clamp_gap
//
//         'Sum all clamp spacings from truncated matrix
//         tot_temp = 0
//         For i = 1 To UBound(clamp_spacing)
//             tot_temp = tot_temp + clamp_spacing(i)
//         Next i
//         
//         'Reduce spacing of 2nd from bottom clamp so sum of clamp spacings = z
//         clamp_spacing(UBound(clamp_spacing) - 1) = clamp_spacing(UBound(clamp_spacing) - 1) - (tot_temp - z)
//
//         'Cycle through all clamp spacings and check that they're >= 5". If less, move lowest non-compliant clamp to 5" and adjust the clamp above it accordingly, repeat until all clamps pass
//         Dim infExit As Integer: infExit = 0
// TestClampSpacing:
//         For i = 2 To UBound(clamp_spacing)
//             If Round(clamp_spacing(i), 3) < 5 And i >= 2 Then
//                 clamp_spacing(i - 1) = clamp_spacing(i - 1) - (5 - clamp_spacing(i))
//                 clamp_spacing(i) = 5
//                 infExit = infExit + 1
//                 If infExit > 100 Then
//                     MsgBox ("Error: Infinite loop encountered while computing clamp spacings.")
//                     Exit For
//                 Else
//                     GoTo TestClampSpacing
//                 End If
//             End If
//         Next i
//
//         'Store the clamp spacing for this pour in multi_clamp_spacing()
//         For i = 1 To UBound(clamp_spacing)
//             multi_clamp_spacing(i, p) = clamp_spacing(i)
//         Next i
//         
//         'Record total clamp spacing
//         total_clamp_spacings(p) = UBound(clamp_spacing)
//         
//         'Record total number of clamps
//         n_clamps = n_clamps + n_clamps_temp
//         If p >= 2 Then n_clamps = n_clamps + 1 'On pours after the first add a clamp "between" clamp spacings
//     Next p
//     
//     'Assemble a complete clamp spacing list from the pour(s) in multi_clamp_spacing()
//     Dim clamp_spacing_r() 'Reverse order of clamp_spacings will be defined first, then flipped
//     ReDim clamp_spacing_r(1 To 100)
//     ReDim clamp_spacing(1 To 100)
//     j = 1
//     For p = n_pours To 1 Step -1
//         For i = 1 To total_clamp_spacings(p)
//             If multi_clamp_spacing(i, p) <> 0 Then 'Only record non-zero values, most of this array will probably be empty
//                 If p <> n_pours And i = 1 Then 'Combine clamp spacing between pours
//                     clamp_spacing(j - 1) = multi_clamp_spacing(i, p) + multi_clamp_spacing(total_clamp_spacings(p + 1), p + 1)
//                     n_clamps = n_clamps - 1
//                     'MsgBox clamp_spacing(j - 1) & vbNewLine & vbNewLine & multi_clamp_spacing(i, p) & vbNewLine & multi_clamp_spacing(total_clamp_spacings(p + 1), p + 1)
//                 Else
//                     clamp_spacing(j) = multi_clamp_spacing(i, p)
//                     j = j + 1
//                 End If
//             End If
//             
//         Next i
//     Next p
//     
//     'Iif a "pour overlap" occurs, move the top-most clamp of the previus pour downward
//     If pour_overlap <> 0 Then
//          clamp_spacing(total_clamp_spacings(n_pours)) = clamp_spacing(total_clamp_spacings(n_pours)) + pour_overlap
//          clamp_spacing(total_clamp_spacings(n_pours) + 1) = clamp_spacing(total_clamp_spacings(n_pours) + 1) - pour_overlap
//     End If
//
//     'Clean up clamp_spacing by deleting 0 values
//     'Only truncate if the next 2 values are 0
//     For i = 1 To UBound(clamp_spacing)
//         If clamp_spacing(i) = 0 And clamp_spacing(i + 1) = 0 Then
//             ReDim Preserve clamp_spacing(1 To i - 1)
//             Exit For
//         End If
//     Next i
//
//     msg = ""
//     For i = LBound(clamp_spacing) To UBound(clamp_spacing)
//         msg = msg & clamp_spacing(i) & vbNewLine
//     Next i
//     msg = "clamp spacing: " & vbNewLine & msg
//     'MsgBox msg
//     
//     'Check if the top clamp is <2" from the top. If it's easy to move it to 2" by adjusting the top 3 clamp spacings, do so
//     If clamp_spacing(1) < 2 And clamp_spacing(1) >= 0 Then
//         If clamp_spacing(2) >= (2 - clamp_spacing(1)) + 5 Then 'Reduce 2nd-from-top clamp spacing to make room on top
//             clamp_spacing(2) = clamp_spacing(2) - (2 - clamp_spacing(1))
//             clamp_spacing(1) = 2
//         ElseIf clamp_spacing(3) >= (2 - clamp_spacing(1)) + 5 Then 'Reduce 3rd-from-top clamp spacing to make room on top
//             clamp_spacing(3) = clamp_spacing(3) - (2 - clamp_spacing(1))
//             clamp_spacing(1) = 2
//         End If
//     End If
//     
//     'Restore original column height
//     z = z_original
//     
//     'Make array of clamp locations by continuous dimensions
//     Dim clamp_spacing_con()
//     ReDim clamp_spacing_con(1 To UBound(clamp_spacing))
//     clamp_spacing_con(1) = z - clamp_spacing(1)
//     For i = 2 To UBound(clamp_spacing)
//         clamp_spacing_con(i) = clamp_spacing_con(i - 1) - clamp_spacing(i)
//     Next i
//
//
//
//
//
// '#####################################################################################
// '                                     T E X T
// '#####################################################################################
//     
//     'Insert material count text
//     Dim MTextObject1 As AcadMText
//     Dim MTextObject2 As AcadMText
//     Dim objTextStyle As AcadTextStyle
//     Dim pt_click() As Double
//     Dim qty_text As String
//     Dim pt_o(0 To 2) As Double 'lower left corner of sheet border
//     Dim pt1(0 To 2) As Double  'Fabrication notes
//     Dim pt2(0 To 2) As Double  'Title text for fabrication notes and quantities
//     Dim pt3(0 To 2) As Double  'Quantity/components text
//     Dim panel_wt As Double  'Weight of 1 panel
//     Dim BlockRefObj As AcadBlockReference 'Create block reference object used for inserting all blocks
//     
//     'Ask user to select point
//     ThisDrawing.ActiveSpace = acModelSpace
//     pt_click = ThisDrawing.Utility.GetPoint(, "Click to place")
//     
//     'Set origin point (lower left corner of the sheet)
//     pt_o(0) = pt_click(0) + 20.23568507: pt_o(1) = pt_click(1) + 5.17943146: pt_o(2) = 0
//     
//     'Add border and header for text
//     pt1(0) = pt_o(0) + 548.37926872: pt1(1) = pt_o(1) + 369.60966303
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt1, "VBA_SCISSOR_CLAMP_NOTES_FRAME", 1#, 1#, 1#, 0) 'Insert plywood sheets
//     
//     'Call module to calculate total weight
//     col_wt = wt_total(x, y, z, n_studs_x, n_studs_y, stud_type, n_clamps, clamp_size, 0)
//     panel_wt = wt_panel(x, y, z, n_studs_x, n_studs_y, stud_type)
//     
//     'Define some other points for text
//     pt1(0) = pt_o(0) + 448.4: pt1(1) = pt_o(1) + 360
//     pt2(0) = pt1(0) + 0: pt2(1) = pt1(1) + 7
//     pt3(0) = pt1(0) + 0: pt3(1) = pt1(1) - 68
//     
//
//     'FABRICATION NOTES AND QUANTITIES
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
//     Dim strFabNotes As String
//     
//     strFabNotes = _
//     "• COLUMN SIZE = " & x & "'' X " & y & "''" & vbCrLf & vbCrLf & _
//     "• NUMBER OF COLUMN FORMS = " & n_col & "-EA" & vbCrLf & vbCrLf & _
//     "• COLUMN FORM WEIGHT (APPROXIMATE) = " & col_wt & "-LBS" & vbCrLf & _
//     "• COLUMN PANEL WEIGHT (SINGLE PANEL) = " & panel_wt & "-LBS" & vbCrLf & vbCrLf & _
//     "• PLYWOOD = 3/4'' PLYFORM (" & ply_name & "), CLASS-1 (MIN)" & vbCrLf & vbCrLf & _
//     "• COLUMN FORMS AND CLAMP SPACING LAYOUTS FOR SCISSOR CLAMPS ARE DESIGNED FOR A POUR RATE = "
//     If Round(z, 3) <= Round(clamp_matrix(row_num, 0), 3) Then
//         strFabNotes = strFabNotes & "FULL LIQUID HEAD U.N.O." & vbCrLf & vbCrLf
//     Else
//         strFabNotes = strFabNotes & ConvertFtIn(ConvertToNum(clamp_matrix(row_num, 0))) & vbCrLf & vbCrLf
//     End If
//     strFabNotes = strFabNotes & _
//     "• CONTACT THE MCC ENGINEER PRIOR TO ANY CHANGES OR MODIFICATIONS TO THE DETAILS ON THIS SHEET."
//     
//     Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt1, 100, strFabNotes)
//     MTextObject1.StyleName = "Arial Narrow"
//     MTextObject1.Height = 2
//     
//     'Plywood quantities
//     For i = 1 To UBound(ply_cuts, 2)
//         qty_text = qty_text & "• (" & ply_cuts(3, i) * n_col & "-EA) = (" & n_col & "-COL) X (" & ply_cuts(3, i) & "-EA/COL) @ " & ConvertFtIn(ply_cuts(1, i)) & " WIDE X " & ConvertFtIn(ply_cuts(2, i)) & " LONG 3/4'' PLYWOOD" & vbCrLf
//     Next i
//     qty_text = "PLYWOOD" & vbCrLf & qty_text & vbCrLf
//     
//     'Stud quantities
//     qty_text = qty_text & "STUDS" & vbCrLf & "• (" & n_studs_total * n_col & "-EA) = (" & n_col & "-COL) X (" & n_studs_total & "-EA/COL) @ " & ConvertFtIn(z - 0.25) & " " & stud_name_full
//     
//     'Clamp quantities
//     qty_text = qty_text & vbCrLf & vbCrLf & "SCISSOR CLAMP SETS (2 CLAMPS PER SET)" & vbCrLf
//     qty_text = qty_text & "• (" & n_clamps * n_col & "-EA) = (" & n_col & "-COL) X (" & n_clamps & "-EA/COL) @ " & clamp_name & " SETS"
//     
//     'Chamfer quantity
//     'qty_text = qty_text & "3/4'' GATES PLASTIC CHAMFER (BASED ON 12' CHAMFER LENGTHS)" & vbCrLf & _
//     '"•   (" & n_col * n_chamf & "-EA) = (" & n_col & "-COL) X (" & n_chamf & "-EA/COL) @ " & chamf_length & "'-0'' LONG PIECES"
//     
//     'Set all count and fab note texts
//     Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt3, 100, qty_text)
//     MTextObject1.Height = 2
//     MTextObject1.StyleName = "Arial Narrow" 'Forces use of Arial Narrow text style
//         
//         
//         
//         
//
//     '#####################################################################################
//     '                                 D R A W I N G
//     '#####################################################################################
//
//     Dim ObjDynBlockProp As AcadDynamicBlockReferenceProperty
//     Dim LineObj As AcadLine
//     Dim PlyTemp As Double
//     Dim ptA(0 To 2) As Double 'lower left corner of side A elevation view plywood
//     Dim ptB(0 To 2) As Double 'lower left corner of side B elevation view plywood
//     Dim ptE(0 To 2) As Double 'Lower left corner of clamp elevation view (right side of ply edge)
//     
//     Dim pt4(0 To 2) As Double 'Plywood seam locations for sides A and B
//     Dim pt5(0 To 2) As Double 'Stud locations on turned out clamp in plan view
//     Dim pt6(0 To 2) As Double 'stud locations for sides A, B, and W
//     Dim pt7(0 To 2) As Double 'UNUSED
//     Dim pt8(0 To 2) As Double 'Elevation section ply base for side A
//     Dim pt9(0 To 2) As Double 'Elevation section ply base for side B
//     Dim pt10(0 To 2) As Double 'Elevation section ply base for side W
//     Dim pt11(0 To 2) As Double 'stud locations at elevation section for sides A, B, and W
//     Dim pt12(0 To 2) As Double 'ply+chamfer for plan view
//     Dim pt13(0 To 2) As Double 'origin point on plan view
//     Dim pt14(0 To 2) As Double 'stud locs for plan view
//     Dim pt15(0 To 2) As Double 'clamp locs for plan view
//     Dim pt16(0 To 2) As Double 'brace locs for plan view
//     Dim pt17(0 To 2) As Double 'Text locs for plan view
//     Dim pt18(0 To 2) As Double 'Chamfer locations for elevation views
//     Dim pt19(0 To 2) As Double
//     Dim pt20(0 To 2) As Double 'stud locations at base of clamp elevation and framing lumber at base
//     Dim pt21(0 To 2) As Double 'Clamp locations in elevation
//     Dim pt22(0 To 2) As Double 'Clamp dimensions in inches from bottom
//     Dim pt23(0 To 2) As Double 'Start point for top clamp dim line
//     Dim pt24(0 To 2) As Double 'End point for top clamp dim line
//     Dim pt25(0 To 2) As Double '"CLAMP" and "TOP CLAMP" labels
//     Dim pt26(0 To 2) As Double 'Lifting sling points
//     Dim pt27(0 To 2) As Double 'Coil rod and nuts for top clamps
//     Dim pt28(0 To 2) As Double 'Brace related locations
//     Dim pt29(0 To 2) As Double 'Screw heads for ply elevations
//     Dim pt30(0 To 2) As Double 'Misc. notes and details
//     Dim pt31(0 To 2) As Double 'Reinforcing angles for wide columns
//     Dim pt32(0 To 2) As Double 'Lower left text for column count
//     Dim pt33(0 To 2) As Double 'A & B elevation sights for plan views
//     Dim pt34(0 To 2) As Double 'Insertion point for timestamp (lower left corner, outside of sheet)
//     Dim pt_blk(0 To 2) As Double
//     
//     'Control which elevation views are drawn and space them accordingly
//     Dim DrawB As Boolean: DrawB = True    'Side B (if different)
//
//     'Define reference points for each view
//     ptA(0) = pt_o(0) + 300: ptA(1) = pt_o(1) + 25
//     ptB(0) = pt_o(0) + 185: ptB(1) = pt_o(1) + 25
//     ptE(0) = pt_o(0) + 54.5: ptE(1) = pt_o(1) + 25
//     DrawB = True
//
//     'ELEVATION VIEWS
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
//     
//     'Profile of other side's studs (left)
//     Call DrawRectangle(ptE(0) - ply_thk, ptE(1) + stud_base_gap, ptE(0) - ply_thk - 1.5, ptE(1) + z)
//     
//     'Profile of other side's studs (right)
//     Call DrawRectangle(ptE(0) + ply_width_e + ply_thk, ptE(1) + stud_base_gap, ptE(0) + ply_width_e + ply_thk + 1.5, ptE(1) + z)
//     
//     'Draw studs
//         'Side A
//         For i = 1 To n_studs_x
//             pt6(0) = ptA(0) + ply_width_x - 3.5 - stud_spacing_x(i)
//             pt6(1) = ptA(1) + stud_base_gap
//             Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt6, stud_face_block, 1#, 1#, 1#, pi / 2) 'Insert studs
//             Call ChangeProp(BlockRefObj, "Distance1", z - stud_base_gap)
//         Next i
//         
//         'Side B
//         For i = 1 To n_studs_y
//             pt6(0) = ptB(0) + ply_width_y - 3.5 - stud_spacing_y(i)
//             pt6(1) = ptB(1) + stud_base_gap
//             Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt6, stud_face_block, 1#, 1#, 1#, pi / 2) 'Insert studs
//             Call ChangeProp(BlockRefObj, "Distance1", z - stud_base_gap)
//         Next i
//         
//     'Draw plywood sheets
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("PLYA")
//     pt4(1) = ptA(1)
//     For i = 1 To UBound(ply_seams)
//         pt4(0) = ptA(0)
//         pt4(1) = pt4(1) + ply_seams(i)
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt4, "VBA_PLY_SHEET", 1#, 1#, 1#, 0) 'Insert plywood sheets
//         Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
//         Call ChangeProp(BlockRefObj, "Distance2", ply_seams(i))
//     
//         pt4(0) = ptB(0)
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt4, "VBA_PLY_SHEET", 1#, 1#, 1#, 0) 'Insert plywood sheets
//         Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
//         Call ChangeProp(BlockRefObj, "Distance2", ply_seams(i))
//     Next i
//
//
//     'Draw elevation view (clamps view) specifc items
//     'Studs
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
//     For i = 1 To n_studs_e
//         pt20(0) = ptE(0) + stud_spacing_e(i): pt20(1) = ptE(1) + stud_base_gap
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt20, stud_face_block, 1#, 1#, 1#, pi / 2) 'Insert studs
//         Call ChangeProp(BlockRefObj, "Distance1", z - stud_base_gap): Call ChangeProp(BlockRefObj, "Visibility1", "Solid")
//     Next i
//     
//     'Plywood
//     'Outer ply boundry of elevation view
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("PLYA")
//     Call DrawRectangle(ptE(0), ptE(1), ptE(0) + ply_width_e, ptE(1) + z)
//     
//     'Plywood edge view
//     pt21(0) = ptE(0): pt21(1) = ptE(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt21, "VBA_PLY", 1#, 1#, 1#, pi / 2) 'left ply
//     Call ChangeProp(BlockRefObj, "Distance1", z)
//     pt21(0) = ptE(0) + ply_width_e + ply_thk: pt21(1) = ptE(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt21, "VBA_PLY", 1#, 1#, 1#, pi / 2) 'right ply
//     Call ChangeProp(BlockRefObj, "Distance1", z)
//     
//     'Clamps and dims
//     'Add top clamp dim
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("FMTEXTA")
//     pt21(0) = ptE(0) - 2.25: pt21(1) = ptE(1) + z
//     Call DrawDimLin(pt21(0) + x + 4.5 + clamp_width, pt21(1) - clamp_spacing(1), pt21(0) + ply_width_e + 3.75 + ply_thk, pt21(1), pt21(0) + clamp_L + 5, (pt21(1) * 2 - clamp_spacing(1)) / 2, pi / 2)
//     
//     For i = 1 To n_clamps
//         'Draw clamps and dims
//         pt21(1) = pt21(1) - clamp_spacing(i)
//         If i Mod 2 = 0 Then
//             pt19(0) = pt21(0): pt19(1) = pt21(1)
//             Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt19, clamp_block_pr, 1#, 1#, 1#, 0) 'Draw clamp profile view
//         Else
//             pt19(0) = pt21(0) + 4.5 + x: pt19(1) = pt21(1)
//             Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt19, clamp_block_pr_flip, 1#, 1#, 1#, 0) 'Draw clamp profile view
//         End If
//         Call ChangeProp(BlockRefObj, "Distance1", 4.5 + ply_width_e)
//         
//         'Draw dimensions between clamps
//         If i <> 1 Then
//             If i Mod 2 = 0 Then
//                 Call DrawDimLin(pt21(0) + clamp_L, pt21(1), pt21(0) + x + 4.5 + clamp_width, pt21(1) + clamp_spacing(i), pt21(0) + clamp_L + 5, (pt21(1) * 2 + clamp_spacing(i)) / 2, -pi / 2)
//             Else
//                 Call DrawDimLin(pt21(0) + x + 4.5 + clamp_width, pt21(1), pt21(0) + clamp_L, pt21(1) + clamp_spacing(i), pt21(0) + clamp_L + 5, (pt21(1) * 2 + clamp_spacing(i)) / 2, -pi / 2)
//             End If
//         End If
//
//         'Draw inches from bottom
//         pt22(0) = pt21(0) + clamp_L + 7: pt22(1) = ptE(1) + clamp_spacing_con(i) + 1.25   'pt21(0) - clamp_width - 8
//         clamp_str = CStr(clamp_spacing_con(i)) & """" 'ConvertFtIn(clamp_spacing_con(i))
//         Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt22, 10, clamp_str)
//         MTextObject2.Height = 2.5
//     Next i
//     
//     'Add bottom clamp dim
//     If i Mod 2 = 0 Then
//         Call DrawDimLin(pt21(0) + x + 4.5 + clamp_width, ptE(1) + bot_clamp_gap, pt21(0) + ply_width_e + 2.25 + ply_thk + 2.75, ptE(1), pt21(0) + clamp_L + 5, (2 * ptE(1) + bot_clamp_gap) / 2, pi / 2)
//     Else
//         Call DrawDimLin(pt21(0) + clamp_L, ptE(1) + bot_clamp_gap, pt21(0) + ply_width_e + 2.25 + ply_thk + 2.75, ptE(1), pt21(0) + clamp_L + 5, (2 * ptE(1) + bot_clamp_gap) / 2, pi / 2)
//     End If
//     
//     'Mark allowable pour heights
//     Dim strOrdinal As String
//     If n_pours >= 2 Then
//         ThisDrawing.ActiveLayer = ThisDrawing.Layers("NOTES")
//         For i = 1 To n_pours - 1
//             'Draw a dotted line at that height
//             pt20(0) = ptE(0) - 24: pt20(1) = ptE(1) + clamp_matrix(row_num, 0) * i
//             pt21(0) = ptE(0) + x + 58: pt21(1) = ptE(1) + clamp_matrix(row_num, 0) * i
//             Set LineObj = ThisDrawing.ModelSpace.AddLine(pt20, pt21)
//             LineObj.Linetype = "HIDDEN"
//             'color.ColorIndex = acYellow
//             'LineObj.TrueColor = color
//             
//             'Add text
//             'Determine which to use for the ordinal count: st, nd, rd, or th
//             ThisDrawing.ActiveLayer = ThisDrawing.Layers("FMTEXTA")
//             If i >= 11 And i <= 19 Then
//                 strOrdinal = "TH"
//             ElseIf Right(i, 1) = 1 Then
//                 strOrdinal = "ST"
//             ElseIf Right(i, 1) = 2 Then
//                 strOrdinal = "ND"
//             ElseIf Right(i, 1) = 3 Then
//                 strOrdinal = "RD"
//             Else
//                 strOrdinal = "TH"
//             End If
//             pt20(0) = ptE(0) + x + 40: pt20(1) = ptE(1) + clamp_matrix(row_num, 0) * i + 3.25
//             Set mtextobj = ThisDrawing.ModelSpace.AddMText(pt20, 24, "MAX HT OF " & vbNewLine & i & strOrdinal & " POUR")
//             mtextobj.Height = 2.5
//             
//             'Add height
//             pt20(0) = ptE(0) - 32: pt20(1) = ptE(1) + clamp_matrix(row_num, 0) * i + 1.25
//             Set mtextobj = ThisDrawing.ModelSpace.AddMText(pt20, 8, i * clamp_matrix(row_num, 0) & """")
//             mtextobj.Height = 2.5
//         Next i
//     End If
//     
//     'Base framing lumber
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
//     pt20(0) = ptE(0) - ply_thk - 1.5 - 3.5: pt20(1) = ptE(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt20, "VBA_2X4", 1#, 1#, 1#, 0) 'Insert studs
//     pt20(0) = ptE(0) + ply_width_e: pt20(1) = ptE(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt20, "VBA_2X4", 1#, 1#, 1#, 0) 'Insert studs
//     pt20(0) = ptE(0) - ply_thk - 1.5 - 6: pt20(1) = ptE(1) + 3
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt20, "VBA_2X4_SIDE", 1#, 1#, 1#, 0) 'Insert studs
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_e + ply_thk + 1.5 + 12)
//     pt30(0) = ptE(0) + ply_width_e + 6: pt30(1) = ptE(1) + 2.25
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt30, "VBA_DOWN_PLATE_NOTES_SCISSOR", 1#, 1#, 1#, 0)
//
//
//     'Draw panel sections
//     'Side A
//     pt8(0) = ptA(0) + ply_width_x: pt8(1) = ptA(1) + z + 18
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt8, "VBA_PLY", 1#, 1#, 1#, pi) 'Insert ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
//     For i = 1 To n_studs_x
//         pt11(0) = ptA(0) + ply_width_x - 3.5 - stud_spacing_x(i)
//         pt11(1) = ptA(1) + z + 18
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt11, stud_block, 1#, 1#, 1#, 0) 'Insert studs
//     Next i
//     
//     'Side B
//     pt9(0) = ptB(0) + ply_width_y: pt9(1) = ptB(1) + z + 18
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt9, "VBA_PLY", 1#, 1#, 1#, pi) 'Insert ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
//     For i = 1 To n_studs_y
//         pt11(0) = ptB(0) + ply_width_y - 3.5 - stud_spacing_y(i)
//         pt11(1) = ptB(1) + z + 18
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt11, stud_block, 1#, 1#, 1#, 0) 'Insert studs
//     Next i
//
//     'Dim sections
//         'Side A
//         Call DrawDim(pt8(0), pt8(1) - ply_thk, pt8(0) - ply_width_x, pt8(1) - ply_thk, (pt8(0) - ply_width_x - pt8(0)) / 2, pt8(1) - ply_thk - 5) 'Dimension plywood
//         For i = 1 To n_studs_x
//             Call DrawDimLin(pt8(0) - stud_spacing_x(UBound(stud_spacing_x) - i + 1), pt8(1) + 1.5, pt8(0) - ply_width_x, pt8(1), ((pt8(0) - ply_width_x) - (pt8(0) - stud_spacing_x(UBound(stud_spacing_x) - i + 1))) / 2, pt8(1) + 4 + i * 4, 0)   'Dimension studs from rightmost (in plan) stud
//         Next i
//         Call DrawDimLin(pt8(0) - stud_spacing_x(UBound(stud_spacing_x)) - 3.5, pt8(1) + 1.5, pt8(0) - ply_width_x, pt8(1), ((pt8(0) - ply_width_x) - (pt8(0) - stud_spacing_x(UBound(stud_spacing_x)) - 3.5)) / 2, pt8(1) + 4, 0)   'Dimension leftmost (in plan) stud to face of ply
//         Call DrawDimLin(pt8(0) - stud_start_offset_x, pt8(1) + 1.5, pt8(0), pt8(1), ((pt8(0)) - (pt8(0) - stud_start_offset_x)) / 2, pt8(1) + 4 * (n_studs_x + 1), 0)    'dimension stud_start_gap
//         
//         'Side B
//         Call DrawDim(pt9(0), pt9(1) - ply_thk, pt9(0) - ply_width_y, pt9(1) - ply_thk, (pt9(0) - ply_width_y - pt9(0)) / 2, pt9(1) - ply_thk - 5) 'Dimension plywood
//         For i = 1 To n_studs_y
//             Call DrawDimLin(pt9(0) - stud_spacing_y(UBound(stud_spacing_y) - i + 1), pt9(1) + 1.5, pt9(0) - ply_width_y, pt9(1), ((pt9(0) - ply_width_y) - (pt9(0) - stud_spacing_y(UBound(stud_spacing_y) - i + 1))) / 2, pt9(1) + 4 + i * 4, 0)   'Dimension studs from rightmost (in plan) stud
//         Next i
//         Call DrawDimLin(pt9(0) - stud_spacing_y(UBound(stud_spacing_y)) - 3.5, pt9(1) + 1.5, pt9(0) - ply_width_y, pt9(1), ((pt9(0) - ply_width_y) - (pt9(0) - stud_spacing_y(UBound(stud_spacing_y)) - 3.5)) / 2, pt9(1) + 4, 0)   'Dimension leftmost (in plan) stud to face of ply
//         Call DrawDimLin(pt9(0) - stud_start_offset_y, pt9(1) + 1.5, pt9(0), pt9(1), ((pt9(0)) - (pt9(0) - stud_start_offset_y)) / 2, pt9(1) + 4 * (n_studs_y + 1), 0)    'dimension stud_start_gap
//
//         'Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt9, "VBA_TOP_SECTION_DETAILS1", 1#, 1#, 1#, 0) 'Insert text notes, blow up detail at top section
//     
//     'Dim plywood
//         'Side A
//         PlyTemp = ptA(1) + ply_seams(1)
//         Call DrawDimSuffix(ptA(0) + ply_width_x, ptA(1), ptA(0) + ply_width_x, PlyTemp, ptA(0) + ply_width_x + 6, (PlyTemp - ptA(1)) / 2, " PLYWOOD")
//         If UBound(ply_seams) >= 2 Then 'If more than one ply panel, draw the rest, bottom up
//             For i = 2 To UBound(ply_seams)
//                 PlyTemp = PlyTemp + ply_seams(i)
//                 Call DrawDimSuffix(ptA(0) + ply_width_x, PlyTemp - ply_seams(i), ptA(0) + ply_width_x, PlyTemp, ptA(0) + ply_width_x + 6, (PlyTemp - (PlyTemp - ply_seams(i))) / 2, " PLYWOOD")
//             Next i
//         End If
//         Call DrawDimSuffix(ptA(0), ptA(1), ptA(0), ptA(1) + z, ptA(0) - 9, (2 * ptA(1) + z) / 2, " OVERALL HEIGHT")
//         Call DrawDimSuffix(ptA(0), ptA(1) + stud_base_gap, ptA(0), ptA(1) + z, ptA(0) - 4.5, ((ptA(1) + z) + (ptA(1) + stud_base_gap)) / 2, " STUD")
//         Call DrawDimSuffix(ptA(0) + ply_width_x, ptA(1) + z, ptA(0), ptA(1) + z, (2 * ptA(0) + ply_width_x) / 2, ptA(1) + z + 4, " PLYWOOD")
//         
//         'Side B
//         PlyTemp = ptB(1) + ply_seams(1)
//         Call DrawDimSuffix(ptB(0) + ply_width_y, ptB(1), ptB(0) + ply_width_y, PlyTemp, ptB(0) + ply_width_y + 6, (PlyTemp - ptB(1)) / 2, " PLYWOOD")
//         If UBound(ply_seams) >= 2 Then 'If more than one ply panel, draw the rest, bottom up
//             For i = 2 To UBound(ply_seams)
//                 PlyTemp = PlyTemp + ply_seams(i)
//                 Call DrawDimSuffix(ptB(0) + ply_width_y, PlyTemp - ply_seams(i), ptB(0) + ply_width_y, PlyTemp, ptB(0) + ply_width_y + 6, (PlyTemp - (PlyTemp - ply_seams(i))) / 2, " PLYWOOD")
//             Next i
//         End If
//         Call DrawDimSuffix(ptB(0), ptB(1), ptB(0), ptB(1) + z, ptB(0) - 9, (2 * ptB(1) + z) / 2, " OVERALL HEIGHT")
//         Call DrawDimSuffix(ptB(0), ptB(1) + stud_base_gap, ptB(0), ptB(1) + z, ptB(0) - 4.5, ((ptB(1) + z) + (ptB(1) + stud_base_gap)) / 2, " STUD")
//         Call DrawDimSuffix(ptB(0) + ply_width_y, ptB(1) + z, ptB(0), ptB(1) + z, (2 * ptB(0) + ply_width_y) / 2, ptB(1) + z + 4, " PLYWOOD")
//
//
//
//             
//     'PLAN VIEWS
//     pt13(0) = pt_o(0) + 427 + x - chamf_thk: pt13(1) = pt_o(1) + 42 + ply_thk 'origin point for plan views (lower right corner)
//
//     'Draw clamps
//     pt15(0) = pt13(0) - x - 2.25: pt15(1) = pt13(1) - 2.25
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt15, clamp_block, 1#, 1#, 1#, 0) 'Insert clamp (operator typical clamps)
//     Call ChangeProp(BlockRefObj, "Distance1", x + 4.5)
//     Call ChangeProp(BlockRefObj, "Distance2", y + 4.5)
//     
//     'Draw plywood
//     pt12(0) = pt13(0) - x: pt12(1) = pt13(1) - ply_thk
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY", 1#, 1#, 1#, 0) 'Bottom ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
//     
//     pt12(0) = pt12(0) - ply_thk: pt12(1) = pt12(1) + ply_width_y + ply_thk - 2.25
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY", 1#, 1#, 1#, -pi / 2) 'Left ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
//     
//     pt12(0) = pt12(0) + ply_thk + ply_width_x: pt12(1) = pt12(1) + ply_thk - 2.25
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY", 1#, 1#, 1#, pi) 'Top ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_x)
//     
//     pt12(0) = pt12(0) + ply_thk: pt12(1) = pt12(1) - ply_thk - ply_width_y + 2.25
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt12, "VBA_PLY", 1#, 1#, 1#, pi / 2) 'Right ply
//     Call ChangeProp(BlockRefObj, "Distance1", ply_width_y)
//
//
//     'Draw studs
//     pt14(1) = pt13(1) - ply_thk - 1.5
//     For i = 1 To n_studs_x
//         pt14(0) = pt13(0) - x + stud_spacing_x(i)
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block, 1#, 1#, 1#, 0) 'Insert bottom studs
//     Next i
//     pt14(0) = pt13(0) - x + stud_spacing_x(UBound(stud_spacing_x))
//
//     pt14(0) = pt13(0) - x - ply_thk - 1.5
//     For i = 1 To n_studs_y
//         pt14(1) = pt13(1) + y + 2.25 - stud_spacing_y(i)
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block, 1#, 1#, 1#, -pi / 2) 'Insert left studs
//     Next i
//     
//     pt14(1) = pt13(1) + y + ply_thk + 1.5
//     For i = 1 To n_studs_x
//         pt14(0) = pt13(0) - stud_spacing_x(i)
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block, 1#, 1#, 1#, pi) 'Insert top studs
//     Next i
//     
//     pt14(0) = pt13(0) + ply_thk + 1.5
//     For i = 1 To n_studs_y
//         pt14(1) = pt13(1) - 2.25 + stud_spacing_y(i)
//         Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, stud_block, 1#, 1#, 1#, pi / 2) 'Insert right studs
//     Next i
//
//     'Draw elevation tag
//     pt14(0) = pt13(0) - x / 2: pt14(1) = pt13(1) - 12
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, "VBA_ELEV_A", 1#, 1#, 1#, 0)
//     
//     'Draw hatch
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("CONC")
//     Dim hatchObj As AcadHatch
//     Dim OuterLoop(0) As AcadEntity
//     Dim pl_pts(0 To 9) As Double
//     
//     'Set boundry of hatch (must close hatch by returning to initial point)
//     pl_pts(0) = pt13(0):     pl_pts(1) = pt13(1)
//     pl_pts(2) = pt13(0) - x: pl_pts(3) = pt13(1)
//     pl_pts(4) = pt13(0) - x: pl_pts(5) = pt13(1) + y
//     pl_pts(6) = pt13(0):     pl_pts(7) = pt13(1) + y
//     pl_pts(8) = pt13(0):     pl_pts(9) = pt13(1)
//     
//     Set OuterLoop(0) = ThisDrawing.ModelSpace.AddLightWeightPolyline(pl_pts) 'Assign the pline object as the hatch's outer loop (inner loop optional and skipped here)
//     Set hatchObj = ThisDrawing.ModelSpace.AddHatch(0, "AR-CONC", True) 'Set hatch
//     hatchObj.AppendOuterLoop (OuterLoop)
//     OuterLoop(0).Delete
//     
//     'Draw chamfer in each corner
//     pt14(0) = pt13(0): pt14(1) = pt13(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, "VBA_3-4_CHAMFER", 1#, 1#, 1#, 0) 'Lower right
//
//     pt14(0) = pt13(0) - x: pt14(1) = pt13(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, "VBA_3-4_CHAMFER", 1#, 1#, 1#, -pi / 2) 'Lower left
//     
//     pt14(0) = pt13(0) - x: pt14(1) = pt13(1) + y
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, "VBA_3-4_CHAMFER", 1#, 1#, 1#, pi) 'Upper left
//     
//     pt14(0) = pt13(0): pt14(1) = pt13(1) + y
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt14, "VBA_3-4_CHAMFER", 1#, 1#, 1#, pi / 2) 'Upper right
//     
//     
//     'Draw text in plan view
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
//     pt17(0) = pt13(0) - x: pt17(1) = pt13(1) + y - 0.5
//     Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, x, "SIDE A")
//     MTextObject1.StyleName = "Arial Narrow": MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75 * Abs(Round(x, 3) >= 12) + (1.75 - 0.125 * (12 - x)) * Abs(Round(x, 3) < 12) 'Uses a text size of 1.75 when the x side of the column is >= 12", text size is reduced by 0.125" for every 1" the x side is below 12"
//
//     pt17(0) = pt13(0) - x + 0.5: pt17(1) = pt13(1)
//     Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, y, "SIDE B")
//     MTextObject1.StyleName = "Arial Narrow": MTextObject1.Rotation = pi / 2: MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75 * Abs(Round(y, 3) >= 12) + (1.75 - 0.125 * (12 - y)) * Abs(Round(y, 3) < 12) 'Uses a text size of 1.75 when the y side of the column is >= 12", text size is reduced by 0.125" for every 1" the y side is below 12"
//     
//     pt17(0) = pt13(0) - x: pt17(1) = pt13(1) + 2.25
//     Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, x, x)
//     MTextObject1.StyleName = "Arial Narrow": MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75 * Abs(Round(x, 3) >= 12) + (1.75 - 0.125 * (12 - x)) * Abs(Round(x, 3) < 12) 'Uses a text size of 1.75 when the x side of the column is >= 12", text size is reduced by 0.125" for every 1" the x side is below 12"
//     
//     pt17(0) = pt13(0) - 2.25: pt17(1) = pt13(1)
//     Set MTextObject1 = ThisDrawing.ModelSpace.AddMText(pt17, y, y)
//     MTextObject1.StyleName = "Arial Narrow": MTextObject1.Rotation = pi / 2: MTextObject1.AttachmentPoint = acAttachmentPointTopCenter: MTextObject1.Height = 1.75 * Abs(Round(y, 3) >= 12) + (1.75 - 0.125 * (12 - y)) * Abs(Round(y, 3) < 12) 'Uses a text size of 1.75 when the y side of the column is >= 12", text size is reduced by 0.125" for every 1" the y side is below 12"
//
//
//     'Add stud holdback detail and callout
//     'Detail on top of sheet
//     pt_blk(0) = ptB(0) + (x + 4.5) / 2 - 14: pt_blk(1) = pt_o(1) + 315
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_STUD_HOLDBACK_DETAIL", 1, 1, 1, 0)
//     
//     'Callout on panel
//     pt_blk(0) = ptB(0) + x + 4.5: pt_blk(1) = ptB(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_STUD_HOLDBACK_CALLOUT", 1, 1, 1, 0)
//     
//     'Add nailing note
//     pt_blk(0) = ptB(0) + x + 0.375: pt_blk(1) = ptB(1) + 0.95 * ply_seams(1)
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_NAILING_NOTES_SCISSOR", 1, 1, 1, 0)
//     
//     'Add detail references to each elevation view
//     'Plan view detail reference
//     pt_blk(0) = pt13(0) - ply_width_x / 2: pt_blk(1) = pt_o(1) + 9
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
//     AttList = BlockRefObj.GetAttributes
//     For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
//         If AttList(i).TextString = "X" Then
//             AttList(i).TextString = "D"
//         End If
//         If AttList(i).TextString = "DESCRIPTION" Then
//             AttList(i).TextString = "PLAN VIEW"
//         End If
//         If AttList(i).TextString = "REFERENCE" Then
//             AttList(i).TextString = ""
//         End If
//     Next i
//     
//     'Side A detail reference
//     pt_blk(0) = ptA(0) + ply_width_x / 2: pt_blk(1) = pt_o(1) + 9
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
//     AttList = BlockRefObj.GetAttributes
//     For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
//         If AttList(i).TextString = "X" Then
//             AttList(i).TextString = "C"
//         End If
//         If AttList(i).TextString = "DESCRIPTION" Then
//             AttList(i).TextString = "SIDE ""A"" PANEL"
//         End If
//         If AttList(i).TextString = "REFERENCE" Then
//             AttList(i).TextString = "VIEWED FROM PLYWOOD FACE"
//         End If
//     Next i
//     
//     'Side B detail reference
//     pt_blk(0) = ptB(0) + ply_width_y / 2: pt_blk(1) = pt_o(1) + 9
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
//     AttList = BlockRefObj.GetAttributes
//     For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
//         If AttList(i).TextString = "X" Then
//             AttList(i).TextString = "B"
//         End If
//         If AttList(i).TextString = "DESCRIPTION" Then
//             AttList(i).TextString = "SIDE ""B"" PANEL"
//         End If
//         If AttList(i).TextString = "REFERENCE" Then
//             AttList(i).TextString = "VIEWED FROM PLYWOOD FACE"
//         End If
//     Next i
//         
//     'Elevation view detail reference
//     pt_blk(0) = ptE(0) + ply_width_e / 2: pt_blk(1) = pt_o(1) + 9
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_blk, "VBA_DETAIL_REF", 0.75, 0.75, 0.75, 0)
//     AttList = BlockRefObj.GetAttributes
//     For i = LBound(AttList) To UBound(AttList) 'Assign attribute values
//         If AttList(i).TextString = "X" Then
//             AttList(i).TextString = "A"
//         End If
//         If AttList(i).TextString = "DESCRIPTION" Then
//             AttList(i).TextString = "COLUMN FORM ELEVATION"
//         End If
//         If AttList(i).TextString = "REFERENCE" Then
//             AttList(i).TextString = ""
//         End If
//     Next i
//     
//     'Add lower left fab count below "column form elevation"
//     ThisDrawing.ActiveLayer = ThisDrawing.Layers("FORM")
//     count_str = "FAB " & n_col & "-EA"
//     pt32(0) = ptE(0) + ply_width_e / 2: pt32(1) = pt_o(1) + 7
//     Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt32, 30, count_str)
//     MTextObject2.Height = 2.5
//
//
//     
// '#####################################################################################
// '                                S H E E T I N G
// '#####################################################################################
//    
//     Dim ptVP1(0 To 2) As Double
//     Dim ptVP2(0 To 2) As Double
//     Dim pt35(0 To 2) As Double
//     Dim AreaName As String
//     Dim oprop As AcadDynamicBlockReferenceProperty
//     
//     'Insert the McClone area block
//     AreaName = CStr(ColumnCreator.sAreaBox.Value)
//     pt35(0) = pt_o(0) + 548.37926872: pt35(1) = pt_o(1) + 186: pt35(2) = 0
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt35, "VBA_OFFICE_LOCATION", (2 / 3), (2 / 3), (2 / 3), 0)
//     
//     Props = BlockRefObj.GetDynamicBlockProperties
//     For i = LBound(Props) To UBound(Props)
//         Set oprop = Props(i) 'Get object properties from border block
//         If oprop.PropertyName = "Visibility1" Then 'Finds visibility state property
//             oprop.Value = ColumnCreator.sAreaBox.Value
//         End If
//     Next i
//     
//     'Insert border
//     Set BlockRefObj = ThisDrawing.ModelSpace.InsertBlock(pt_click, "VBA_24X36_LAYOUT", 1#, 1#, 1#, 0) 'Insert 3/4" = 1'-0" sheet boundry
//     BlockRefObj.ScaleEntity pt_click, 12 / 0.75
//     
//     'Sets Sheet issued for: "" text (defined as a block reference property)
//     Props = BlockRefObj.GetDynamicBlockProperties
//     For i = LBound(Props) To UBound(Props)
//         Set oprop = Props(i) 'Get object properties from border block
//         If oprop.PropertyName = "Visibility1" Then 'Finds visibility state property
//             oprop.Value = ColumnCreator.sSheetIssuedForBox.Value
//         End If
//     Next i
//     
//     'Place timestamp in lower left corner, outside of sheet boundry
//     pt34(0) = pt_o(0) - 1.65: pt34(1) = pt_o(1): pt34(2) = pt_o(2)
//     ThisDrawing.ActiveTextStyle = ThisDrawing.TextStyles("VBA_txt_narrow")
//     Set MTextObject2 = ThisDrawing.ModelSpace.AddMText(pt34, 18, Now())
//     MTextObject2.Height = 1.25
//     MTextObject2.Rotate pt34, pi / 2
//     
//     'Places other text (defined as attributes)
//     AttList = BlockRefObj.GetAttributes
//     For i = LBound(AttList) To UBound(AttList)
//         ' Check for the correct attribute tag.
//         If AttList(i).TextString = "SHEET NAME" Then
//             AttList(i).TextString = "" & ColumnCreator.sProjectNameBox.Value
//         End If
//         If AttList(i).TextString = "PROJECT TITLE" Then
//             AttList(i).TextString = "" & ColumnCreator.sProjectTitleBox.Value
//         End If
//         If AttList(i).TextString = "PROJECT ADDRESS" Then
//             AttList(i).TextString = "" & ColumnCreator.sProjectAddressBox.Value
//         End If
//         If AttList(i).TextString = "DATE" Then
//             AttList(i).TextString = "" & ColumnCreator.sDateBox.Value
//         End If
//         If AttList(i).TextString = "SCALE" Then
//             AttList(i).TextString = "3/4'' = 1'-0''"
//         End If
//         If AttList(i).TextString = "DRAWN BY" Then
//             AttList(i).TextString = "" & ColumnCreator.sDrawnByBox.Value
//         End If
//         If AttList(i).TextString = "JOB NUMBER" Then
//             AttList(i).TextString = "" & ColumnCreator.sJobNoBox.Value
//         End If
//         If AttList(i).TextString = "SHEET_#" Then
//             AttList(i).TextString = "" & ColumnCreator.sSheetBox.Value
//         End If
//         If AttList(i).TextString = "SUFFIX" Then
//             AttList(i).TextString = "" & ColumnCreator.sSuffixBox.Value
//         End If
//     Next i
//
//
// last_line:
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
// ErrorHandler:
//     If Err.Number = -2147352567 Then 'This error occurs when the user presses the ESC key while inside the GetPoint utility
//         ColumnCreator.show
//         Exit Function
//     ElseIf Err.Number = -2145386422 Then 'This error occurs when the path to Egnyte isn't found
//         MsgBox ("Error: VBA block library not found at: " & FileToInsert & vbNewLine & vbNewLine & "Check Egnyte (Z:\) connection. Try closing and relaunching Egnyte Connect.")
//         Exit Function
//     ElseIf Err.Number = -2145386445 Then 'This error occurs when a block definition that doesn't exist is attempted to be inserted.
//         If BlockExists = True Then
//             MsgBox ("This drawing seems to be missing blocks for this VBA script. This is normal if you are using new button features on an old drawing. The block library will be re-downloaded and this script terminated." & vbNewLine & vbNewLine & "Please run the button again.")
//             'Insert the block library then delete it
//             FileToInsert = FileLocationPrefix & "VBA_Block_Library.dwg"
//             Set BlockDwg = ThisDrawing.ModelSpace.InsertBlock(insertionPnt, FileToInsert, 1#, 1#, 1#, 0)
//             BlockDwg.Delete
//             Exit Function
//         Else
//             MsgBox ("Error: This drawing seems to be missing blocks for this VBA script despite the block library having been downloaded. The library may be missing the necessary blocks, or something mysterious has gone horribly wrong." & vbNewLine & vbNewLine & "Library location: " & FileToInsert)
//             Exit Function
//         End If
//     Else 'Generic error message
//         MsgBox ("An error occured." & vbNewLine & "error " & Err.Number & vbNewLine & Err.Description)
//         Exit Function
//     End If
//     
// SkipErrorHandler:
//
// End Function
//
//
//
// Public Function ChangeProp(BlockRefObj, PropName, PropValue)
//
//     Props = BlockRefObj.GetDynamicBlockProperties
//     Dim oprop As AcadDynamicBlockReferenceProperty
//     
//     For i = LBound(Props) To UBound(Props)
//         Set oprop = Props(i) 'Get object properties from block one at a time
//         If oprop.PropertyName = PropName Then 'compare to property name
//             oprop.Value = PropValue 'if name matches, assign new PropValue
//         End If
//     'Exit For
//     Next i
// End Function
// Public Function DrawDim(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double)
//
//     Dim dimObj As AcadDimAligned
//     Dim point1(0 To 2) As Double 'first point
//     Dim point2(0 To 2) As Double 'second point
//     Dim location(0 To 2) As Double 'text location
//     
//     ' Define the dimension
//     point1(0) = x1#:    point1(1) = y1#:    point1(2) = 0#
//     point2(0) = x2#:    point2(1) = y2#:    point2(2) = 0#
//     location(0) = x3#:  location(1) = y3#:  location(2) = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimAligned(point1, point2, location)
// End Function
// Public Function DrawDimSuffix(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, txt As String)
//
//     Dim dimObj As AcadDimAligned
//     Dim point1(0 To 2) As Double
//     Dim point2(0 To 2) As Double
//     Dim location(0 To 2) As Double
//     
//     ' Define the dimension
//     point1(0) = x1#:    point1(1) = y1#:    point1(2) = 0#
//     point2(0) = x2#:    point2(1) = y2#:    point2(2) = 0#
//     location(0) = x3#:  location(1) = y3#:  location(2) = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimAligned(point1, point2, location)
//     dimObj.TextSuffix = txt
//     
// End Function
// Public Function DrawDimLin(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, angle As Double)
//
//     Dim dimObj As AcadDimRotated
//     Dim point1(0 To 2) As Double 'first point
//     Dim point2(0 To 2) As Double 'second point
//     Dim location(0 To 2) As Double 'text location
//     
//     ' Define the dimension
//     point1(0) = x1#:    point1(1) = y1#:    point1(2) = 0#
//     point2(0) = x2#:    point2(1) = y2#:    point2(2) = 0#
//     location(0) = x3#:  location(1) = y3#:  location(2) = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimRotated(point1, point2, location, angle)
// End Function
// Public Function DrawDimLinSuffixLeader(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, txt As String, angle As Double)
//
//     Dim dimObj As AcadDimRotated
//     Dim point1(0 To 2) As Double
//     Dim point2(0 To 2) As Double
//     Dim location(0 To 2) As Double
//     Dim txtLocation(0 To 2) As Double
//     
//     ' Define the dimension
//     point1(0) = x1#: point1(1) = y1#: point1(2) = 0#
//     point2(0) = x2#: point2(1) = y2#: point2(2) = 0#
//     location(0) = x3#: location(1) = y3#: location(2) = 0#
//     txtLocation(0) = x4#: txtLocation(1) = y4#: txtLocation(2) = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimRotated(point1, point2, location, angle)
//     dimObj.TextMovement = acMoveTextAddLeader
//     dimObj.TextPosition = txtLocation
//     dimObj.TextSuffix = txt
// End Function
// Public Function DrawDimLinLeader(x1 As Double, y1 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double, x4 As Double, y4 As Double, angle As Double)
//
//     Dim dimObj As AcadDimRotated
//     Dim point1(0 To 2) As Double
//     Dim point2(0 To 2) As Double
//     Dim location(0 To 2) As Double
//     Dim txtLocation(0 To 2) As Double
//     
//     ' Define the dimension
//     point1(0) = x1#: point1(1) = y1#: point1(2) = 0#
//     point2(0) = x2#: point2(1) = y2#: point2(2) = 0#
//     location(0) = x3#: location(1) = y3#: location(2) = 0#
//     txtLocation(0) = x4#: txtLocation(1) = y4#: txtLocation(2) = 0#
//     
//     ' Create an aligned dimension object in model space
//     Set dimObj = ThisDrawing.ModelSpace.AddDimRotated(point1, point2, location, angle)
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
// Public Function DrawRectangle(xA As Double, yA As Double, xB As Double, yB As Double)
//     Dim plineObj As AcadLWPolyline
//     Dim pts(0 To 7) As Double
//     
//     pts(0) = xA#:   pts(1) = yA
//     pts(2) = xB#:   pts(3) = yA
//     pts(4) = xB#:   pts(5) = yB
//     pts(6) = xA#:   pts(7) = yB
//
//     Set plineObj = ThisDrawing.ModelSpace.AddLightWeightPolyline(pts)
//     plineObj.Closed = True
//
// End Function
// Public Function FindMax(Arr As Variant)
//     Dim RetMax As Double
//     RetMax = Arr(LBound(Arr)) 'initialize with first value
//     
//     For i = LBound(Arr) To UBound(Arr)
//     If Arr(i) > RetMax Then
//         RetMax = Arr(i)
//     End If
//     Next i
//     FindMax = RetMax
// End Function