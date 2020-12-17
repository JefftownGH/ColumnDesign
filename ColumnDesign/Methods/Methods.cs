using System;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ColumnDesign.UI;
using static ColumnDesign.Modules.ConvertFeetInchesToNumber;
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
            WinGap = 6    ;
            WinStudOff = 1;
            int[,] stud_matrix;
            var ply_thk = 0.75;
            var chamf_thk = 0.75;
            var min_stud_gap = 2.125;
            var max_wt_single_top  = 2000 ;
            var stud_base_gap  = 0.25;
            double bot_clamp_gap  = 8;
            var stud_start_offset  = 0.125;
            
            x = ConvertToNum(_ui.WidthX.Text);
            y =  ConvertToNum(_ui.LengthY.Text);
            z =  ConvertToNum(_ui.HeightZ.Text);
            n_col = (int) ConvertToNum(_ui.Quantity.Text);
            // ply_name = ColumnCreator.PlyNameBox
            if (WinX)
            {
                x= ConvertToNum(_ui.LengthY.Text);
                y=ConvertToNum(_ui.WidthX.Text);
                WinX = false;
                WinY = true;
            }
            //TODO other  code
                
            // #####################################################################################
            //                                                                                           S T U D S
            // #####################################################################################
            var row_num = 0;
            int col_num_x = 0;
            int col_num_y = 0;
            int n_studs_total;
            stud_matrix = ImportMatrix(@$"{GlobalNames.WtFileLocationPrefix}Columns\n_stud_matrix.csv");
            for (var i = 0; i < stud_matrix.GetLength(0); i++)
            {
                if (stud_matrix[i,0]>z)
                {
                    row_num = i;
                    break;
                }
            }

            for (var i = 0; i < stud_matrix.GetLength(1); i++)
            {
                if (stud_matrix[0,i]==x)
                {
                    col_num_x = i;
                    break;
                }
            }

            for (var i = 0; i < stud_matrix.GetLength(1); i++)
            {
                if (stud_matrix[0,i]==y)
                {
                    col_num_y = i;
                    break;
                }
            }

            if (col_num_x==0)
            {
                for (var i = 0; i < stud_matrix.GetLength(1); i++)
                {
                    if (stud_matrix[0,i]>x)
                    {
                        if (stud_matrix[row_num,i]>=stud_matrix[row_num,i-1])
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

            if (col_num_y==0)
            {
                for (var i = 0; i < stud_matrix.GetLength(1); i++)
                {
                    if (stud_matrix[0,i]>y)
                    {
                        if (stud_matrix[row_num,i]>=stud_matrix[row_num,i-1])
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
           if (avg_gap_x<min_stud_gap)
           {
               push_studs_x = 1;
           }
           avg_gap_y = (y + ply_thk - stud_start_offset - n_studs_y * 3.5) / (n_studs_y - 1);
           if (avg_gap_y<min_stud_gap)
           {
               push_studs_y = 1;
           }
           double[] stud_spacing_x=new double[n_studs_x-1];
           double[] stud_spacing_y=new double[n_studs_y-1];
           double available_2x2_x;
           double available_2x2_y;
           double min_2x2_gap = 1.625;
           int AFB_x2;
           for (var i = 0; i < n_studs_x; i++)
           {
               stud_spacing_x[i] = stud_start_offset + (i - 1) * (3.5 + avg_gap_x);
           }

           for (var i = 0; i < n_studs_y; i++)
           {
               stud_spacing_y[i] = stud_start_offset + (i - 1) * (3.5 + avg_gap_y);
           }

           if (push_studs_x==1)
           {
               if (avg_gap_x * (n_studs_x - 1) >= 2 * min_stud_gap)
               {
                   stud_spacing_x[1] = stud_start_offset + 3.5 + min_stud_gap;
                   stud_spacing_x[stud_spacing_x.Length - 2] = stud_spacing_x[stud_spacing_x.Length - 1] - 3.5 - min_stud_gap;
                   AFB_x2 = 1;
               }
               else
               {
                   stud_spacing_x[stud_spacing_x.Length - 2] = stud_spacing_x[stud_spacing_x.Length - 1] - 3.5 - min_stud_gap;
                   AFB_x2 = 0;
               }
           }

           if (push_studs_y==1)
           {
               if (avg_gap_y * (n_studs_y - 1) >= 2 * min_stud_gap)
               {
                   stud_spacing_y[1] = stud_start_offset + 3.5 + min_stud_gap;
                   stud_spacing_y[stud_spacing_y.Length - 2] = stud_spacing_y[stud_spacing_y.Length-1] - 3.5 - min_stud_gap;
               }
               else
               {
                   stud_spacing_y[stud_spacing_y.Length - 2] = stud_spacing_y[stud_spacing_y.Length-1] - 3.5 - min_stud_gap;
               }
           }

           for (var i = 0; i < n_studs_x-1; i++)
           {
               if (stud_spacing_x[i+1]-stud_spacing_x[i]<3.5)
               {
                   stud_spacing_x[i] = stud_spacing_x[i + 1] - 3.5;
               }
           }
           for (var i = 0; i < n_studs_y-1; i++)
           {
               if (stud_spacing_y[i+1]-stud_spacing_y[i]<3.5)
               {
                   stud_spacing_y[i] = stud_spacing_y[i + 1] - 3.5;
               }
           }

           for (var i = 0; i < n_studs_x-1; i++)
           {
               if (stud_spacing_x[i+1]-stud_spacing_x[i]<3.5)
               {
                   TaskDialog.Show("Error","Error: Studs overlap, check and manually correct drawing.");
                   goto EndOfStudChecks;
               }
           }
           for (var i = 0; i < n_studs_y-1; i++)
           {
               if (stud_spacing_y[i+1]-stud_spacing_y[i]<3.5)
               {
                   TaskDialog.Show("Error","Error: Studs overlap, check and manually correct drawing.");
                   goto EndOfStudChecks;
               }
           }

           if (window)
           {
               for (var i = 0; i < n_studs_x-1; i++)
               {
                   if (stud_spacing_x[i+1]-stud_spacing_x[i]<5)
                   {
                       TaskDialog.Show("Error",
                           "Warning: Insufficient clearance between studs for a 2x2 window lock. Check and manually correct drawing if necessary.");
                       goto EndOfStudChecks;
                   }
               }
               for (var i = 0; i < n_studs_y-1; i++)
               {
                   if (stud_spacing_y[i+1]-stud_spacing_y[i]<5)
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
            //TODO PLYWOOD
            ply_width_x = x + 1.5;
            ply_width_y = y + 1.5;
            int n_studs_w;
            double ply_width_w=0;
            double[] stud_spacing_w = new double[0];
            int n_studs_e;
            double ply_width_e=0;
            double[] stud_spacing_e= new double[0];
            if (WinX)
            {
                n_studs_w = n_studs_x;
                n_studs_e = n_studs_x;
                ply_width_w = ply_width_x;
                ply_width_e = ply_width_x;
                Array.Resize(ref stud_spacing_w, stud_spacing_x.Length-1);
                Array.Resize(ref stud_spacing_e, stud_spacing_x.Length-1);
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
                Array.Resize(ref stud_spacing_w, stud_spacing_y.Length-1);
                Array.Resize(ref stud_spacing_e, stud_spacing_y.Length-1);
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
                Array.Resize(ref stud_spacing_e, stud_spacing_x.Length-1);
                for (var i = 0; i < stud_spacing_x.Length; i++)
                {
                    stud_spacing_e[i] = stud_spacing_x[i];
                }
            }

            if (ply_width_x>48||ply_width_y>48)
            {
                max_ply_ht = 48;
            }
            else
            {
                max_ply_ht = 96;
            }
            //TODO PLYWOOD

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
            var bigSide = new[] {x, y, z}.Max();
            if (_ui.RbColumn824.IsChecked == true || _ui.RbColumnCustom.IsChecked == true && bigSide <= 24)
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
            for (var i = 0; i < clamp_spacing.Length-1; i++)
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
            clamp_spacing[clamp_spacing.Length-1] = (int)bot_clamp_gap;
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
                }

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
            var clamp_spacing_con = new int[clamp_spacing.Length - 1];
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

                            clamp_spacing_con[i] = (int)WinPos - (int)win_clamp_bot_max;
                            clamp_spacing_con[clamp_spacing_con.Length - 1] = 0;
                            if (clamp_spacing_con[i + 1] - clamp_spacing_con[i] < 8)
                            {
                                clamp_spacing_con[i + 1] = (clamp_spacing_con[i + 2] + clamp_spacing_con[i]) / 2;
                            }

                            goto LowerWinClampSet;
                        }

                        if (WinPos - clamp_spacing_con[i] <= win_clamp_bot_max)
                        {
                            clamp_spacing_con[i] = (int) WinPos - (int)win_clamp_bot_max;
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
                                for (var l = 0; l < clamp_spacing_con.Length - j; l++)
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
                Array.Resize(ref clamp_spacing, n_clamps);
                clamp_spacing_con[0] = (int) z - clamp_spacing_con[0];
                for (var i = 1; i < n_clamps; i++)
                {
                    clamp_spacing_con[i] = clamp_spacing_con[i - 1] - clamp_spacing_con[i];
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
            if (_ui.WindowX.IsChecked == true || _ui.WindowY.IsChecked == true)
            {
                if (clamp_spacing_con[1] < WinPos)
                {
                    throw new Exception(
                        "Warning:\nPour window is only secured by 1 clamp. Consult with an engineering manager");
                }
            }
            
            // #####################################################################################
            //                                                                  M I S C E L L A N E O U S
            // #####################################################################################
            //TODO MISCELLANEOUS

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
            if (window==false && Math.Abs(x - y) < 0.001)
            {
                ptA[0] = pt_o[0] + 245;
                ptA[1] = pt_o[1] + 25;
                ptE[0] = pt_o[0] + 100;
                ptE[1] = pt_o[1] + 25;
            }
            else if (window==true && Math.Abs(x - y) < 0.001)
            {
                ptW[0] = pt_o[0] + 300;
                ptW[1] = pt_o[1] + 25;
                ptA[0] = pt_o[0] + 185;
                ptA[1] = pt_o[1] + 25;
                ptE[0] = pt_o[0] + 54.5;
                ptE[1] = pt_o[1] + 25;
                DrawW = true;
            }
            else if (window==false && Math.Abs(x - y) > 0.001)
            {
                ptA[0] = pt_o[0] + 300;
                ptA[1] = pt_o[1] + 25;
                ptB[0] = pt_o[0] + 185;
                ptB[1] = pt_o[1] + 25;
                ptE[0] = pt_o[0] + 54.5;
                ptE[1] = pt_o[1] + 25;
                DrawB = true;
            }
            else if (window==true && Math.Abs(x - y) > 0.001)
            {
                ptW[0] = pt_o[0] + 340 - 0.5 * ply_width_w;
                ptW[1] = pt_o[1] + 25;
                ptA[0] = pt_o[0] + 267.25 - 0.5 * ply_width_x;
                ptA[1] = pt_o[1] + 25;
                ptB[0] = pt_o[0] + 190 - 0.5 * ply_width_y;
                ptB[1] = pt_o[1] + 25;
                ptE[0] = pt_o[0] + 54.5 - 0.5 * (ply_width_w -24);
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