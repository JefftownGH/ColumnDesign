using System;
using System.Linq;
using System.Windows;
using ColumnDesign.UI;
using ColumnDesign.ViewModel;
using static ColumnDesign.Modules.ConvertFeetInchesToNumber;
using static ColumnDesign.Modules.ConvertNumberToFeetInches;
using static ColumnDesign.Modules.GetPlySeamsFunction;
using static ColumnDesign.Modules.ReadSizesFunction;

namespace ColumnDesign.Modules
{
    public class sUpdatePly_Function
    {
        public static void sUpdatePly(ColumnCreatorView ui, ColumnCreatorViewModel vm)
        {
            var x = ConvertToNum(vm.SWidthX);
            var y = ConvertToNum(vm.SLengthY);
            var z = ConvertToNum(vm.SHeightZ);
            if (x == 0 || y == 0 || z == 0) return;
            double[] ply_seams = { };
            double temp_ht = 0;
            if (ui.SBoxPlySeams.Text.Equals(""))
            {
                ply_seams = GetPlySeams(x, y, z, ui);
                ui.STxtPlyError.Visibility = Visibility.Collapsed;
            }
            else
            {
                ply_seams = ReadSizes(ui.SBoxPlySeams.Text);
                if (sValidatePlySeams(ui, ply_seams, x, y, z) != 1)
                {
                    return;
                }
            }

            if (!ui.BoxPlySeams.Text.Equals(""))
            {
                goto SkipPlyUpdate;
            }

            string strPlySeams = "";
            for (var i = 0; i < ply_seams.Length; i++)
            {
                if (i != ply_seams.Length - 1)
                {
                    strPlySeams += ConvertFtIn(ply_seams[i]) + ", ";
                }
                else
                {
                    strPlySeams += ConvertFtIn(ply_seams[i]);
                }
            }

            ui.SBoxPlySeams.Text = strPlySeams;

            SkipPlyUpdate:
            double[] pt_draw = new double[3];
            pt_draw[0] = 200;
            pt_draw[1] = 216;
            pt_draw[2] = 0;
            double px_z;
            double px_x;
            double px_y;
            double px_z_max;
            if (z > 96)
            {
                px_z_max = 200;
                px_x = px_z_max * (x / z);
                px_y = px_z_max * (y / z);
            }
            else
            {
                px_z_max = 200 * z / 96;
                px_x = px_z_max * (x / z) * z / 96;
                px_y = px_z_max * (y / z) * z / 96;
            }

            double px_width = ui.SSlblAxis.Text switch
            {
                "X" => px_x,
                "Y" => px_y,
                _ => throw new ArgumentException("Error: Horizontal dimension not x or y")
            };
            //TODO  'Draw "frame" of column: left, right, and top

//         'Left
//         ColumnCreator.s_img_line_1.Left = pt_draw(0) - px_width / 2
//         ColumnCreator.s_img_line_1.top = pt_draw(1) - px_z_max
//         ColumnCreator.s_img_line_1.Width = 2
//         ColumnCreator.s_img_line_1.Height = px_z_max
//         ColumnCreator.s_img_line_1.Visible = True
//     
//         'Right
//         ColumnCreator.s_img_line_2.Left = pt_draw(0) + px_width / 2
//         ColumnCreator.s_img_line_2.top = pt_draw(1) - px_z_max
//         ColumnCreator.s_img_line_2.Width = 2
//         ColumnCreator.s_img_line_2.Height = px_z_max
//         ColumnCreator.s_img_line_2.Visible = True
//     
//         'Top
//         ColumnCreator.s_img_line_3.Left = pt_draw(0) - px_width / 2
//         ColumnCreator.s_img_line_3.top = pt_draw(1) - px_z_max
//         ColumnCreator.s_img_line_3.Width = px_width
//         ColumnCreator.s_img_line_3.Height = 2
//         ColumnCreator.s_img_line_3.Visible = True
//         
//         'Draw plywood seams
//         Dim cumulative_z As Double
//         For i = 1 To UBound(ply_seams) - 1
//             Coll(i).Left = pt_draw(0) - px_width / 2
//             cumulative_z = 0
//             For j = 1 To i
//                 cumulative_z = cumulative_z + ply_seams(j)
//             Next j
//             Coll(i).top = pt_draw(1) - px_z_max * (cumulative_z / z)
//             Coll(i).Width = px_width
//             Coll(i).Height = 2
//             Coll(i).Visible = True
//         Next i
//         
//         'Dimension plywood seams
//         For i = 1 To UBound(ply_seams)
//             Coll(i + 10).Left = pt_draw(0) + px_width / 2 + 10
//             cumulative_z = 0
//             For j = 1 To i
//                 cumulative_z = cumulative_z + ply_seams(j)
//             Next j
//             Coll(i + 10).top = pt_draw(1) - px_z_max * ((cumulative_z - ply_seams(i) / 2) / z) - Coll(i + 10).Height / 2 + 2
//             Coll(i + 10).Caption = ConvertFtIn(ply_seams(i))
//             Coll(i + 10).Visible = True
//         Next i
//     End If
        }

        public static int sValidatePlySeams(ColumnCreatorView ui, double[] ply_seams, double x, double y,
            double z)
        {
            var temp_ht = ply_seams.Sum();
            if (Math.Abs(temp_ht - z) > 0.001)
            {
                if (temp_ht > z)
                {
                    ui.STxtPlyError.Text = $"Remove {ConvertFtIn(temp_ht - z)}";
                }
                else if (temp_ht < z)
                {
                    ui.STxtPlyError.Text = $"Add {ConvertFtIn(z - temp_ht)}";
                }

                ui.STxtPlyError.Visibility = Visibility.Visible;
                return 0;
            }

            double ply_width_x;
            double ply_width_y;
            double min_ply_ht = 6;
            double max_ply_ht;
            ply_width_x = x;
            ply_width_y = y + 4.5;
            if (ply_width_x > 48 || ply_width_y > 48) max_ply_ht = 48;
            else max_ply_ht = 96;
            foreach (var t in ply_seams)
            {
                if (t > max_ply_ht)
                {
                    ui.STxtPlyError.Text = $"{ConvertFtIn(t)} ply too tall. Max: {ConvertFtIn(max_ply_ht)}";
                    ui.STxtPlyError.Visibility = Visibility.Visible;
                    return 0;
                }

                if (t < min_ply_ht)
                {
                    ui.STxtPlyError.Text = $"{ConvertFtIn(t)} ply too small. Min: {ConvertFtIn(min_ply_ht)}";
                    ui.STxtPlyError.Visibility = Visibility.Visible;
                    return 0;
                }

                ui.STxtPlyError.Visibility = Visibility.Collapsed;
            }

            ui.STxtPlyError.Visibility = Visibility.Collapsed;
            return 1;
        }

        public static void sCheckHeight(double x, double y, double z, ColumnCreatorView ui)
        {
            double big_dim;
            double max_ht;
            if (x == 0 || y == 0||z==0)
            {
                ui.SMaxHeight.Text = "N/A";
                return;
            }

            if (x >= y)
            {
                big_dim = x;
            }
            else
            {
                big_dim = y;
            }

            max_ht = big_dim switch
            {
                <= 8 => 196,
                <= 10 => 196,
                <= 12 => 196,
                <= 14 => 196,
                <= 16 => 196,
                <= 18 => 197,
                <= 20 => 197,
                <= 22 => 193,
                <= 24 => 195,
                <= 26 => 175,
                <= 28 => 175,
                <= 30 => 175,
                <= 32 => 133,
                <= 34 => 133,
                <= 36 => 180,
                <= 38 => 127,
                <= 40 => 127,
                <= 42 => 127,
                <= 44 => 127,
                <= 46 => 127,
                _ => 127
            };
            ui.SMaxHeight.Text = ConvertFtIn(max_ht);
        }
    }
}