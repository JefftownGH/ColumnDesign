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
    public static class UpdatePly_Function
    {
        public static void UpdatePly(ColumnCreatorView ui, ColumnCreatorViewModel vm)
        {
            if (ui.WindowX.IsEnabled == false) return;
            var x = ConvertToNum(vm.WidthX);
            var y = ConvertToNum(vm.LengthY);
            var z = ConvertToNum(vm.HeightZ);

            double[] ply_seams = { };

            vm.WinDim1 = ConvertFtIn(z / 5);
            vm.WinDim2 = ConvertFtIn(z - ConvertToNum(vm.WinDim1));
            goto UpdatePlyColorCheck;

            UpdatePlyColorCheck:
            double temp_ht = 0;
            if (ui.BoxPlySeams.Text.Equals(""))
            {
                ply_seams = GetPlySeams(x, y, z, ui);
                ui.TxtPlyError.Visibility = Visibility.Collapsed;
            }
            else
            {
                ply_seams = ReadSizes(ui.BoxPlySeams.Text);
                if (ValidatePlySeams(ui, ply_seams, x, y, z) != 1)
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

            ui.BoxPlySeams.Text = strPlySeams;
            
            SkipPlyUpdate:
            double[] pt_draw = new double[3];
            pt_draw[0] = 200;
            pt_draw[1] = 216;
            pt_draw[2] = 0;
            const double px_z_max = 200;
            double px_x = px_z_max * (x / z);
            double px_y = px_z_max * (y / z);
            double px_width = vm.SlblAxis switch
            {
                "X" => px_x,
                "Y" => px_y,
                _ => throw new ArgumentException("Error: Horizontal dimension not x or y")
            };
            //TODO  'Draw "frame" of column: left, right, and top
            // ColumnCreator.img_line_1.Left = pt_draw(0) - px_width / 2
            // ColumnCreator.img_line_1.top = pt_draw(1) - px_z_max
            // ColumnCreator.img_line_1.Width = 2
            // ColumnCreator.img_line_1.Height = px_z_max
            // ColumnCreator.img_line_1.Visible = True
            //
            // 'Right
            // ColumnCreator.img_line_2.Left = pt_draw(0) + px_width / 2
            // ColumnCreator.img_line_2.top = pt_draw(1) - px_z_max
            // ColumnCreator.img_line_2.Width = 2
            // ColumnCreator.img_line_2.Height = px_z_max
            // ColumnCreator.img_line_2.Visible = True
            //
            // 'Top
            // ColumnCreator.img_line_3.Left = pt_draw(0) - px_width / 2
            // ColumnCreator.img_line_3.top = pt_draw(1) - px_z_max
            // ColumnCreator.img_line_3.Width = px_width
            // ColumnCreator.img_line_3.Height = 2
            // ColumnCreator.img_line_3.Visible = True
            //
            // 'Draw plywood seams
            //     Dim cumulative_z As Double
            // For i = 1 To UBound(ply_seams) - 1
            // Coll(i).Left = pt_draw(0) - px_width / 2
            // cumulative_z = 0
            // For j = 1 To i
            //     cumulative_z = cumulative_z + ply_seams(j)
            // Next j
            // Coll(i).top = pt_draw(1) - px_z_max * (cumulative_z / z)
            // Coll(i).Width = px_width
            // Coll(i).Height = 2
            // Coll(i).Visible = True
            // Next i
            //
            // 'Dimension plywood seams
            // For i = 1 To UBound(ply_seams)
            // Coll(i + 10).Left = pt_draw(0) + px_width / 2 + 10
            // cumulative_z = 0
            // For j = 1 To i
            //     cumulative_z = cumulative_z + ply_seams(j)
            // Next j
            // Coll(i + 10).top = pt_draw(1) - px_z_max * ((cumulative_z - ply_seams(i) / 2) / z) - Coll(i + 10).Height / 2 + 2
            // Coll(i + 10).Caption = ConvertFtIn(ply_seams(i))
            // Coll(i + 10).Visible = True
            // Next i
                
            if (ui.WindowX.IsChecked==true || ui.WindowX.IsChecked ==true )
            {
                if (vm.WinDim1.Equals("Z1"))
                {
                    vm.WinDim1 = ConvertFtIn(z / 5);
                    vm.WinDim2 = ConvertFtIn(z - ConvertToNum(vm.WinDim1));
                }

                if (ConvertToNum(vm.WinDim1) > z)
                {
                    vm.WinDim1 = ConvertFtIn(z);
                    vm.WinDim2 = ConvertFtIn(0);
                }
                else if (ConvertToNum(vm.WinDim2) > z)
                {
                    vm.WinDim2 = ConvertFtIn(z);
                    vm.WinDim1 = ConvertFtIn(0);
                }
            }
        }
        
        //     'Input source of 0 is for changes from the x, y, and z boxes or from changing the axis
//     'Input source of 1 is for changes from boxPlySeams
//     'Input source of 2 is for changes from the window checkboxes or the window position dimensions
//     'Input source of 3 is for changes from the z box only
//         
//         ColumnCreator.img_line_1.Left = pt_draw(0) - px_width / 2
//         ColumnCreator.img_line_1.top = pt_draw(1) - px_z_max
//         ColumnCreator.img_line_1.Width = 2
//         ColumnCreator.img_line_1.Height = px_z_max
//         ColumnCreator.img_line_1.Visible = True
//     
//         'Right
//         ColumnCreator.img_line_2.Left = pt_draw(0) + px_width / 2
//         ColumnCreator.img_line_2.top = pt_draw(1) - px_z_max
//         ColumnCreator.img_line_2.Width = 2
//         ColumnCreator.img_line_2.Height = px_z_max
//         ColumnCreator.img_line_2.Visible = True
//     
//         'Top
//         ColumnCreator.img_line_3.Left = pt_draw(0) - px_width / 2
//         ColumnCreator.img_line_3.top = pt_draw(1) - px_z_max
//         ColumnCreator.img_line_3.Width = px_width
//         ColumnCreator.img_line_3.Height = 2
//         ColumnCreator.img_line_3.Visible = True
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
//         
//         'Create window seam

//         If ColumnCreator.chkWinX.Value = True Or ColumnCreator.chkWinY.Value = True Then
//             If InputSource = 0 Or InputSource = 2 Or InputSource = 3 Then
//                 ColumnCreator.img_line_win.Left = pt_draw(0) - px_width / 2
//                 ColumnCreator.img_line_win.top = pt_draw(1) - px_z_max + ConvertToNum(ColumnCreator.WinDim1.Value) * px_z_max / z
//                 ColumnCreator.img_line_win.Width = px_width
//                 ColumnCreator.img_line_win.Height = 3 'Thicc
//                 ColumnCreator.WinDim1.Left = pt_draw(0) - px_width / 2 - 40
//                 ColumnCreator.WinDim1.top = pt_draw(1) - px_z_max + (ConvertToNum(ColumnCreator.WinDim1.Value) * px_z_max / z) / 2 - ColumnCreator.WinDim1.Height / 2
//                 'Move dimensions into place (lower dim)
//                 ColumnCreator.WinDim2.Left = pt_draw(0) - px_width / 2 - 40
//                 ColumnCreator.WinDim2.top = pt_draw(1) - (ConvertToNum(ColumnCreator.WinDim2.Value) * px_z_max / z) / 2 - ColumnCreator.WinDim1.Height / 2
// UpdatePlyColorCheck:
//             End If
//             ColumnCreator.img_line_win.Visible = True
//             ColumnCreator.WinDim1.Visible = True
//             ColumnCreator.WinDim2.Visible = True
//         Else
//             'Hide window dimensions if no window option is checked
//             ColumnCreator.WinDim1.Visible = False
//             ColumnCreator.WinDim2.Visible = False
//         End If
// End Function

        public static int ValidatePlySeams(ColumnCreatorView ui, double[] ply_seams, double x, double y,
            double z)
        {
            var temp_ht = ply_seams.Sum();

            if (Math.Abs(temp_ht - z) > 0.001)
            {
                if (temp_ht > z)
                {
                    ui.TxtPlyError.Text = $"Remove {ConvertFtIn(temp_ht - z)}";
                }
                else if (temp_ht < z)
                {
                    ui.TxtPlyError.Text = $"Add {ConvertFtIn(z - temp_ht)}";
                }

                ui.TxtPlyError.Visibility = Visibility.Visible;
                return 0;
            }


            double ply_width_x;
            double ply_width_y;
            double min_ply_ht = 6;
            double max_ply_ht;
            ply_width_x = x + 1.5;
            ply_width_y = y + 1.5;
            if (ply_width_x > 48 || ply_width_y > 48) max_ply_ht = 48;
            else max_ply_ht = 96;
            foreach (var t in ply_seams)
            {
                if (t > max_ply_ht)
                {
                    ui.TxtPlyError.Text = $"{ConvertFtIn(t)} ply too tall. Max: {ConvertFtIn(max_ply_ht)}";
                    ui.TxtPlyError.Visibility = Visibility.Visible;
                    return 0;
                }

                if (t < min_ply_ht)
                {
                    ui.TxtPlyError.Text = $"{ConvertFtIn(t)} ply too small. Min: {ConvertFtIn(min_ply_ht)}";
                    ui.TxtPlyError.Visibility = Visibility.Visible;
                    return 0;
                }

                ui.TxtPlyError.Visibility = Visibility.Collapsed;
            }

            ui.TxtPlyError.Visibility = Visibility.Collapsed;
            return 1;
        }
    }
}