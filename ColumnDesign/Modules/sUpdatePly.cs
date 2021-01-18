using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using ColumnDesign.UI;
using ColumnDesign.ViewModel;
using static ColumnDesign.Modules.ConvertFeetInchesToNumber;
using static ColumnDesign.Modules.ConvertNumberToFeetInches;
using static ColumnDesign.Modules.GetPlySeamsFunction;
using static ColumnDesign.Modules.ReadSizesFunction;
using static ColumnDesign.Modules.UpdatePly_Function;

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

            if (!ui.SBoxPlySeams.Text.Equals(""))
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
            ui.STreeLines.Children.Clear();
            ui.STreeLines.RowDefinitions.Clear();
            ui.STreeLines.RowDefinitions.Add(new RowDefinition());
            ui.STreeValues.Children.Clear();
            ui.STreeValues.RowDefinitions.Clear();   
            var lineStyle = ui.FindResource("TreeLine") as Style;
            var textStyle = ui.FindResource("TreeTextBlock") as Style;
            var gcd = MultiGcd(ply_seams);
            var lineCount = (int) z/ gcd;
            for (var i = 0; i < lineCount-1; i++)
            {
                ui.STreeLines.RowDefinitions.Add(new RowDefinition());
            }

            for (var i = 0; i < lineCount*2+1; i++)
            {
                ui.STreeValues.RowDefinitions.Add(new RowDefinition());
            }

            var sumPly = 0d;
            for (var i = 0; i < ply_seams.Length-1; i++)
            {
                sumPly += ply_seams[i];
                var line = new Line
                {
                    X1 = 0,
                    X2 = 1,
                    VerticalAlignment = VerticalAlignment.Center,
                    Style = lineStyle,
                };
                Grid.SetRow(line, ui.STreeLines.RowDefinitions.Count-(int)sumPly/gcd);
                ui.STreeLines.Children.Add(line);
            }
            sumPly = 0d;
            foreach (var t in ply_seams)
            {
                sumPly += t;

                var value = new TextBlock
                {
                    Text = ConvertFtIn(t),
                    Style = textStyle,
                };
                Grid.SetRow(value, ui.STreeValues.RowDefinitions.Count-1-(int)sumPly/gcd);
                ui.STreeValues.RowDefinitions[ui.STreeValues.RowDefinitions.Count - 1 - (int) sumPly / gcd].Height = new GridLength(20);
                ui.STreeValues.Children.Add(value);
                sumPly += t;
            }
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