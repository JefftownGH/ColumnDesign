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

namespace ColumnDesign.Modules
{
    public static class UpdatePly_Function
    {
        public static void UpdatePly(ColumnCreatorView ui, ColumnCreatorViewModel vm)
        {
            var x = ConvertToNum(vm.WidthX);
            var y = ConvertToNum(vm.LengthY);
            var z = ConvertToNum(vm.HeightZ);
            if (x == 0 || y == 0 || z == 0) return;
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
            var pt_draw = new double[3];
            pt_draw[0] = 200;
            pt_draw[1] = 216;
            pt_draw[2] = 0;
            const double px_z_max = 200;
            var px_x = px_z_max * (x / z);
            var px_y = px_z_max * (y / z);
            double px_width;
            switch (vm.SlblAxis)
            {
                case "X":
                    px_width = px_x;
                    break;
                case "Y":
                    px_width = px_y;
                    break;
                default:
                    throw new ArgumentException("Error: Horizontal dimension not x or y");
            }
            
            if (vm.WindowX || vm.WindowY)
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

            ui.TreeLines.Children.Clear();
            ui.TreeLines.RowDefinitions.Clear();
            ui.TreeLines.RowDefinitions.Add(new RowDefinition());
            ui.TreeValues.Children.Clear();
            ui.TreeValues.RowDefinitions.Clear();   
            var lineStyle = ui.FindResource("TreeLine") as Style;
            var textStyle = ui.FindResource("TreeTextBlock") as Style;
            var gcd = MultiGcd(ply_seams);
            var lineCount = (int) z/ gcd;
            var winHeight = (int) ConvertToNum(vm.WinDim1);
            if (vm.WindowX || vm.WindowY)
            {
                gcd = Gcd(lineCount, winHeight );
                lineCount = (int) z/ gcd;
            }
            for (var i = 0; i < lineCount-1; i++)
            {
                ui.TreeLines.RowDefinitions.Add(new RowDefinition());
            }

            for (var i = 0; i < lineCount*2+1; i++)
            {
                ui.TreeValues.RowDefinitions.Add(new RowDefinition());
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
                Grid.SetRow(line, ui.TreeLines.RowDefinitions.Count-(int)sumPly/gcd);
                ui.TreeLines.Children.Add(line);
            }

            if (vm.WindowX || vm.WindowY)
            {
                var line = new Line
                {
                    X1 = 0,
                    X2 = 1,
                    VerticalAlignment = VerticalAlignment.Center,
                    Style = lineStyle,
                    Stroke = new SolidColorBrush(Colors.Blue)
                };
                Grid.SetRow(line, winHeight/gcd);
                ui.TreeLines.Children.Add(line);
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
                Grid.SetRow(value, ui.TreeValues.RowDefinitions.Count-1-(int)sumPly/gcd);
                ui.TreeValues.RowDefinitions[ui.TreeValues.RowDefinitions.Count - 1 - (int) sumPly / gcd].Height = new GridLength(20);
                ui.TreeValues.Children.Add(value);
                sumPly += t;
            }

        }

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

        public static int Gcd(int a, int b)
        {
            while (b != 0)
            {
                var temp = b;
                b = a % b;
                a = temp;
            }
            return a;
        }

        public static int MultiGcd(double[] n)
        {
            if (n.Length == 0) return 0;
            int i, gcd = (int)n[0];
            for (i = 0; i < n.Length - 1; i++)
                gcd = Gcd(gcd, (int)n[i + 1]);
            return gcd;
        }
    }
}