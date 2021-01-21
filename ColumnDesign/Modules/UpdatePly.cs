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
        public static void UpdatePly(ColumnCreatorView ui, ColumnCreatorViewModel vm, bool editWin = false)
        {
            var x = ConvertToNum(vm.WidthX);
            var y = ConvertToNum(vm.LengthY);
            var z = ConvertToNum(vm.HeightZ);
            if (x == 0 || y == 0 || z == 0) return;
            double[] ply_seams = { };
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
            
            ui.TreeLines.Children.Clear();
            ui.TreeLines.RowDefinitions.Clear();
            ui.TreeLines.RowDefinitions.Add(new RowDefinition());
            ui.TreeValues.Children.Clear();
            ui.TreeValues.RowDefinitions.Clear();   
            ui.GridWinDim.RowDefinitions.Clear();
            var lineStyle = ui.FindResource("TreeLine") as Style;
            var textStyle = ui.FindResource("TreeTextBlock") as Style;
            var gcd = MultiGcd(ply_seams);
            var lineCount = (int)Math.Round(z/ gcd);
            var winHeight1 = (int) ConvertToNum(vm.WinDim1);
            var winHeight2 = (int) ConvertToNum(vm.WinDim2);
            if (winHeight1>z || winHeight2>z) return;
            if (vm.WindowX || vm.WindowY)
            {
                if (winHeight1 != 0 && winHeight2 != 0)
                {
                    gcd = Gcd(lineCount, winHeight1);
                    lineCount = (int) Math.Round(z / gcd);
                }

                for (var i = 0; i < lineCount*2+1; i++)
                {
                    ui.GridWinDim.RowDefinitions.Add(new RowDefinition());
                }
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
                    StrokeThickness = 2,
                    Style = lineStyle,
                };
                Grid.SetRow(line, ui.TreeLines.RowDefinitions.Count-(int)Math.Round(sumPly/gcd));
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
                    StrokeThickness = 2,
                    Stroke = new SolidColorBrush(Colors.Red)
                };
                int windowRow;
                if ( winHeight1==0)
                {
                    windowRow = 0;
                    line.VerticalAlignment = VerticalAlignment.Top;
                }
                else if ( winHeight2==0)
                {
                    windowRow  =ui.TreeLines.RowDefinitions.Count-1;
                    line.VerticalAlignment = VerticalAlignment.Bottom;
                }
                else
                {
                   windowRow  = (int)Math.Round((double)winHeight1 / gcd,0);
                }
                Grid.SetRow(line, windowRow);
                ui.TreeLines.Children.Add(line);

                var winDim1Row = (int)Math.Round((double)windowRow / 2,0);
                var winDim2Row = (int)Math.Round((double)(ui.GridWinDim.RowDefinitions.Count+windowRow)/2,0);
                Grid.SetRow(ui.WinDim1, winDim1Row);
                Grid.SetRow(ui.WinDim2, winDim2Row);
                ui.GridWinDim.RowDefinitions[winDim1Row].Height = new GridLength(20);
                ui.GridWinDim.RowDefinitions[winDim2Row].Height = new GridLength(20);
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
                var textRow = ui.TreeValues.RowDefinitions.Count - 1 - (int)Math.Round(sumPly / gcd, 0);
                Grid.SetRow(value, textRow);
                ui.TreeValues.RowDefinitions[textRow].Height = new GridLength(20);
                ui.TreeValues.Children.Add(value);
                sumPly += t;
            }

            if (editWin) return;
            vm.WinDim1 = ConvertFtIn(z / 5);
            vm.WinDim2 = ConvertFtIn(z - ConvertToNum(vm.WinDim1));
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