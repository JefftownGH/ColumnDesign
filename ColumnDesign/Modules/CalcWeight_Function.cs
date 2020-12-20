using System;

namespace ColumnDesign.Modules
{
    public static class CalcWeightFunction
    {
        private const double WtPly = 2.2;
        private const double Wt2X4 = 1.31;
        private const double WtLvl = 1.8;
        private const double Wt0824Clamp = 76;
        private const double Wt1236Clamp = 100;
        private const double Wt2448Clamp = 123;
        private const double Wt36ScissorClamp = 40;
        private const double Wt48ScissorClamp = 56;
        private const double Wt60ScissorClamp = 85;
        private const double Wt0711Brace = 52;
        private const double Wt1119Brace = 76;
        private const double WtExtra = 50;

        public static int wt_total(double x, double y, double z, double nStudsX, double nStudsY,
            double studType, double nClamps, double clampSize, double braceSize)
        {
            var wtStud = studType switch
            {
                1 => Wt2X4,
                2 => WtLvl,
                _ => throw new ArgumentException("Invalid stud type detected in weight calculation")
            };
            var wtClamp = clampSize switch
            {
                1 => Wt0824Clamp,
                2 => Wt1236Clamp,
                3 => Wt2448Clamp,
                4 => Wt36ScissorClamp,
                5 => Wt48ScissorClamp,
                6 => Wt60ScissorClamp,
                _ => throw new ArgumentException("Invalid clamp size detected in weight calculation")
            };
            var wtBrace = braceSize switch
            {
                0 => 0,
                1 => Wt0711Brace,
                2 => Wt1119Brace,
                _ => throw new ArgumentException("Invalid brace size detected in weight calculation")
            };
            var wtTotal = WtPly * ((x + 2.25) * z * 2 / 144 + (y + 2.25) * z * 2 / 144) +
                           wtStud * (z / 12) * 2 * (nStudsX + nStudsY) +
                           wtClamp * nClamps + wtBrace * 3;
            wtTotal = Math.Ceiling((wtTotal + WtExtra) / 100);
            return (int)wtTotal * 100;
        }

        public static double wt_panel(double x, double y, double z, double nStudsX, double nStudsY,
            double studType)
        {
            var wtStud = studType switch
            {
                1 => Wt2X4,
                2 => WtLvl,
                _ => throw new ArgumentException("Invalid stud type detected in weight calculation")
            };
            var largerSide = x > y ? x : y;
            var mostStuds = nStudsX > nStudsY ? nStudsX : nStudsY;
            var wtPanel = WtPly * ((largerSide + 4.5) * z / 144) + wtStud * (z / 12) * mostStuds;

            wtPanel = Math.Ceiling(wtPanel * 1.1 / 10);
            return wtPanel * 10;
        }
    }
}