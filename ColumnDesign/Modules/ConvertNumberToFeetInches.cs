using System;
using System.Globalization;

namespace ColumnDesign.Modules
{
    public static class ConvertNumberToFeetInches
    {
        public static string ConvertFtIn(double num)
        {
            var xFra = "";
            string strNeg;
            if (num < 0)
            {
                strNeg = "-";
                num *= -1;
            }
            else strNeg = "";

            num = Math.Round(num * 8 - 0.000000001, 0) / 8;
            var xFt = ((int) Math.Round(num / 12, 3)).ToString(CultureInfo.InvariantCulture);
            var xIn = (num - 12 * (int) Math.Round(num / 12, 3)).ToString(CultureInfo.InvariantCulture);
            if (Math.Round(num - (int) Math.Round(num, 3), 3) != 0)
            {
                xFra = Dec2Fraction(12 * (num / 12 - (int) (num / 12)) - (int) (num - 12 * (int) (num / 12)));
                xIn = ((int) (num - 12 * (int) (num / 12))).ToString();
            }

            if (xFt.Equals("0"))
            {
                return Math.Round(num - (int) Math.Round(num, 3), 3) != 0
                    ? $"{strNeg}{xIn} {xFra}\""
                    : $"{strNeg}{xIn}\"";
            }

            return Math.Round(num - (int) Math.Round(num, 3), 3) != 0
                ? $"{strNeg}{xFt}'-{xIn} {xFra}\""
                : $"{strNeg}{xFt}'-{xIn}\"";
        }

        private static string Dec2Fraction(double dFraction)
        {
            double lUpperPart = 1;
            double lLowerPart = 1;
            dFraction = Math.Round(dFraction, 3);
            var df = lUpperPart / lLowerPart;
            while (Math.Abs(df - dFraction) > 0.003)
            {
                if (df < dFraction) lUpperPart += 1;
                else
                {
                    lLowerPart += 1;
                    lUpperPart = dFraction * lLowerPart;
                }

                df = lUpperPart / lLowerPart;
            }

            return $"{lUpperPart}/{lLowerPart}";
        }
    }
}