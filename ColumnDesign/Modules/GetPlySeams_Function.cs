using System;

namespace ColumnDesign.Modules
{
    public class GetPlySeamsFunction
    {
        public static double[] GetPlySeams(double x, double y, double z, int tabControlIndex)
        {
            double plyWidthX;
            double plyWidthY;
            const double plyTopHtMin = 6;
            double[] plySeamsFun = { };
            double maxPlyHt;
            if (tabControlIndex == 0)
            {
                plyWidthX = x + 1.5;
                plyWidthY = y + 1.5;
            }
            else
            {
                plyWidthX = x;
                plyWidthY = y + 4.5;
            }

            if (plyWidthX > 48 || plyWidthY > 48) maxPlyHt = 48;
            else maxPlyHt = 96;
            var plyTopHt = z - maxPlyHt * (int) (z / maxPlyHt);
            var plyBotN = (int) (z / maxPlyHt);
            Array.Resize(ref plySeamsFun, plyBotN + (int) Math.Abs(plyTopHt) != 0 ? 1 : 0);
            if (z <= maxPlyHt)
            {
                plySeamsFun[0] = z;
                return plySeamsFun;
            }

            for (var i = 0; i < plyBotN; i++) plySeamsFun[i] = maxPlyHt;

            if (plyTopHt != 0) plySeamsFun[plySeamsFun.Length - 1] = plyTopHt;

            if (plyTopHt < plyTopHtMin && plyTopHt > 0)
            {
                plyTopHt += 48 * maxPlyHt / 96;
                plyBotN = (int) ((z - 48 * maxPlyHt / 96 - plyTopHt) / maxPlyHt);
                Array.Resize(ref plySeamsFun, plyBotN + 2);
                for (var i = 0; i < plyBotN; i++) plySeamsFun[i] = maxPlyHt;
                plySeamsFun[plySeamsFun.Length - 2] = 48 * maxPlyHt / 96;
                plySeamsFun[plySeamsFun.Length - 1] = plyTopHt;
            }

            return plySeamsFun;
        }
    }
}