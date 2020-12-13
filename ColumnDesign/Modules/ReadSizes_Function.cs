using System;
using static ColumnDesign.Modules.ConvertFeetInchesToNumber;

namespace ColumnDesign.Modules
{
    public static class ReadSizesFunction
    {
        public static double[] ReadSizes(string csvPly)
        {
            var tempStr1 = "";
            string[] tempStr2 = {""};
            double[] tempOutput = {0};
            var j = 0;
            if (csvPly.Equals("")) return new double[] {0};
            for (var i = 0; i < csvPly.Length; i++)
            {
                if (csvPly.Substring(i, 1).Equals(","))
                {
                    Array.Resize(ref tempStr2, j + 1);
                    tempStr2[j] = tempStr1;
                    j += 1;
                    tempStr1 = "";
                }
                else if (i + 1 == csvPly.Length)
                {
                    Array.Resize(ref tempStr2, j + 1);
                    tempStr1 += csvPly.Substring(i, 1);
                    tempStr2[j] = tempStr1;
                    break;
                }
                else tempStr1 += csvPly.Substring(i, 1);
            }

            if (tempStr2.Length >= 11)
            {
                //TODO task dialog "Error: Too many ply sheets defined (max: 11)"
                return new double[] {0};
            }

            Array.Resize(ref tempOutput, tempStr2.Length);
            for (var i = 0; i < tempStr2.Length; i++) tempOutput[i] = ConvertToNum(tempStr2[i]);

            return tempOutput;
        }
    }
}