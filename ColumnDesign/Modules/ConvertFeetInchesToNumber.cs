using System.Globalization;

namespace ColumnDesign.Modules
{
    public static class ConvertFeetInchesToNumber
    {
        public static double ConvertToNum(string stringInput)
        {
            var negTog = 1;
            double outNum = 0;
            var strTemp1 = "";
            var strTemp2 = "";
            var strTemp3 = "";
            var strTemp4 = "";
            if (stringInput == null || stringInput.Equals("") || string.Equals(stringInput, " ") ||
                string.Equals(stringInput, ".")) return 0;

            for (var iCnt0 = 0; iCnt0 < stringInput.Length; iCnt0++)
            {
                if (int.TryParse(stringInput.Substring(iCnt0, 1), out _))
                {
                    if (stringInput.Substring(0, 1).Equals("-"))
                    {
                        negTog = -1;
                        stringInput = stringInput.Substring(1, stringInput.Length);
                    }

                    var foundNum = 0;
                    int iCnt1;
                    for (iCnt1 = 0; iCnt1 < stringInput.Length; iCnt1++)
                    {
                        if (int.TryParse(stringInput.Substring(iCnt1, 1), out _) ||
                            stringInput.Substring(iCnt1, 1).Equals("."))
                        {
                            strTemp1 += stringInput.Substring(iCnt1, 1);
                            foundNum = 1;
                        }
                        else if (foundNum == 1) break;
                    }

                    foundNum = 0;
                    int iCnt2;
                    for (iCnt2 = iCnt1 + 1; iCnt2 < stringInput.Length; iCnt2++)
                    {
                        if (int.TryParse(stringInput.Substring(iCnt2, 1), out _) ||
                            stringInput.Substring(iCnt2, 1).Equals("."))
                        {
                            strTemp2 += stringInput.Substring(iCnt2, 1);
                            foundNum = 1;
                        }
                        else if (foundNum == 1) break;
                    }

                    foundNum = 0;
                    int iCnt3;
                    for (iCnt3 = iCnt2 + 1; iCnt3 < stringInput.Length; iCnt3++)
                    {
                        if (int.TryParse(stringInput.Substring(iCnt3, 1), out _) ||
                            stringInput.Substring(iCnt3, 1).Equals("."))
                        {
                            strTemp3 += stringInput.Substring(iCnt3, 1);
                            foundNum = 1;
                        }
                        else if (foundNum == 1) break;
                    }

                    int iCnt4;
                    for (iCnt4 = iCnt3 + 1; iCnt4 < stringInput.Length; iCnt4++)
                    {
                        if (int.TryParse(stringInput.Substring(iCnt4, 1), out _) ||
                            stringInput.Substring(iCnt4, 1).Equals("."))
                        {
                            strTemp4 += stringInput.Substring(iCnt4, 1);
                            foundNum = 1;
                        }
                        else if (foundNum == 1) break;
                    }

                    if (!strTemp1.Equals("") && !strTemp1.Equals(".") && !strTemp2.Equals("") &&
                        !strTemp2.Equals(".") && !strTemp3.Equals("") && !strTemp3.Equals(".") &&
                        !strTemp4.Equals("") && !strTemp4.Equals("."))
                    {
                        if (double.Parse(strTemp4, CultureInfo.InvariantCulture) == 0) outNum = 0;
                        else
                        {
                            outNum = 12 * double.Parse(strTemp1, CultureInfo.InvariantCulture) +
                                     double.Parse(strTemp2, CultureInfo.InvariantCulture) +
                                     double.Parse(strTemp3, CultureInfo.InvariantCulture) /
                                     double.Parse(strTemp4, CultureInfo.InvariantCulture);
                        }
                    }
                    else if (!strTemp1.Equals("") && !strTemp1.Equals(".") && !strTemp2.Equals("") &&
                             !strTemp2.Equals(".") && !strTemp3.Equals("") && !strTemp3.Equals("."))
                    {
                        if (double.Parse(strTemp3, CultureInfo.InvariantCulture) == 0) outNum = 0;
                        else
                        {
                            outNum = 12 * double.Parse(strTemp1, CultureInfo.InvariantCulture) +
                                     double.Parse(strTemp2, CultureInfo.InvariantCulture) /
                                     double.Parse(strTemp3, CultureInfo.InvariantCulture);
                        }
                    }
                    else if (!strTemp1.Equals("") && !strTemp1.Equals(".") && !strTemp2.Equals("") &&
                             !strTemp2.Equals("."))
                    {
                        outNum = 12 * double.Parse(strTemp1, CultureInfo.InvariantCulture) +
                                 double.Parse(strTemp2, CultureInfo.InvariantCulture);
                    }
                    else if (!strTemp1.Equals("") && !strTemp1.Equals("."))
                    {
                        if (iCnt1 < stringInput.Length && stringInput.Substring(iCnt1, 1).Equals("'"))
                        {
                            outNum = 12 * double.Parse(strTemp1, CultureInfo.InvariantCulture);
                        }

                        else outNum = double.Parse(strTemp1, CultureInfo.InvariantCulture);
                    }
                    else return 0;

                    return negTog * outNum;
                }

                if (iCnt0 == stringInput.Length) return 0;
            }

            return outNum;
        }
    }
}