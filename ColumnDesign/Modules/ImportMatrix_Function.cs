using System;
using System.IO;
using System.Linq;

namespace ColumnDesign.Modules
{
    public static class ImportMatrixFunction
    {
        public static int[,] ImportMatrix(string filepath)
        {
            var i = 0;
            var nRowsTot = 0;
            var myString = "";
            var returnMatrix = new int[,] { };
            using (var reader = File.OpenText(filepath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    myString = line;
                    nRowsTot += 1;
                }
            }

            var tempRow = myString.Split(',');
            returnMatrix = ResizeArray<int>(returnMatrix, nRowsTot, tempRow.Length);
            using (var reader = File.OpenText(filepath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    myString = line;
                    tempRow = myString.Split(',');
                    for (var j = 0; j < tempRow.Length; j++)
                    {
                        returnMatrix[i, j] = int.Parse(tempRow[j]);
                    }

                    i++;
                }
            }
            return returnMatrix;
        }

        public static T[,] ResizeArray<T>(T[,] original, int rows, int cols)
        {
            var newArray = new T[rows, cols];
            var minRows = Math.Min(rows, original.GetLength(0));
            var minCols = Math.Min(cols, original.GetLength(1));
            for (var i = 0; i < minRows; i++)
            {
                for (var j = 0; j < minCols; j++)
                    newArray[i, j] = original[i, j];
            }

            return newArray;
        }
    }
}