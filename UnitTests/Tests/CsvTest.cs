using System;
using ColumnDesign;
using ColumnDesign.Modules;
using NUnit.Framework;
using static ColumnDesign.Modules.ReadSizesFunction;

namespace UnitTests.Tests
{
    public class CsvTest
    {
        
        [Test]
        public void ReadSizesFunction()
        {
            Assert.AreEqual(ReadSizes(""), new double[] {0});
            Assert.AreEqual(ReadSizes("10,12"), new double[] {10, 12});
            Assert.Throws<ArgumentException>(()=> ReadSizes("1,2,3,4,5,6,7,8,9,10,11,12"));
        }
        [Test]
        public void ReadFile()
        {
            Assert.That(()=>ImportMatrixFunction.ImportMatrix(@$"{GlobalNames.WtFileLocationPrefix}Columns\clamp_matrix.csv"), Throws.Nothing);
        }
    }
}