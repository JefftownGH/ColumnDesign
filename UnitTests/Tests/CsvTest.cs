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
            Assert.AreEqual(ReadSizes("1,2,3,4,5,6,7,8,9,10,11,12"), new double[] {0});
        }
    }
}