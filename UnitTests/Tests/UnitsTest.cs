using NUnit.Framework;
using static ColumnDesign.Modules.ConvertFeetInchesToNumber;
using static ColumnDesign.Modules.ConvertNumberToFeetInches;

namespace UnitTests.Tests
{
    public class UnitsTest
    {
        [Test]
        public void ConvertStringInchesToNum()
        {
            Assert.AreEqual(ConvertToNum("125"), 125, 0.1);
            Assert.AreEqual(ConvertToNum("125.5"), 125.5, 0.1);
            Assert.AreEqual(ConvertToNum("10'"), 120, 0.1);
            Assert.AreEqual(ConvertToNum("10'-5.5\""), 125.5, 0.1);
            Assert.AreEqual(ConvertToNum("10'5.5"), 125.5, 0.1);
            Assert.AreEqual(ConvertToNum("10 5.5"), 125.5, 0.1);
            Assert.AreEqual(ConvertToNum("10'.5"), 120.5, 0.1);
            Assert.AreEqual(ConvertToNum("10'-5 1/2\""), 125.5, 0.1);
            Assert.AreEqual(ConvertToNum("10'5 1/2\""), 125.5, 0.1);
            Assert.AreEqual(ConvertToNum("10' 1/2\""), 120.5, 0.1);
            Assert.AreEqual(ConvertToNum("10' 5.5/2\""), 122.75, 0.1);
            Assert.AreEqual(ConvertToNum("10' 5.5 5/2\""), 128, 0.1);
            Assert.AreEqual(ConvertToNum("10' 5.5 5.5/2\""), 128.25, 0.1);
        }

        [Test]
        public void ConvertNumToStringInches()
        {
            Assert.AreEqual(ConvertFtIn(125), "10'-5\"");
            Assert.AreEqual(ConvertFtIn(125.5), "10'-5 1/2\"");
            Assert.AreEqual(ConvertFtIn(122.75), "10'-2 1.5/2\"");
        }
    }
}