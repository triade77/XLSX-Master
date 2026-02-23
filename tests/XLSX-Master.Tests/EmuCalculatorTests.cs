using Microsoft.VisualStudio.TestTools.UnitTesting;
using XlsxMaster.Helpers;

namespace XlsxMaster.Tests
{
    [TestClass]
    public class EmuCalculatorTests
    {
        [TestMethod]
        public void ParseAnchor_StandardInput_ReturnsCorrectIndices()
        {
            var (colFrom, rowFrom, colTo, rowTo) = EmuCalculator.ParseAnchor("E1:M20");

            Assert.AreEqual(4, colFrom);   // E = 5th column → 0-based = 4
            Assert.AreEqual(0, rowFrom);   // 행 1 → 0-based = 0
            Assert.AreEqual(12, colTo);    // M = 13th column → 0-based = 12
            Assert.AreEqual(19, rowTo);    // 행 20 → 0-based = 19
        }

        [TestMethod]
        public void ParseAnchor_LowercaseInput_ReturnsCorrectIndices()
        {
            var (colFrom, rowFrom, colTo, rowTo) = EmuCalculator.ParseAnchor("a1:b10");

            Assert.AreEqual(0, colFrom);
            Assert.AreEqual(0, rowFrom);
            Assert.AreEqual(1, colTo);
            Assert.AreEqual(9, rowTo);
        }

        [TestMethod]
        [ExpectedException(typeof(System.ArgumentException))]
        public void ParseAnchor_InvalidFormat_ThrowsArgumentException()
        {
            EmuCalculator.ParseAnchor("INVALID");
        }

        [TestMethod]
        [ExpectedException(typeof(System.ArgumentException))]
        public void ParseAnchor_EmptyString_ThrowsArgumentException()
        {
            EmuCalculator.ParseAnchor("");
        }

        [TestMethod]
        public void ColumnLetterToIndex_SingleLetter()
        {
            Assert.AreEqual(0, EmuCalculator.ColumnLetterToIndex("A"));
            Assert.AreEqual(1, EmuCalculator.ColumnLetterToIndex("B"));
            Assert.AreEqual(25, EmuCalculator.ColumnLetterToIndex("Z"));
        }

        [TestMethod]
        public void ColumnLetterToIndex_DoubleLetters()
        {
            Assert.AreEqual(26, EmuCalculator.ColumnLetterToIndex("AA"));
            Assert.AreEqual(27, EmuCalculator.ColumnLetterToIndex("AB"));
        }

        [TestMethod]
        public void ColumnIndexToEmu_ZeroIndex_ReturnsZero()
        {
            Assert.AreEqual(0L, EmuCalculator.ColumnIndexToEmu(0));
        }

        [TestMethod]
        public void ColumnIndexToEmu_OneIndex_ReturnsDefaultWidth()
        {
            Assert.AreEqual(EmuCalculator.DefaultColumnWidthEmu, EmuCalculator.ColumnIndexToEmu(1));
        }
    }
}
