using System;
using System.Linq;
using FluentAssertions;
using Xunit;

namespace SpreadsheetCellRef.Tests
{
    public class CellRefTests
    {
        [Fact]
        public void CreatesCorrectColumn()
        {
            var cellRef = new CellRef("AB23");
            cellRef.Column.Should().Be("AB");
        }
        
        [Fact]
        public void CreatesCorrectColumnNumber()
        {
            var cellRef = new CellRef("AB23");
            cellRef.ColumnNumber.Should().Be(28);
        }
        
        [Fact]
        public void CreatesCorrectRow()
        {
            var cellRef = new CellRef("AB23");
            cellRef.Row.Should().Be(23);
        }

        
        [Fact]
        public void SendingInvalidStringToConstructorThrowsMeaningfulException()
        {
            Action action = () =>
            {
                // ReSharper disable once UnusedVariable
                var cellRef = new CellRef("Not a valid cell ref");
            };
            action.Should().Throw<FormatException>().WithMessage("Input string was not in a correct format. Was expecting cell address in ColumnLetterRowNumber format (e.g. 'AM56') but received 'Not a valid cell ref'");
        }

        [Fact]
        public void IntegerConstructorWorks()
        {
            var cellRef = new CellRef(1,1);
            cellRef.ToString().Should().Be("A1");
        }
        
        [Fact]
        public void RowNumberWorks()
        {
            var cellRef = new CellRef(86,86);
            cellRef.ColumnNumber.Should().Be(86);
        }

        [Fact]
        public void AddColumnsWorks()
        {
            var a1 = new CellRef("A1");
            var g1 = a1.AddColumns(6);
            g1.ToString().Should().Be("G1");
        }

        [Fact]
        public void AddLargeColumnsWorks()
        {
            var aa1 = new CellRef("AA1");
            var ag1 = aa1.AddColumns(6);
            ag1.ToString().Should().Be("AG1");
        }


        [Fact]
        public void AddRowsWorks()
        {
            var a1 = new CellRef("A1");
            var a7 = a1.AddRows(6);
            a7.ToString().Should().Be("A7");
        }

        [Fact]
        public void ColumnNumberWorks()
        {
            var aa1 = new CellRef("AA1");
            aa1.ColumnNumber.Should().Be(27);
        }

        [Fact]
        public void ColumnNameToNumberWorks()
        {
            CellRef.ColumnNameToNumber("A").Should().Be(1);
            CellRef.ColumnNameToNumber("Z").Should().Be(26);
            CellRef.ColumnNameToNumber("AA").Should().Be(27);
        }

        [Fact]
        public void NumberToColumnNameWorks()
        {
            CellRef.NumberToColumnName(27).Should().Be("AA");
        }

        [Fact]
        public void RangeWorks()
        {
            var range = CellRef.Range("A1", "B2").ToArray();
            range[0].ToString().Should().Be("A1");
            range[1].ToString().Should().Be("B1");
            range[2].ToString().Should().Be("A2");
            range[3].ToString().Should().Be("B2");
        }

        [Fact]
        public void SingleColumnRangeWorks()
        {
            var range = CellRef.Range("A1", "A4").ToArray();
            range[0].ToString().Should().Be("A1");
            range[1].ToString().Should().Be("A2");
            range[2].ToString().Should().Be("A3");
            range[3].ToString().Should().Be("A4");
        }

        [Fact]
        public void ToStringWorks()
        {
            var aa1 = new CellRef("AA1");
            aa1.ToString().Should().Be("AA1");
        }

        [Fact]
        public void EqualityWorks()
        {
            var cellRefA1 = new CellRef("A1");
            var cellRefA1FromInts = new CellRef(1,1);
            (cellRefA1 == cellRefA1FromInts).Should().BeTrue();
            (cellRefA1 != cellRefA1FromInts).Should().BeFalse();
        }
        
        [Fact]
        public void InequalityWorks()
        {
            var cellRefA1 = new CellRef("A1");
            var cellRefA2FromInts = new CellRef(1,2);
            (cellRefA1 != cellRefA2FromInts).Should().BeTrue();
            (cellRefA1 == cellRefA2FromInts).Should().BeFalse();
            cellRefA1.Equals(cellRefA2FromInts).Should().BeFalse();
        }
    }
}