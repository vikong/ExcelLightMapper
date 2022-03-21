using ExcelLightMapper;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests
{
	[TestClass]
	public class MapperTests
	{
		[DataTestMethod]
		[DataRow("A", 1)]
		[DataRow("B", 2)]
		[DataRow("c", 3)]
		[DataRow("AA", 27)]
		[DataRow("AZ", 52)]
		[DataRow("BA", 53)]
		public void GetColumnNumber_Returns_Expected(string columnName, int columnIndex)
		{
			Assert.AreEqual(columnIndex, ExcelReader.GetColumnNumber(columnName));
		}

		[TestMethod]
		public void Mapper_FromRow_SetsRowRange()
		{
			var mapper = new RowMapper<Primitive>();
			mapper.FromRows(3, 6);

			Assert.IsFalse(mapper.IsRequestedRow(0));
			Assert.IsFalse(mapper.IsRequestedRow(1));

			Assert.IsTrue(mapper.IsRequestedRow(2));
			Assert.IsTrue(mapper.IsRequestedRow(3));
			Assert.IsTrue(mapper.IsRequestedRow(5));

			Assert.IsFalse(mapper.IsRequestedRow(6));
			Assert.IsFalse(mapper.IsRequestedRow(10));
		}

		[TestMethod]
		public void Mapper_WithoutRow_AcceptAllRow()
		{
			var mapper = new RowMapper<Primitive>();

			Assert.IsTrue(mapper.IsRequestedRow(0));
			Assert.IsTrue(mapper.IsRequestedRow(2));
			Assert.IsTrue(mapper.IsRequestedRow(300));
		}

		[TestMethod]
		public void Mapper_WithFirstRow_SetsRowRange()
		{
			var mapper = new RowMapper<Primitive>();
			mapper.FromRows(3);

			Assert.IsFalse(mapper.IsRequestedRow(0));
			Assert.IsFalse(mapper.IsRequestedRow(1));

			Assert.IsTrue(mapper.IsRequestedRow(2));
			Assert.IsTrue(mapper.IsRequestedRow(300));
		}

		[TestMethod]
		public void Mapper_CustomRange_SetsRowRange()
		{
			string sheet = "Primitives";
			var mapper = new RowMapper<Primitive>();
			mapper.Range(c => c.Sheet == sheet && c.Column == 1);

			Assert.IsTrue(mapper.IsRequestedRange(new ExCell(1, 0) { Sheet = sheet }));
			Assert.IsTrue(mapper.IsRequestedRange(new ExCell(100, 0) { Sheet = sheet }));

			Assert.IsFalse(mapper.IsRequestedRange(new ExCell(1, 0) { Sheet = "Third sheet" }));
			Assert.IsFalse(mapper.IsRequestedRange(new ExCell(1, 30) { Sheet = sheet }));
		}

		[TestMethod]
		public void Mapper_MapColumnNumber_AddsColumnMap()
		{
			var mapper = new RowMapper<Primitive>();
			int column = 1;
			int columnIndex = column - 1;
			mapper.MapColumn(column, f => f.IntValue);
			Assert.IsTrue(mapper.TryGetCellMap(columnIndex, out ICellMap<Primitive> cellMap));
			Assert.AreEqual(columnIndex, cellMap.ColumnIndex);
		}

		[TestMethod]
		public void Mapper_MapColumnName_AddsColumnMap()
		{
			var mapper = new RowMapper<Primitive>();
			string columnName = "A";
			int columnIndex = 0;
			mapper.MapColumn(columnName, f => f.IntValue);
			Assert.IsTrue(mapper.TryGetCellMap(columnIndex, out ICellMap<Primitive> cellMap));
			Assert.AreEqual(columnIndex, cellMap.ColumnIndex);
		}
	}
}