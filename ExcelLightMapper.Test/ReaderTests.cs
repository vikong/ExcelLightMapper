using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using ExcelLightMapper;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTests
{

	[TestClass]
	public class ReaderTests
	{
		private static string DefaultTestXlsx
			=> Helper.ResourcePath("Primitives.xlsx");

		[TestMethod]
		public void GetRows_Returns_DateValues()
		{
			var mapper = new RowMapper<Primitive>();
			mapper.MapColumn("D", f => f.DateValue);
			mapper.FromRows(3, 6);
			var actual = ExcelReader
				.GetRows(DefaultTestXlsx, "Primitives", mapper)
				.Select(e => e.DateValue)
				.ToArray();
			var expected = new DateTime[]
			{
				new DateTime(2020, 12, 4),
				new DateTime(2017, 12, 31),
				default,
				default
			};
			CollectionAssert.AreEqual(expected, actual);
		}

		[TestMethod]
		public void GetRows_Returns_IntValues()
		{
			var mapper = new RowMapper<Primitive>();
			mapper.MapColumn(1, f => f.IntValue);
			mapper.FromRows(3, 6);
			var actual = ExcelReader
				.GetRows(DefaultTestXlsx, "Primitives", mapper)
				.Select(e => e.IntValue)
				.ToArray();
			var expected = new int?[]
			{
				1,
				2,
				null,
				null
			};
			CollectionAssert.AreEqual(expected, actual);
		}

		[TestMethod]
		public void GetRows_Returns_FuncValues()
		{
			var mapper = new RowMapper<Primitive>()
				.FromRows(3, 6);
			mapper.MapColumn("F", f => f.FuncColumn);
			var actual = ExcelReader
				.GetRows(DefaultTestXlsx, "Primitives", mapper)
				.Select(e => e.FuncColumn)
				.ToArray();
			var expected = new Decimal?[]
			{
				27.13M,
				12.12M,
				null,
				39.25M
			};
			CollectionAssert.AreEqual(expected, actual);
		}

		[TestMethod]
		public void ReadRow_FromSheet_ReturnsExpected()
		{
			var mapper = new RowMapper<Primitive>();
			mapper.MapColumn(2, f => f.StringValue);
			var actual = ExcelReader.GetRows(DefaultTestXlsx, "Third Sheet", mapper.FromRows(3, 6))
				.Select(e => e.StringValue)
				.ToArray();
			Assert.AreEqual(4, actual.Length);
			var expected = new String[]
				{
					"s3-1", "s3-2", "s3-3", "s3-4"
				};
			CollectionAssert.AreEqual(expected, actual);
		}

		[TestMethod]
		public void ReadRow_FromAllSheets_ReturnsExpected()
		{
			var mapper = new RowMapper<Primitive>();
			mapper.FromRows(3, 4);
			mapper.MapColumn(2, f => f.StringValue);
			var actual = ExcelReader.GetRows(DefaultTestXlsx, mapper)
				.Select(e => e.StringValue)
				.ToArray();
			Assert.AreEqual(4, actual.Length);
			var expected = new String[]
				{
					"a", "b", "s3-1", "s3-2"
				};
			CollectionAssert.AreEqual(expected, actual);
		}

		[TestMethod]
		public void ReadRows_ActionConvertion_ReturnsExpected()
		{
			var mapper = new RowMapper<Primitive>();
			mapper.MapColumn(8, (o, v) => {
				string strVal = v.ToString();
				if (strVal.Equals("Success", StringComparison.InvariantCultureIgnoreCase)|| strVal == "1")
				{ o.BoolValue = true; }
				else if (strVal.Equals("Fault", StringComparison.InvariantCultureIgnoreCase) || strVal == "0")
				{ o.BoolValue = false; }
				else
				{ 
					o.BoolValue = null;
					o.StringValue = "undefined";
				}
			});

			var actual = ExcelReader.GetRows(DefaultTestXlsx, "Primitives", mapper.FromRows(3, 6))
				.ToArray();
				
			var expected = new bool?[] { true, false, null, true };

			CollectionAssert.AreEqual(expected, actual.Select(p => p.BoolValue).ToArray());
			Assert.AreEqual(1, actual.Count(p=>p.StringValue== "undefined"));
		}
		[TestMethod]
		public void ReadRows_CustomConvertion_ReturnsExpected()
		{
			var mapper = new RowMapper<Primitive>();
			mapper.MapColumn(1, f => f.StringValue)
				.From(x => x != null ? x.ToString() : String.Empty);

			var actual = ExcelReader.GetRows(DefaultTestXlsx, "Primitives", mapper.FromRows(3, 6))
				.Select(p => p.StringValue)
				.ToArray();

			var expected = new string[] { "1", "2", "totally_invalid", String.Empty };

			CollectionAssert.AreEqual(expected, actual);
		}

		[TestMethod]
		public void GetRows_WithHandler_TreatsError()
		{
			var mapper = new RowMapper<Primitive>();
			mapper.MapColumn(1, f => f.IntValue).IgnoreErrors();
			List<ExceptionInfo> actual = new List<ExceptionInfo>();
			mapper.OnError((e) => actual.Add(e));

			var rows = ExcelReader
				.GetRows(DefaultTestXlsx, "Primitives", mapper.FromRows(3, 6))
				.Select(e => e.IntValue)
				.ToArray();

			var expected = new ExCell[] {
				new ExCell(4,0){Sheet="Primitives"}
			};
			
			Assert.AreEqual(1, actual.Count);

			CollectionAssert.AreEqual(expected, actual.Select(x=>x.Cell).ToArray());
		}
		[TestMethod]
		public void GetRows_WithHandler_ThrowsError()
		{
			var mapper = new RowMapper<Primitive>();
			mapper.MapColumn(1, f => f.IntValue).IgnoreErrors(false);
			List<ExceptionInfo> actual = new List<ExceptionInfo>();
			mapper.OnError((e) => actual.Add(e));
			FormatException expectedExcetpion = null;
			try
			{
				var rows = ExcelReader
					.GetRows(DefaultTestXlsx, "Primitives", mapper.FromRows(3, 6))
					.Select(e => e.IntValue)
					.ToArray();
			}
			catch (FormatException ex)
			{
				expectedExcetpion = ex;
			}

			var expected = new ExCell[] {
				new ExCell(4,0){Sheet="Primitives"}
			};

			Assert.IsNotNull(expectedExcetpion);
			Assert.AreEqual(1, actual.Count);

			CollectionAssert.AreEqual(expected, actual.Select(x => x.Cell).ToArray());
		}


		[TestMethod]
		public void ExcelReader_Rows_ReturnsExpected()
		{
			var filePath = Helper.ResourcePath("Primitives.xlsx");
			Stopwatch sw = new Stopwatch();
			sw.Start();
			var mapper = new RowMapper<Primitive>();
			//mapper.MapSheet("Primitives");
			List<ExceptionInfo> exceptions = new List<ExceptionInfo>();
			mapper.FromRows(3, 6);
			mapper.OnError((e) => exceptions.Add(e));
			//mapper.Sheet("Primitives")
			//	.Range(cell => cell.Row >= 3 && cell.Row <= 6);
			//mapper.Range((cell) =>
			//	{ return cell.Sheet == "Primitives" && cell.Row >= 3 && cell.Row <= 6; }
			//);

			//mapper.MapColumn(0, f => f.IntValue)
			//	.OnError((e) => exceptions.Add(e));

			mapper.MapColumn(0, (i, v) => i.IntValue = Convert.ToInt32(v));

			mapper.MapColumn(1, f => f.StringValue);

			mapper.MapColumn(2, f => f.DecimalValue);

			var actual = ExcelReader.GetRows<Primitive>(
				filePath, "Primitives", mapper)
				.ToArray();

			sw.Stop();
			Console.WriteLine($"Elapsed time: {sw.Elapsed}");
			foreach (var item in actual)
			{
				Console.WriteLine(item);
			}
			foreach (var item in exceptions)
			{
				Console.WriteLine($"{item.Cell} {item.Message} '{item.RawValue}'");
			}
		}
	}

}