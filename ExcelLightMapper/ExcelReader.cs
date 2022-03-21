namespace ExcelLightMapper
{
	using System;
	using System.Collections.Generic;
	using System.IO;
	using ExcelDataReader;

	//public class ExcelData
	//{
	//	public string Sheet { get; set; }
	//	public object[] Columns { get; set; }
	//}

	public static class ExcelReader
	{
		/// <summary>
		/// Converts Excel's column name to column number
		/// </summary>
		/// <param name="name">Column name</param>
		/// <returns>Column number</returns>
		/// <example>GetColumnNumber("B"); //returns 2</example>
		public static int GetColumnNumber(string name)
		{
			int number = 0;
			int pow = 1;
			name = name.ToUpper();
			for (int i = name.Length - 1; i >= 0; i--)
			{
				number += (name[i] - 'A' + 1) * pow;
				pow *= 26;
			}

			return number;
		}

		public static IExcelDataReader OpenReader(FileStream stream)
		{
			return ExcelReaderFactory.CreateReader(stream);
		}

		public static IExcelDataReader OpenReader(string filePath)
		{
			IExcelDataReader excelReader;
			var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

			if (Path.GetExtension(filePath).ToUpper() == ".XLS")
			{
				//1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
				excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
			}
			else
			{
				//1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
				excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
			}
			return excelReader;
		}

		/// <summary>
		/// Moves Excel reader to sheet (moves forvard only!)
		/// </summary>
		/// <param name="reader">Reader</param>
		/// <param name="sheetName">Name of sheet</param>
		/// <returns>true if sheet was positioned</returns>
		private static bool MoveToSheet(this IExcelDataReader reader, string sheetName)
		{
			while (!string.Equals(reader.Name, sheetName, StringComparison.InvariantCultureIgnoreCase)
				&& reader.NextResult())
			{ }
			return string.Equals(reader.Name, sheetName, StringComparison.InvariantCultureIgnoreCase);
		}

		public static IEnumerable<T> GetRows<T>(IExcelDataReader reader, string sheetName, RowMapper<T> mapper)
		{
			if (sheetName != null && !reader.MoveToSheet(sheetName))
			{
				yield break;
			}

			Func<T> instanceCreator = mapper.InstanceCreator;
			int row = 0;
			string currentSheet = null;
			do
			{
				while (reader.Read())
				{
					if (reader.Name == currentSheet)
					{
						row++;
					}
					else
					{
						currentSheet = reader.Name;
						if (sheetName != null && !sheetName.Equals(currentSheet, StringComparison.InvariantCultureIgnoreCase))
						{
							yield break;
						}
						row = 0;
					}

					if (!mapper.IsRequestedRow(row)) { continue; }

					T instance = instanceCreator();
					int fieldCount = reader.FieldCount;
					for (int column = 0; column < fieldCount; column++)
					{
						if (!mapper.IsRequestedRange(new ExCell(row, column) { Sheet = currentSheet }))
						{
							continue;
						}
						if (mapper.TryGetCellMap(column, out ICellMap<T> cellMap))
						{
							var val = reader.GetValue(column);
							try
							{
								cellMap.SetValue(instance, val);
							}
							catch (Exception ex)
							{
								if (mapper.ExceptionHandler != null)
								{
									var exInfo = new ExceptionInfo
									{
										Cell = new ExCell(row, column) { Sheet = currentSheet },
										Message = ex.Message,
										RawValue = val
									};
									mapper.ExceptionHandler(exInfo);
								}
								if (!cellMap.IgnoreErrors) { throw; }
							}
						}
					}
					yield return instance;
				}
			} while (reader.NextResult());
		}

		public static IEnumerable<T> GetRows<T>(string filePath, string sheetName, RowMapper<T> mapper)
		{
			using (IExcelDataReader reader = OpenReader(filePath))
			{
				foreach (var item in GetRows(reader, sheetName, mapper))
				{
					yield return item;
				}
			}
		}

		public static IEnumerable<T> GetRows<T>(string filePath, RowMapper<T> mapper)
			=> GetRows(filePath, null, mapper);

		//public static List<ExcelData> Read(string filePath)
		//{
		//	List<ExcelData> result = new List<ExcelData>();
		//	using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
		//	{
		//		using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
		//		{
		//			do
		//			{
		//				while (reader.Read())
		//				{
		//					var row = new ExcelData() { Sheet = reader.Name };
		//					int fieldCount = reader.FieldCount;
		//					object[] col = new object[fieldCount];
		//					for (int i = 0; i < fieldCount; i++)
		//					{
		//						col[i] = reader.GetValue(i);
		//					}
		//					row.Columns = col;
		//					result.Add(row);
		//				}
		//			} while (reader.NextResult());
		//		}
		//	}
		//	return result;
		//}
	}
}