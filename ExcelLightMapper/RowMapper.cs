namespace ExcelLightMapper
{
	using System;
	using System.Collections.Generic;
	using System.Linq.Expressions;

	public class RowMapper<T>
	{
		private readonly Dictionary<int, ICellMap<T>> mappedColumns;

		private Func<int, bool> rowFilter;
		private Func<ExCell, bool> rangeFilter { get; set; }

		private Func<T> instanceCreator;

		public Action<ExceptionInfo> ExceptionHandler { get; set; }

		public Func<T> InstanceCreator
		{
			get { return instanceCreator; }
			set { instanceCreator = value ?? throw new ArgumentNullException(); }
		}

		public RowMapper()
		{
			instanceCreator = Expression.Lambda<Func<T>>(
				Expression.New(typeof(T).GetConstructor(Type.EmptyTypes)))
				.Compile();
			mappedColumns = new Dictionary<int, ICellMap<T>>();
			rangeFilter = _ => true;
			rowFilter = _ => true;
		}

		public void AddCellMap(ICellMap<T> cellMap)
		{
			mappedColumns.Add(cellMap.ColumnIndex, cellMap);
		}

		public bool IsRequestedRow(int rowIndex)
			=> rowFilter(rowIndex);

		public bool IsRequestedRange(ExCell cell)
			=> rangeFilter(cell);

		public RowMapper<T> Range(Expression<Func<ExCell, bool>> rangeFilter)
		{
			this.rangeFilter = rangeFilter.Compile();
			return this;
		}

		public RowMapper<T> FromRows(uint startingRow, uint? lastRow = null)
		{
			Expression<Func<int, bool>> rowFilter;
			if (lastRow.HasValue)
			{
				rowFilter = r => r >= startingRow - 1 && r <= lastRow.Value - 1;
			}
			else
			{
				rowFilter = r => r >= startingRow - 1;
			}
			this.rowFilter = rowFilter.Compile();
			return this;
		}

		public bool TryGetCellMap(int columnIndex, out ICellMap<T> cellMap)
		{
			return mappedColumns.TryGetValue(columnIndex, out cellMap);
		}

		public ICellMapper<T, TProp> MapColumn<TProp>(string columnName, Expression<Func<T, TProp>> mapExpression)
		{
			return MapColumn(ExcelReader.GetColumnNumber(columnName), mapExpression);
		}

		public ICellMapper<T, TProp> MapColumn<TProp>(int column, Expression<Func<T, TProp>> mapExpression)
		{
			var cMap = new CellMap<T, TProp>(mapExpression) { ColumnIndex = column - 1 };
			AddCellMap(cMap);

			return cMap;
		}

		public ICellMapper<T> MapColumn(int column, Action<T, object> mapAction)
		{
			var cMap = new CellMap<T>(mapAction) { ColumnIndex = column - 1 };
			AddCellMap(cMap);

			return cMap;
		}
	}

	public static class RowMapperExtensions
	{
		public static RowMapper<T> OnError<T>(this RowMapper<T> mapper, Action<ExceptionInfo> errorHandler)
		{
			mapper.ExceptionHandler = errorHandler;
			return mapper;
		}

		//public static RowMapper<T> MapSheet<T>(this RowMapper<T> mapper, String sheetName)
		//{
		//	Func<string, bool> sheetFiler =
		//		(sh) => String.Compare(sh, sheetName, true)==0;
		//	mapper.SheetFilter = sheetFiler;
		//	return mapper;
		//}
		//public static RowMapper<T> MapSheets<T>(this RowMapper<T> mapper, IEnumerable<String> sheetNames)
		//{
		//	HashSet<string> sheets = new HashSet<string>(sheetNames.Select(s=>s.Trim().ToUpperInvariant()));
		//	Func<string, bool> sheetFiler =
		//		(sh) => sheets.Contains(sh.ToUpperInvariant());

		//	mapper.SheetFilter = sheetFiler;
		//	return mapper;
		//}
	}

	//public class RowMapper<TOut>: IRowMapper<TOut>
	//   {
	//       public String Sheet { get; set; }
	//       private Dictionary<Int32, Action<TOut, Object>> mapFunctions;
	//       public RowMapper()
	//       {
	//           mapFunctions = new Dictionary<int, Action<TOut, Object>>();
	//       }

	//       public RowMapper<TOut> MapColumn(Int32 column, Action<TOut, Object> func)
	//       {
	//           mapFunctions.Add(column, func);
	//           return this;
	//       }

	//       public RowMapper<TOut> MapColumn(Int32 column, Expression<Func<TOut, String>> propertyExpression)
	//       {
	//           var setter = ExpressionExt.GetSetter(propertyExpression);
	//           Action<TOut, Object> func = (t, o) => {
	//               //MemberInfo member = ExpressionExt.GetMemberInfo(propertyExpression);
	//               //var parameter = Expression.Parameter(typeof(double), "value");
	//               //Console.WriteLine(parameter.Name);
	//               setter(t, (String)o);
	//           };
	//           return MapColumn(column, func);
	//       }
	//       public bool TryGetColumnMapper(Int32 column, out Action<TOut, Object> converter)
	//       {
	//           return mapFunctions.TryGetValue(column, out converter);
	//       }

	//   }
}