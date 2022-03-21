namespace ExcelLightMapper
{
	using System;
	using System.Linq.Expressions;

	public interface ICellMap<T>
	{
		int ColumnIndex { get; }
		bool IgnoreErrors { get; set; }

		void SetValue(T instance, object value);
	}

	public class CellMap<T> : ICellMap<T>, ICellMapper<T>
	{
		public int ColumnIndex { get; set; }
		protected Action<T, object> Setter { get; set; }
		public bool IgnoreErrors { get; set; } = true;

		public CellMap(Action<T, object> setter)
		{
			Setter = setter;
		}

		public void SetValue(T instance, object value)
		{
			Setter(instance, value);
		}

		ICellMapper<T> ICellMapper<T>.IgnoreErrors(bool ignore)
		{
			IgnoreErrors = ignore;
			return this;
		}
	}

	public class CellMap<T, TProp> : ICellMap<T>, ICellMapper<T, TProp>
	{
		public int ColumnIndex { get; set; }
		protected Action<T, TProp> Setter { get; set; }
		protected Func<object, TProp> Converter { get; set; }
		public bool IgnoreErrors { get; set; } = true;

		public CellMap()
		{ }

		public CellMap(Expression<Func<T, TProp>> propertySelector)
		{
			Converter = Helper.ConvertTo<TProp>;
			Setter = Helper.GetSetter(propertySelector);
		}

		ICellMapper<T> ICellMapper<T>.IgnoreErrors(bool ignore)
		{
			IgnoreErrors = ignore;
			return this;
		}

		public void SetValue(T instance, object value)
		{
			TProp TPropValue = Converter(value);
			Setter(instance, TPropValue);
		}

		ICellMapper<T, TProp> ICellMapper<T, TProp>.From(Func<object, TProp> valueConverter)
		{
			Converter = valueConverter;
			return this;
		}
	}
}