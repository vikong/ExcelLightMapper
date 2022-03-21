using System;

namespace ExcelLightMapper
{
	public interface ICellMapper<T>
	{
		ICellMapper<T> IgnoreErrors(bool ignore = true);
	}

	public interface ICellMapper<T, TProp> : ICellMapper<T>
	{
		ICellMapper<T, TProp> From(Func<object, TProp> valueConverter);
	}


}