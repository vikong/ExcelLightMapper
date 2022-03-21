using System;

namespace UnitTests
{
	public class Primitive
	{
		public Int32 Row { get; set; }
		public Int32? IntValue { get; set; }

		public String StringValue { get; set; }
		public Decimal? DecimalValue { get; set; }
		public Boolean? BoolValue { get; set; }
		public DateTime DateValue { get; set; }
		public Decimal? FuncColumn { get; set; }
	}
}