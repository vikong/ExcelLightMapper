# ExcelLightMapper
Lightweight mapper for [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)
The library reads rows from Excel sheet(s) and maps it to an your POCOs objects.
A flexible fluent mapping system allows you to customize the way the row is mapped to an object.
```csharp
public class PocoObject
{
    public Int32? IntValue { get; set; }
    public String StringValue { get; set; }
    public Decimal? DecimalValue { get; set; }
    public Boolean? BoolValue { get; set; }
    public DateTime DateValue { get; set; }
}

//...

// create map
var mapper = new RowMapper<Primitive>();

// specify required range, if need
mapper.FromRows(3);

// simple mapping
mapper.MapColumn(1, f => f.DateValue);
			
// map with column name
mapper.MapColumn("B", f => f.IntValue);

// custom mapping
mapper.MapColumn("C", f => f.StringValue).From(v => v.ToString());

// custom function mapping
mapper.MapColumn(4, (o, val) => {
    var d = Convert.ToDecimal(val);
    o.DecimalValue += d;
});

// treat convertion errors
List<ExceptionInfo> actual = new List<ExceptionInfo>();
mapper.OnError((e) => actual.Add(e));

// read rows
IEnumerable<Primitive> rows = ExcelReader.GetRows("Excel_Data.xlsx", "SheetName", mapper);

```
