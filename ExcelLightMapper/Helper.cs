namespace ExcelLightMapper
{
	using System;
	using System.Globalization;
	using System.Linq.Expressions;
	using System.Reflection;
	using System.Runtime.CompilerServices;
	using System.ComponentModel.DataAnnotations;
	using System.Linq;

	public static class Helper
	{
		[MethodImpl(MethodImplOptions.AggressiveInlining)]
		public static T ConvertTo<T>(object val)
		{
			if (val is T tVal)
			{
				return tVal;
			}
			else if (val == null && (!typeof(T).IsValueType || Nullable.GetUnderlyingType(typeof(T)) != null))
			{
				return default;
			}
			else if (val is Array array && typeof(T).IsArray)
			{
				var elementType = typeof(T).GetElementType();
				var result = Array.CreateInstance(elementType, array.Length);
				for (int i = 0; i < array.Length; i++)
					result.SetValue(Convert.ChangeType(array.GetValue(i), elementType, CultureInfo.InvariantCulture), i);
				return (T)(object)result;
			}
			else
			{
				var convertToType = Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T);
				return (T)Convert.ChangeType(val, convertToType, CultureInfo.InvariantCulture);
			}

		}

		/// <summary>
		/// Convert a lambda expression for a getter into a setter
		/// </summary>
		public static Action<T, TProperty> GetSetter<T, TProperty>(Expression<Func<T, TProperty>> expression)
		{
			var memberExpression = (MemberExpression)expression.Body;
			var property = (PropertyInfo)memberExpression.Member;
			var setMethod = property.GetSetMethod();

			var parameterT = Expression.Parameter(typeof(T), "x");
			var parameterTProperty = Expression.Parameter(typeof(TProperty), "y");

			var newExpression =
				Expression.Lambda<Action<T, TProperty>>(
					Expression.Call(parameterT, setMethod, parameterTProperty),
					parameterT,
					parameterTProperty
				);

			return newExpression.Compile();
		}

		static Expression<Action<T, object>> MakeSetter<T>(Expression<Func<T, object>> property)
		{
			var member = (MemberExpression)property.Body;
			var source = property.Parameters[0];
			var value = Expression.Parameter(typeof(object), "value");
			var body = Expression.Assign(member, Expression.Convert(value, member.Type));
			return Expression.Lambda<Action<T, object>>(body, source, value);
		}

		public static string GetDisplayName(this Enum value)
		{
			var attribute = value.GetType()
				.GetField(value.ToString())
				.GetCustomAttributes(typeof(DisplayAttribute), false)
				.FirstOrDefault() as DisplayAttribute;

			return attribute == null ? value.ToString() : attribute.GetName();
		}

	}

	public static class ExpressionExtensions
	{
		public static MemberInfo GetMemberInfo<T, U>(Expression<Func<T, U>> expression)
		{
			var member = expression.Body as MemberExpression;
			if (member != null)
				return member.Member;

			throw new ArgumentException("Expression is not a member access", "expression");
		}

		private static void SetPropertyValue(object target, string propName, object value)
		{
			var propInfo = target.GetType().GetProperty(propName,
								 BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly);

			if (propInfo != null)
			{
				propInfo.SetValue(target, value, null);
			}
		}
	}

}
