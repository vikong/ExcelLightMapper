using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace UnitTests
{
	internal static class Helper
	{
		public static String ResourcePath(String name)
		{ 
			return Path.GetFullPath(Path.Combine("Resources", name));
		}
	}
}
