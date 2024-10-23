using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCC_Sub
{
	class DimUnit
	{
		public static double PtToMM(double value)
		{
			return value * 0.352777778;
		}

		public static double MMToPt(double value)
		{
			return value * 2.83464566929134;
		}

		public static double InToMM(double value)
		{
			return value * 25.4;
		}

		public static double MMToIn(double value)
		{
			return value * 0.03937;
		}
	}
}
