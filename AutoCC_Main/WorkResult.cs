using System.Collections.Generic;

namespace AutoCC_Main
{
	class WorkResult
	{
		public static bool m_result = false;
		public static string m_drawer_code = string.Empty;
		public static List<string> m_error_codes = null;

		public static void SetErrorCodes(string error_codes)
		{
			if (error_codes != "")
			{
				if (m_error_codes == null)
					m_error_codes = new List<string>();

				char[] separator = { '|' };
				string[] results = error_codes.Split(separator);
				foreach (var item in results)
				{
					m_error_codes.Add(item);
				}
			}
		}


		public static void AddErrorCode(string error_code)
		{
			if (m_error_codes == null)
				m_error_codes = new List<string>();

			m_error_codes.Add(error_code);
		}

		public static string ErrorCodesToString()
		{
			if (m_error_codes != null)
			{
				System.Text.StringBuilder error_string = new System.Text.StringBuilder();
				if (m_error_codes != null)
				{
					foreach (var error_code in m_error_codes)
					{
						if (error_string.Length != 0)
							error_string.Append("|");
						error_string.Append(error_code);
					}
				}

				return error_string.ToString();
			}

			return "";
		}

		public static void Clear()
		{
			m_result = false;
			m_drawer_code = string.Empty;
			m_error_codes = null;
		}
	}
}