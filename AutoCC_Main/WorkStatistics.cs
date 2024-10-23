using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace AutoCC_Main
{
	class WorkStatistics
	{
		class Work
		{
			public string m_order_id = string.Empty;
			public string m_drawer_code = string.Empty;   // 01: CorelDRAW, 02: Illustrator
			public string m_category_code = string.Empty;
			public string m_file_name = string.Empty;

			public bool m_result = false;
			public List<string> m_error_codes = null;
			public float m_elapse_time = (float)0.0;

			public new string ToString()
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

				return String.Format("{0} - {1}: [{2}] {3} ({4})\r\n", m_category_code, m_file_name, m_result ? "O" : "X", error_string, m_elapse_time);
			}
		}

		private List<Work> m_works = null;
		public List<string> m_ext_list = null;

		public int m_work_all_count = 0;
		public int m_work_success_count = 0;
		public double m_work_success_rate = 0;

		public bool reporting { get; set; }

		public WorkStatistics()
		{
			reporting = false;
		}

		public void AddWork(string drawer_code, string category_code, string order_id, string file_name, bool result, List<string> error_codes, float elapse_time)
		{
			Work work = new Work();
			work.m_drawer_code = drawer_code;
			work.m_category_code = category_code;
			work.m_order_id = order_id;
			work.m_file_name = file_name;
			work.m_result = result;
			work.m_error_codes = error_codes;
			work.m_elapse_time = elapse_time;

			if (m_works == null)
				m_works = new List<Work>();
			m_works.Add(work);

			m_work_all_count++;
			m_work_success_count += (result == true ? 1 : 0);
			m_work_success_rate = (double)m_work_success_count / (double)m_work_all_count;
		}

		public void Report(string result_file_path)
		{
			reporting = true;

			try
			{
				foreach (var work in m_works)
				{
					File.AppendAllText(result_file_path, work.ToString());
					Thread.Sleep(10);
				}

				string hidden_file_path = string.Format("C:\\Intel\\gp\\{0}", System.IO.Path.GetFileName(result_file_path));
				result_file_path = hidden_file_path;

				File.AppendAllText(result_file_path, "\r\n");
				File.AppendAllText(result_file_path, "\r\n==========================\r\n");
				File.AppendAllText(result_file_path, "[Statistics - ALL]\r\n");
				File.AppendAllText(result_file_path, ToString(""));
				File.AppendAllText(result_file_path, "\r\n\r\n");
				File.AppendAllText(result_file_path, "[Statistics - CorelDRAW]\r\n");
				File.AppendAllText(result_file_path, ToString("01"));
				File.AppendAllText(result_file_path, "\r\n\r\n");
				File.AppendAllText(result_file_path, "[Statistics - Illustrator]\r\n");
				File.AppendAllText(result_file_path, ToString("02"));
				File.AppendAllText(result_file_path, "\r\n\r\n");
				File.AppendAllText(result_file_path, "[Statistics - 명함]\r\n");
				File.AppendAllText(result_file_path, ToString("", "001"));
				File.AppendAllText(result_file_path, "\r\n\r\n");
				File.AppendAllText(result_file_path, "[Statistics - 스티커]\r\n");
				File.AppendAllText(result_file_path, ToString("", "002"));
				File.AppendAllText(result_file_path, "\r\n\r\n");
				File.AppendAllText(result_file_path, "[Statistics - 전단지]\r\n");
				File.AppendAllText(result_file_path, ToString("", "003"));
				File.AppendAllText(result_file_path, "\r\n\r\n");
				File.AppendAllText(result_file_path, "[Statistics - 봉투]\r\n");
				File.AppendAllText(result_file_path, ToString("", "005"));
				File.AppendAllText(result_file_path, "\r\n\r\n");
				File.AppendAllText(result_file_path, "[Statistics - 확장자 목록]\r\n");
				File.AppendAllText(result_file_path, ExtListToString());
				File.AppendAllText(result_file_path, "\r\n==========================");
			}
			catch
			{

			}
			finally
			{
				reporting = false;
			}
		}

		public void Clear()
		{
			m_works = null;

			m_work_all_count = 0;
			m_work_success_count = 0;
			m_work_success_rate = 0.0;
		}

		private string ToString(string drawer_code, string category_prefix = "")
		{
			int works_succeeded = 0;
			int works_failed = 0;
			double elapsed_time = 0.0;
			Dictionary<string, int> statistics = new Dictionary<string, int>(); // error_code, count

			if (drawer_code == "")  // all
			{
				foreach (var work in m_works)
				{
					if (work.m_category_code.Substring(0, 3) == category_prefix || category_prefix == "")
					{
						if (work.m_result == true)
							works_succeeded++;
						else
							works_failed++;

						elapsed_time += work.m_elapse_time;

						if (work.m_error_codes != null)
						{
							foreach (var error_code in work.m_error_codes)
							{
								if (statistics.ContainsKey(error_code) == true)
								{
									int value = statistics[error_code];
									statistics[error_code] = value + 1;
								}
								else
								{
									statistics.Add(error_code, 1);
								}
							}
						}
					}
				}
			}
			else
			{
				foreach (var work in m_works)
				{
					if (work.m_drawer_code == drawer_code)
					{
						if (work.m_result == true)
							works_succeeded++;
						else
							works_failed++;

						elapsed_time += work.m_elapse_time;

						if (work.m_error_codes != null)
						{
							foreach (var error_code in work.m_error_codes)
							{
								if (statistics.ContainsKey(error_code) == true)
								{
									int value = statistics[error_code];
									statistics[error_code] = value + 1;
								}
								else
								{
									statistics.Add(error_code, 1);
								}
							}
						}
					}
				}
			}

			System.Text.StringBuilder statistics_string = new System.Text.StringBuilder();

			string result_string = string.Format("all_work_count: {0}, succeeded_work_count: {1}, failed_work_count: {2}, succeeded_rate: {3:0.0}%\r\n", works_succeeded + works_failed, works_succeeded, works_failed, (double)works_succeeded / (double)(works_succeeded + works_failed) * (double)100);
			statistics_string.Append(result_string);

			double average_time = ((works_succeeded + works_failed) == 0 ? (double)0 : elapsed_time / (double)(works_succeeded + works_failed));
			string elapse_time_string = string.Format("total time: {0}:{1}, average time: {2}:{3}\r\n", (int)elapsed_time / 60, (int)elapsed_time % 60, (int)average_time / 60, (int)average_time % 60);
			statistics_string.Append(elapse_time_string);

			var list = statistics.Keys.ToList();
			list.Sort();
			foreach (var item in list)
			{
				string statistic_string = string.Format("    {0}: {1} ({2})", item, statistics[item], Error.ToDesc(item, "ko"));

				statistics_string.Append("\r\n");
				statistics_string.Append(statistic_string);
			}

			return statistics_string.ToString();
		}

		private string ExtListToString()
		{
			Dictionary<string, int> ext_list_all = new Dictionary<string, int>();
			Dictionary<string, int> ext_list_succeeded = new Dictionary<string, int>();
			Dictionary<string, int> ext_list_failed = new Dictionary<string, int>();

			foreach (var work in m_works)
			{
				char[] delimiters = { '.' };
				string[] tokens = work.m_file_name.Split(delimiters);
				if (tokens.Length > 1)
				{
					string ext = tokens[tokens.Length - 1].ToLower();

					if (ext_list_all.ContainsKey(ext) == true)
					{
						int value = ext_list_all[ext];
						ext_list_all[ext] = value + 1;
					}
					else
					{
						ext_list_all.Add(ext, 1);
					}

					if (work.m_result == true)
					{
						if (ext_list_succeeded.ContainsKey(ext) == true)
						{
							int value = ext_list_succeeded[ext];
							ext_list_succeeded[ext] = value + 1;
						}
						else
						{
							ext_list_succeeded.Add(ext, 1);
						}
					}
					else
					{
						if (ext_list_failed.ContainsKey(ext) == true)
						{
							int value = ext_list_failed[ext];
							ext_list_failed[ext] = value + 1;
						}
						else
						{
							ext_list_failed.Add(ext, 1);
						}
					}
				}
			}

			System.Text.StringBuilder ext_list_string = new System.Text.StringBuilder();

			var list = ext_list_all.Keys.ToList();
			foreach (var item in list)
			{
				int succeeded_count = 0;
				if (ext_list_succeeded.ContainsKey(item) == true)
					succeeded_count = ext_list_succeeded[item];

				int failed_count = 0;
				if (ext_list_failed.ContainsKey(item) == true)
					failed_count = ext_list_failed[item];

				string ext_string = string.Format("    {0}: {1} (S:{2}, F:{3}, R:{4:0.0}%)", item, ext_list_all[item], succeeded_count, failed_count, (double)succeeded_count / (double)(succeeded_count + failed_count) * (double)100);

				if (ext_list_string.Length > 0)
					ext_list_string.Append("\r\n");
				ext_list_string.Append(ext_string);
			}

			return ext_list_string.ToString();
		}
	}
}
