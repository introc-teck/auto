using System;
using System.Collections.Generic;

namespace AutoCC_Main
{
	public class WorkSchedule
	{
		public WorkSchedule()
		{
		}

		public class DayOfWeekComparer : IComparer<DayOfWeek>
		{
			public int Compare(DayOfWeek a, DayOfWeek b)
			{
				if (a.m_day_index < b.m_day_index)
					return -1;

				return 0;
			}
		}

		public class DayOfWeek
		{
			public int m_day_index = 0;
			public bool m_duty = false;
			public Time m_start_time = new Time(9, 0, 0);
			public Time m_end_time = new Time(18, 0, 0);

			public DayOfWeek()
			{
			}

			public bool IsInRange(Time time)
			{
				bool over_and_equal = false;
				if (time.Hour > m_start_time.Hour)
				{
					over_and_equal = true;
				}
				else if (time.Hour == m_start_time.Hour)
				{
					if (time.Minute > m_start_time.Minute)
					{
						over_and_equal = true;
					}
					else if (time.Minute == m_start_time.Minute)
					{
						if (time.Second >= m_start_time.Second)
						{
							over_and_equal = true;
						}
					}
				}

				if (over_and_equal == false)
					return false;

				bool under_and_equal = false;
				if (time.Hour < m_end_time.Hour)
				{
					under_and_equal = true;
				}
				else if (time.Hour == m_end_time.Hour)
				{
					if (time.Minute < m_end_time.Minute)
					{
						under_and_equal = true;
					}
					else if (time.Minute == m_end_time.Minute)
					{
						if (time.Second <= m_end_time.Second)
						{
							under_and_equal = true;
						}
					}
				}

				return (over_and_equal == true && under_and_equal == true);
			}
		}

		public class DateRangeComparer : IComparer<DateRange>
		{
			public int Compare(DateRange a, DateRange b)
			{
				if (a.m_start_date.Year < b.m_start_date.Year)
				{
					return -1;
				}
				else if (a.m_start_date.Year == b.m_start_date.Year)
				{
					if (a.m_start_date.Month < b.m_start_date.Month)
					{
						return -1;
					}
					else if (a.m_start_date.Month == b.m_start_date.Month)
					{
						if (a.m_start_date.Day < b.m_start_date.Day)
							return -1;
					}
				}

				return 0;
			}
		}

		public class DateRange
		{
			public Date m_start_date = null;
			public Date m_end_date = null;
			public bool m_duty = false;
			public Time m_start_time = new Time(9, 0, 0);
			public Time m_end_time = new Time(18, 0, 0);

			public DateRange()
			{
			}

			public bool IsInRange(Date date)
			{
				bool over_and_equal = false;
				if (date.Year > m_start_date.Year)
				{
					over_and_equal = true;
				}
				else if (date.Year == m_start_date.Year)
				{
					if (date.Month > m_start_date.Month)
					{
						over_and_equal = true;
					}
					else if (date.Month == m_start_date.Month)
					{
						if (date.Day >= m_start_date.Day)
						{
							over_and_equal = true;
						}
					}
				}

				if (over_and_equal == false)
					return false;

				bool under_and_equal = false;
				if (date.Year < m_end_date.Year)
				{
					under_and_equal = true;
				}
				else if (date.Year == m_end_date.Year)
				{
					if (date.Month < m_end_date.Month)
					{
						under_and_equal = true;
					}
					else if (date.Month == m_end_date.Month)
					{
						if (date.Day <= m_end_date.Day)
						{
							under_and_equal = true;
						}
					}
				}

				return (over_and_equal == true && under_and_equal == true);
			}

			public bool IsOverlapped(DateRange date_range)
			{
				Date date = date_range.m_start_date;
				while (true)
				{
					if (IsInRange(date) == true)
						return true;

					if (date >= date_range.m_end_date)
						break;

					DateTime temp = new DateTime(date.Year, date.Month, date.Day);
					DateTime next = temp.AddDays(1);
					date = new Date(next.Year, next.Month, next.Day);
				}

				return false;
			}

			public bool IsInRange(Time time)
			{
				bool over_and_equal = false;
				if (time.Hour > m_start_time.Hour)
				{
					over_and_equal = true;
				}
				else if (time.Hour == m_start_time.Hour)
				{
					if (time.Minute > m_start_time.Minute)
					{
						over_and_equal = true;
					}
					else if (time.Minute == m_start_time.Minute)
					{
						if (time.Second >= m_start_time.Second)
						{
							over_and_equal = true;
						}
					}
				}

				if (over_and_equal == false)
					return false;

				bool under_and_equal = false;
				if (time.Hour < m_end_time.Hour)
				{
					under_and_equal = true;
				}
				else if (time.Hour == m_end_time.Hour)
				{
					if (time.Minute < m_end_time.Minute)
					{
						under_and_equal = true;
					}
					else if (time.Minute == m_end_time.Minute)
					{
						if (time.Second <= m_end_time.Second)
						{
							under_and_equal = true;
						}
					}
				}

				return (over_and_equal == true && under_and_equal == true);
			}
		}

		public class Date
		{
			public Date()
			{
			}

			public Date(int year, int month, int day)
			{
				Year = year;
				Month = month;
				Day = day;
			}

			public Date(string value)
			{
				char[] delimiters = { '.' };
				string[] items = value.Split(delimiters);
				if (items.Length == 3)
				{
					Year = Convert.ToInt32(items[0]);
					Month = Convert.ToInt32(items[1]);
					Day = Convert.ToInt32(items[2]);
				}
			}

			public int Year { get; set; }
			public int Month { get; set; }
			public int Day { get; set; }

			public new string ToString()
			{
				string value = string.Format("{0:0000}.{1:00}.{2:00}", Year, Month, Day);
				return value;
			}

			public bool IsValid()
			{
				if (Year != 0 && Month != 0 && Day != 0)
					return true;

				return false;
			}

			public static bool operator <(Date date1, Date date2)
			{
				bool result = false;

				if (date1.Year < date2.Year)
				{
					result = true;
				}
				else if (date1.Year == date2.Year)
				{
					if (date1.Month < date2.Month)
					{
						result = true;
					}
					else if (date1.Month == date2.Month)
					{
						if (date1.Day < date2.Day)
						{
							result = true;
						}
					}
				}

				return result;
			}

			public static bool operator >(Date date1, Date date2)
			{
				bool result = false;

				if (date1.Year > date2.Year)
				{
					result = true;
				}
				else if (date1.Year == date2.Year)
				{
					if (date1.Month > date2.Month)
					{
						result = true;
					}
					else if (date1.Month == date2.Month)
					{
						if (date1.Day > date2.Day)
						{
							result = true;
						}
					}
				}

				return result;
			}

			public static bool operator <=(Date date1, Date date2)
			{
				bool result = false;

				if (date1.Year < date2.Year)
				{
					result = true;
				}
				else if (date1.Year == date2.Year)
				{
					if (date1.Month < date2.Month)
					{
						result = true;
					}
					else if (date1.Month == date2.Month)
					{
						if (date1.Day <= date2.Day)
						{
							result = true;
						}
					}
				}

				return result;
			}

			public static bool operator >=(Date date1, Date date2)
			{
				bool result = false;

				if (date1.Year > date2.Year)
				{
					result = true;
				}
				else if (date1.Year == date2.Year)
				{
					if (date1.Month > date2.Month)
					{
						result = true;
					}
					else if (date1.Month == date2.Month)
					{
						if (date1.Day >= date2.Day)
						{
							result = true;
						}
					}
				}

				return result;
			}
		}

		public class Time
		{
			public Time()
			{
			}

			public Time(int hour, int minute, int second)
			{
				Hour = hour;
				Minute = minute;
				Second = second;
			}

			public Time(string value)
			{
				char[] delimiters = { ':' };
				string[] items = value.Split(delimiters);
				if (items.Length == 2)
				{
					Hour = Convert.ToInt32(items[0]);
					Minute = Convert.ToInt32(items[1]);
				}
				else if (items.Length == 3)
				{
					Hour = Convert.ToInt32(items[0]);
					Minute = Convert.ToInt32(items[1]);
					Second = Convert.ToInt32(items[2]);
				}
			}

			public int Hour { get; set; }
			public int Minute { get; set; }
			public int Second { get; set; }

			public static double GetPosition(Time left, Time right, Time target)
			{
				int left_value = left.Hour * 60 + left.Minute;
				int right_value = right.Hour * 60 + right.Minute - left_value;
				int target_value = target.Hour * 60 + target.Minute - left_value;

				return ((double)target_value / (double)right_value);    // 0 ~ 1.0 (left ~ right)
			}

			public new string ToString()
			{
				string value = string.Format("{0:00}:{1:00}:{2:00}", Hour, Minute, Second);
				return value;
			}

			public static bool operator <(Time time1, Time time2)
			{
				bool result = false;

				if (time1.Hour < time2.Hour)
				{
					result = true;
				}
				else if (time1.Hour == time2.Hour)
				{
					if (time1.Minute < time2.Minute)
					{
						result = true;
					}
					else if (time1.Minute == time2.Minute)
					{
						if (time1.Second < time2.Second)
						{
							result = true;
						}
					}
				}

				return result;
			}

			public static bool operator >(Time time1, Time time2)
			{
				bool result = false;

				if (time1.Hour > time2.Hour)
				{
					result = true;
				}
				else if (time1.Hour == time2.Hour)
				{
					if (time1.Minute > time2.Minute)
					{
						result = true;
					}
					else if (time1.Minute == time2.Minute)
					{
						if (time1.Second > time2.Second)
						{
							result = true;
						}
					}
				}

				return result;
			}

			public static bool operator <=(Time time1, Time time2)
			{
				bool result = false;

				if (time1.Hour < time2.Hour)
				{
					result = true;
				}
				else if (time1.Hour == time2.Hour)
				{
					if (time1.Minute < time2.Minute)
					{
						result = true;
					}
					else if (time1.Minute == time2.Minute)
					{
						if (time1.Second <= time2.Second)
						{
							result = true;
						}
					}
				}

				return result;
			}

			public static bool operator >=(Time time1, Time time2)
			{
				bool result = false;

				if (time1.Hour > time2.Hour)
				{
					result = true;
				}
				else if (time1.Hour == time2.Hour)
				{
					if (time1.Minute > time2.Minute)
					{
						result = true;
					}
					else if (time1.Minute == time2.Minute)
					{
						if (time1.Second >= time2.Second)
						{
							result = true;
						}
					}
				}

				return result;
			}
		}

		public List<DayOfWeek> m_days_of_week = new List<DayOfWeek>();
		public List<DateRange> m_date_ranges = new List<DateRange>();

		public Time GetStartTime(DateTime day)
		{
			Date date = new Date(day.Year, day.Month, day.Day);
			for (int i = 0; i < m_date_ranges.Count; i++)
			{
				DateRange date_range = m_date_ranges[i];
				if (date_range.IsInRange(date) == true)
				{
					return date_range.m_start_time;
				}
			}

			int day_index = Convert.ToInt32(day.DayOfWeek);
			for (int i = 0; i < m_days_of_week.Count; i++)
			{
				DayOfWeek day_of_week = m_days_of_week[i];
				if (day_index == day_of_week.m_day_index)
				{
					return day_of_week.m_start_time;
				}
			}

			return null;
		}

		public Time GetEndTime(DateTime day)
		{
			Date date = new Date(day.Year, day.Month, day.Day);
			for (int i = 0; i < m_date_ranges.Count; i++)
			{
				DateRange date_range = m_date_ranges[i];
				if (date_range.IsInRange(date) == true)
				{
					return date_range.m_end_time;
				}
			}

			int day_index = Convert.ToInt32(day.DayOfWeek);
			for (int i = 0; i < m_days_of_week.Count; i++)
			{
				DayOfWeek day_of_week = m_days_of_week[i];
				if (day_index == day_of_week.m_day_index)
				{
					return day_of_week.m_end_time;
				}
			}

			return null;
		}

		public Pair<Time, Time> GetWorkTime(DateTime day)
		{
			Date date = new Date(day.Year, day.Month, day.Day);
			for (int i = 0; i < m_date_ranges.Count; i++)
			{
				DateRange date_range = m_date_ranges[i];
				if (date_range.IsInRange(date) == true)
				{
					return new Pair<Time, Time>(date_range.m_start_time, date_range.m_end_time);
				}
			}

			int day_index = Convert.ToInt32(day.DayOfWeek);
			for (int i = 0; i < m_days_of_week.Count; i++)
			{
				DayOfWeek day_of_week = m_days_of_week[i];
				if (day_index == day_of_week.m_day_index)
				{
					return new Pair<Time, Time>(day_of_week.m_start_time, day_of_week.m_end_time);
				}
			}

			return null;
		}

		public bool GetDuty(DateTime day)
		{
			int duty = -1;  // -1: none, 0: off-duty, 1: on-duty

			Date date = new Date(day.Year, day.Month, day.Day);
			for (int j = 0; j < m_date_ranges.Count; j++)
			{
				DateRange date_range = m_date_ranges[j];
				if (date_range.IsInRange(date) == true)
				{
					duty = (date_range.m_duty == true ? 1 : 0);
					break;
				}
			}

			if (duty == -1)
			{
				int day_index = Convert.ToInt32(day.DayOfWeek);
				for (int j = 0; j < m_days_of_week.Count; j++)
				{
					DayOfWeek day_of_week = m_days_of_week[j];
					if (day_index == day_of_week.m_day_index)
					{
						duty = (day_of_week.m_duty == true ? 1 : 0);
						break;
					}
				}
			}

			return (duty == 1 ? true : false);
		}

		public void GetWeekDuties(DateTime day, ref bool[] duties)
		{
			duties = new bool[7];

			DateTime first_day = day.AddDays(-((double)Convert.ToInt32(day.DayOfWeek)));    // sunday

			for (int i = 0; i < 7; i++)
			{
				DateTime the_day = first_day.AddDays((double)i);

				int duty = -1;  // -1: none, 0: off-duty, 1: on-duty

				Date date = new Date(the_day.Year, the_day.Month, the_day.Day);
				for (int j = 0; j < m_date_ranges.Count; j++)
				{
					DateRange date_range = m_date_ranges[j];
					if (date_range.IsInRange(date) == true)
					{
						duty = (date_range.m_duty == true ? 1 : 0);
						break;
					}
				}

				if (duty == -1)
				{
					int day_index = Convert.ToInt32(the_day.DayOfWeek);
					for (int j = 0; j < m_days_of_week.Count; j++)
					{
						DayOfWeek day_of_week = m_days_of_week[j];
						if (day_index == day_of_week.m_day_index)
						{
							duty = (day_of_week.m_duty == true ? 1 : 0);
							break;
						}
					}
				}

				duties[i] = (duty == 1 ? true : false);
			}
		}

		public bool CheckWorkTime(DateTime work_time)
		{
			bool duty = false;

			Date date = new Date(work_time.Year, work_time.Month, work_time.Day);
			Time time = new Time(work_time.Hour, work_time.Minute, work_time.Second);

			for (int j = 0; j < m_date_ranges.Count; j++)
			{
				DateRange date_range = m_date_ranges[j];
				if (date_range.IsInRange(date) == true)
				{
					if (date_range.m_duty == true)
					{
						if (date_range.IsInRange(time) == true)
						{
							duty = true;
						}
					}

					return duty;
				}
			}

			for (int i = 0; i < 7; i++)
			{
				int day_index = Convert.ToInt32(work_time.DayOfWeek);
				for (int j = 0; j < m_days_of_week.Count; j++)
				{
					DayOfWeek day_of_week = m_days_of_week[j];
					if (day_index == day_of_week.m_day_index)
					{
						if (day_of_week.m_duty == true)
						{
							if (day_of_week.IsInRange(time) == true)
							{
								duty = true;
								break;
							}

							break;
						}
					}
				}

				if (duty == true)
					break;
			}

			return duty;
		}
	}
}
