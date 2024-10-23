using System;
using System.Collections.Generic;
using System.Configuration;

namespace AutoCC_Main
{
	class AppConfiguration
	{





		public static Dictionary<string, string> config = null; //FormSettings 화면에 있는 모든 정보를 저장하는곳

		public static string[] toDay_RunTime = null; // 동작시간 저장

		public static Machine machine = new Machine();




		public static void SaveConfig(string _prop, string _val)
		{
			try
			{

				var configfile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
				var appSetting = configfile.AppSettings.Settings;


				if (appSetting[_prop] == null)
				{
					appSetting.Add(_prop, _val);
				}
				else
				{
					appSetting.Remove(_prop);
					appSetting.Add(_prop, _val);
				}

				configfile.Save(ConfigurationSaveMode.Modified);
				ConfigurationManager.RefreshSection(configfile.AppSettings.SectionInformation.Name);


			}
			catch (Exception ex)
			{

			}
		}

		public static string LoadConfig(string _prop)
		{
			try
			{


				var appSetting = ConfigurationManager.AppSettings;

				string result;

				result = appSetting[_prop] ?? string.Empty;

				if (result != string.Empty) return result;





			}
			catch (Exception ex)
			{

			}

			return "null";
		}


		public static string GetToDay(DateTime dt)
		{


			int nFirstDay = Int32.Parse(dt.DayOfWeek.ToString("d"));



			string retYoil = string.Empty;

			switch (nFirstDay)

			{

				case 0:
					retYoil = "Sunday";

					break;

				case 1:
					retYoil = "Monday";

					break;

				case 2:
					retYoil = "Tuesday";

					break;

				case 3:
					retYoil = "Wednesday";

					break;

				case 4:
					retYoil = "Thursday";

					break;

				case 5:
					retYoil = "Friday";

					break;

				case 6:
					retYoil = "Saturday";

					break;

			}






			return retYoil;
		}


		public static bool chk_work_start()
		{


			bool _rest = false;


			DateTime dt = new DateTime(int.Parse(DateTime.Now.ToString("yyyy")), int.Parse(DateTime.Now.ToString("MM")), int.Parse(DateTime.Now.ToString("dd")),
										int.Parse(DateTime.Now.ToString("HH")), int.Parse(DateTime.Now.ToString("mm")), int.Parse(DateTime.Now.ToString("ss")));



			if (toDay_RunTime == null) toDay_RunTime = LoadConfig(GetToDay(dt)).Split(',');



			if (toDay_RunTime[0].Equals("T"))
			{

				// Console.WriteLine(string.Format("{0:D2}", int.Parse(toDay_RunTime[1])) + "-" + string.Format("{0:D2}", int.Parse(toDay_RunTime[2])));



				DateTime _sd = new DateTime(int.Parse(DateTime.Now.ToString("yyyy")), int.Parse(DateTime.Now.ToString("MM")), int.Parse(DateTime.Now.ToString("dd")),
											int.Parse(toDay_RunTime[1]), int.Parse(toDay_RunTime[2]), 0);


				DateTime _ed = new DateTime(int.Parse(DateTime.Now.ToString("yyyy")), int.Parse(DateTime.Now.ToString("MM")), int.Parse(DateTime.Now.ToString("dd")),
											int.Parse(toDay_RunTime[3]), int.Parse(toDay_RunTime[4]), 0);

				if (dt.TimeOfDay > _sd.TimeOfDay && dt.TimeOfDay < _ed.TimeOfDay)
					_rest = true;
				else
					_rest = false;



				Console.WriteLine(_rest.ToString() + "결과");

			}

			return _rest;


		}



	}
}

