using System;
using System.Windows.Forms;

namespace AutoCC_Main
{
	public partial class FormSettings : Form
	{


		public FormSettings()
		{
			InitializeComponent();

			this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
			this.SetStyle(ControlStyles.EnableNotifyMessage, true);

		}

		protected override void OnLoad(EventArgs e)
		{
			base.OnLoad(e);

			DateTime today = DateTime.Now;

			// machine info

			string _val = AppConfiguration.LoadConfig("Machine_info");

			if (!_val.Equals("null"))
			{
				textBoxMachineIndex.Text = _val.Split('/')[0];
				textBoxMachineCount.Text = _val.Split('/')[1];
			}




			time_Info1.set_name("일");
			time_Info2.set_name("월");
			time_Info3.set_name("화");
			time_Info4.set_name("수");
			time_Info5.set_name("목");
			time_Info6.set_name("금");
			time_Info7.set_name("토");



			time_Info1.set_date(AppConfiguration.LoadConfig("Sunday"));
			time_Info2.set_date(AppConfiguration.LoadConfig("Monday"));
			time_Info3.set_date(AppConfiguration.LoadConfig("Tuesday"));
			time_Info4.set_date(AppConfiguration.LoadConfig("Wednesday"));
			time_Info5.set_date(AppConfiguration.LoadConfig("Thursday"));
			time_Info6.set_date(AppConfiguration.LoadConfig("Friday"));
			time_Info7.set_date(AppConfiguration.LoadConfig("Saturday"));


			if (!AppConfiguration.LoadConfig("Root_Folder").Equals("null"))
			{
				textBox1.Text = AppConfiguration.LoadConfig("Root_Folder");
			}
			else
			{
				AppConfiguration.SaveConfig("Root_Folder", textBox1.Text);

			}


			if (!AppConfiguration.LoadConfig("Target_Folder").Equals("null"))
			{
				textBox2.Text = AppConfiguration.LoadConfig("Target_Folder");
			}
			else
			{
				AppConfiguration.SaveConfig("Target_Folder", textBox2.Text);

			}

			if (!AppConfiguration.LoadConfig("Preview_Folder").Equals("null"))
			{
				textBox3.Text = AppConfiguration.LoadConfig("Preview_Folder");
			}
			else
			{
				AppConfiguration.SaveConfig("Preview_Folder", textBox3.Text);
			}

			if (!AppConfiguration.LoadConfig("Deposit_Folder").Equals("null"))
			{
				textBox4.Text = AppConfiguration.LoadConfig("Deposit_Folder");
			}
			else
			{
				AppConfiguration.SaveConfig("Deposit_Folder", textBox4.Text);
			}

			if (!AppConfiguration.LoadConfig("Ordered_folder").Equals("null"))
			{
				textBox5.Text = AppConfiguration.LoadConfig("Ordered_folder");
			}
			else
			{
				AppConfiguration.SaveConfig("Ordered_folder", textBox5.Text);
			}

			if (!AppConfiguration.LoadConfig("Result_folder").Equals("null"))
			{
				textBox6.Text = AppConfiguration.LoadConfig("Result_folder");
			}
			else
			{
				AppConfiguration.SaveConfig("Result_folder", textBox6.Text);
			}




		}

		private void time_Info3_Load(object sender, EventArgs e)
		{

		}

		private void button1_Click(object sender, EventArgs e)
		{



			AppConfiguration.SaveConfig("Sunday", time_Info1.get_Enable() + "," +
														time_Info1.get_numericUpDownStartHour() + "," +
														time_Info1.get_numericUpDownStartMinute() + "," +
														time_Info1.get_numericUpDownEndHour() + "," +
														time_Info1.get_numericUpDownEndMinute());



			AppConfiguration.SaveConfig("Monday", time_Info2.get_Enable() + "," +
														time_Info2.get_numericUpDownStartHour() + "," +
														time_Info2.get_numericUpDownStartMinute() + "," +
														time_Info2.get_numericUpDownEndHour() + "," +
														time_Info2.get_numericUpDownEndMinute());



			AppConfiguration.SaveConfig("Tuesday", time_Info3.get_Enable() + "," +
														time_Info3.get_numericUpDownStartHour() + "," +
														time_Info3.get_numericUpDownStartMinute() + "," +
														time_Info3.get_numericUpDownEndHour() + "," +
														time_Info3.get_numericUpDownEndMinute());

			AppConfiguration.SaveConfig("Wednesday", time_Info4.get_Enable() + "," +
														time_Info4.get_numericUpDownStartHour() + "," +
														time_Info4.get_numericUpDownStartMinute() + "," +
														time_Info4.get_numericUpDownEndHour() + "," +
														time_Info4.get_numericUpDownEndMinute());

			AppConfiguration.SaveConfig("Thursday", time_Info5.get_Enable() + "," +
														time_Info5.get_numericUpDownStartHour() + "," +
														time_Info5.get_numericUpDownStartMinute() + "," +
														time_Info5.get_numericUpDownEndHour() + "," +
														time_Info5.get_numericUpDownEndMinute());

			AppConfiguration.SaveConfig("Friday", time_Info6.get_Enable() + "," +
														time_Info6.get_numericUpDownStartHour() + "," +
														time_Info6.get_numericUpDownStartMinute() + "," +
														time_Info6.get_numericUpDownEndHour() + "," +
														time_Info6.get_numericUpDownEndMinute());

			AppConfiguration.SaveConfig("Saturday", time_Info7.get_Enable() + "," +
														time_Info7.get_numericUpDownStartHour() + "," +
														time_Info7.get_numericUpDownStartMinute() + "," +
														time_Info7.get_numericUpDownEndHour() + "," +
														time_Info7.get_numericUpDownEndMinute());

			AppConfiguration.SaveConfig("Machine_info", textBoxMachineIndex.Text + "/" + textBoxMachineCount.Text);


			AppConfiguration.SaveConfig("Root_Folder", textBox1.Text);
			AppConfiguration.SaveConfig("Target_Folder", textBox2.Text);
			AppConfiguration.SaveConfig("Preview_Folder", textBox3.Text);
			AppConfiguration.SaveConfig("Deposit_Folder", textBox4.Text);
			AppConfiguration.SaveConfig("Ordered_folder", textBox5.Text);
			AppConfiguration.SaveConfig("Result_folder", textBox6.Text);



			DateTime dt = new DateTime(int.Parse(DateTime.Now.ToString("yyyy")), int.Parse(DateTime.Now.ToString("MM")), int.Parse(DateTime.Now.ToString("dd")),
										 int.Parse(DateTime.Now.ToString("HH")), int.Parse(DateTime.Now.ToString("mm")), int.Parse(DateTime.Now.ToString("ss"))
										 );


			AppConfiguration.toDay_RunTime = AppConfiguration.LoadConfig(AppConfiguration.GetToDay(dt)).Split(',');


			this.Close();

		}

		private void button2_Click(object sender, EventArgs e)
		{



		}
	}
}
