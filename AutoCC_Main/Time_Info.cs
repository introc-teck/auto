using System;
using System.Windows.Forms;

namespace AutoCC_Main
{
	public partial class Time_Info : UserControl
	{


		public Time_Info()
		{
			InitializeComponent();
		}


		public void set_name(string _val)
		{
			label1.Text = _val;
		}


		public string get_Enable()
		{
			if (checkBoxActivate.Checked)
			{
				return "T";
			}
			else
			{
				return "F";

			}
		}

		public string get_numericUpDownStartHour()
		{
			return numericUpDownStartHour.Value.ToString();

		}

		public string get_numericUpDownStartMinute()
		{
			return numericUpDownStartMinute.Value.ToString();

		}

		public string get_numericUpDownEndHour()
		{
			return numericUpDownEndHour.Value.ToString();

		}

		public string get_numericUpDownEndMinute()
		{
			return numericUpDownEndMinute.Value.ToString();

		}


		public void set_date(string _val)
		{
			if (!_val.Equals("null"))
			{
				string[] n1 = _val.Split(',');
				if (n1[0].Equals("T"))
				{

					set_enable();
				}
				else
				{
					set_disable();
				}
				numericUpDownStartHour.Value = decimal.Parse(n1[1]);
				numericUpDownStartMinute.Value = decimal.Parse(n1[2]);
				numericUpDownEndHour.Value = decimal.Parse(n1[3]);
				numericUpDownEndMinute.Value = decimal.Parse(n1[4]);
			}
			else
			{
				set_enable();
				numericUpDownStartHour.Value = 9;
				numericUpDownStartMinute.Value = 0;
				numericUpDownEndHour.Value = 20;
				numericUpDownEndMinute.Value = 10;
			}



		}



		private void checkBoxActivate_CheckedChanged(object sender, EventArgs e)
		{
			if (checkBoxActivate.Checked)
			{

				set_enable();
			}
			else
			{
				set_disable();
			}
		}


		private void set_enable()
		{

			checkBoxActivate.Checked = true;

			label1.Enabled = true;
			label2.Enabled = true;
			numericUpDownStartMinute.Enabled = true;
			numericUpDownStartHour.Enabled = true;
			numericUpDownEndMinute.Enabled = true;
			numericUpDownEndHour.Enabled = true;
		}

		private void set_disable()
		{
			checkBoxActivate.Checked = false;

			label1.Enabled = false;
			label2.Enabled = false;
			numericUpDownStartMinute.Enabled = false;
			numericUpDownStartHour.Enabled = false;
			numericUpDownEndMinute.Enabled = false;
			numericUpDownEndHour.Enabled = false;
		}
	}
}
