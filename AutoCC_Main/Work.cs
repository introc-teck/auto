using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading;

namespace AutoCC_Main
{
	class Work
	{
		//		public string order_id { get; set; }
		public string order_date { get; set; }
		public int order_common_seqno { get; set; }
		public string order_num { get; set; }
		public string PreviewPath { get; set; }
		public string order_detail_dvs_num { get; set; }
		public string drawer_code { get; set; }     // 01: CorelDRAW, 02: Illustrator
		public string category_code { get; set; }
		public string category_name { get; set; }
		public string file_path { get; set; }       // file ordered from user
		public SizeF bleed_size { get; set; }
		public SizeF trim_size { get; set; }
		public int side_count { get; set; }
		public int item_count { get; set; }
		public string finish_codes { get; set; }    // finish id list. ex) aaa|bbb|ccc
		public string marker { get; set; }
		public bool magenta_outline { get; set; }


		public string OPISeq { get; set; }
		public string OPICustomerID { get; set; }       // customer id
		public string Platform { get; set; }
		public string CustomerName { get; set; }        // customer name
		public string DeliveryCode { get; set; }        // delivery code
		public string ExtraTitle { get; set; }          // extra title
		public string Standard { get; set; }            // 0: 규격, 1: 비규격
		public string ImpositionPosition { get; set; }  // 조판될 자릿수. 1-based
		public string PeriodAM { get; set; }            // Y: 오전판
		public string PaperCode { get; set; }           // 종이종류 code
		public double PaperQuantity { get; set; }       // 종이수량. ~매
		public string ColorCode { get; set; }           // 인쇄도수
		public string SizeCode { get; set; }           // 
		public string SubTypeCode { get; set; }
		public string Coating { get; set; }             // 코팅 유무. 0: 비코팅, 1: 코팅


		public Work()
		{
			//			order_id = string.Empty;
			order_date = string.Empty;
			order_common_seqno = 0;
			order_num = string.Empty;
			order_detail_dvs_num = string.Empty;
			drawer_code = string.Empty;
			category_code = string.Empty;
			category_name = string.Empty;
			file_path = string.Empty;
			bleed_size = SizeF.Empty;
			trim_size = SizeF.Empty;
			side_count = 0;
			item_count = 0;
			finish_codes = string.Empty;
			marker = string.Empty;
			magenta_outline = false;


			OPISeq = string.Empty;
			OPICustomerID = string.Empty;
			CustomerName = string.Empty;
			DeliveryCode = string.Empty;
			ExtraTitle = string.Empty;
			Standard = string.Empty;
			ImpositionPosition = string.Empty;
			PeriodAM = string.Empty;
			PaperCode = string.Empty;
			PaperQuantity = 0.0;
			ColorCode = string.Empty;
			SizeCode = string.Empty;
			SubTypeCode = string.Empty;

		}

		public String GetCategoryCode(int step) // 1, 2, 3
		{
			if (category_code.Length != 9)
				return string.Empty;

			return category_code.Substring((step - 1) * 3, 3);
		}


		public string GetCategoryName()
		{
			string category_name = string.Empty;

			switch (category_code)
			{
				case "N11": category_name = "코팅명함"; break;
				case "N12": category_name = "무코팅명함"; break;
				case "N13": category_name = "고품격코팅명함"; break;
				case "N14": category_name = "고품격무코팅명함"; break;
				case "N20": category_name = "수입명함"; break;
				case "N30": category_name = "옵셋카드명함"; break;
				case "S10": category_name = "컬러 일반/강접 스티커"; break;
				case "S20": category_name = "특수지스티커"; break;
				case "S30": category_name = "도무송스티커"; break;
				case "B01": category_name = "합판전단지"; break;
				case "B02": category_name = "독판전단지"; break;
				case "B03": category_name = "초소량인쇄"; break;
				case "E10": category_name = "컬러옵셋봉투"; break;
				case "E20": category_name = "마스타봉투"; break;
			}

			return category_name;
		}

		public string GetFileName(string order_no, string order_title)
		{
			string file_name = string.Empty;

			string temp = string.Empty;

			if (OPICustomerID == "106378" || OPICustomerID == "106853" || OPICustomerID == "106984" || OPICustomerID == "105222")
				temp = "-" + ExtraTitle;

			string temp_imposition_position = ImpositionPosition;
			string period_am_name = "";
			int multiple = 1;

			string category_prefix = category_code.Substring(0, 1);
			if (category_prefix == "N")
			{
				if (category_code == "N13" || category_code == "N14")
				{
					Standard = "1";

					if (temp_imposition_position == "1")
						temp_imposition_position = "2";
				}

				if (PeriodAM == "Y")
					period_am_name = "-오전";

				if (category_code == "N20")
					multiple = (int)PaperQuantity / 200;
				else
					multiple = 1;

				if (PaperCode == "TT")
				{
					if ((int)PaperQuantity != 200)
						multiple = (int)PaperQuantity / 200;
				}

				if (finish_codes != string.Empty)
				{
					if (finish_codes.Contains("0101") == true)
						temp += "-귀돌이(↙↖↗↘)";
					if (finish_codes.Contains("0801") == true)
						temp += "-반재단";
				}

				string delivery_name = GetDeliveryName();
				file_name = Convert.ToInt32(PaperQuantity).ToString() + "][" + CustomerName + temp + period_am_name + "][][" + order_no + "][" + multiple.ToString() + "][0][" + Convert.ToInt32(bleed_size.Width).ToString() + "x" + Convert.ToInt32(bleed_size.Height).ToString() + "][" + Standard + "][" + delivery_name + "][" + temp_imposition_position + "][N][.pdf";
			}
			else if (category_prefix == "S")
			{
				string customer_name = CustomerName;
				if (temp != string.Empty)
					customer_name = customer_name + temp;

				string tag = string.Empty;
				switch (PaperCode)
				{
					case "S2": tag = "-강접"; break;
					case "S3": tag = "-은데드롱"; break;
					case "S4": tag = "-유포지"; break;
					case "S5": tag = "-투명데드롱"; break;
					case "S6": tag = "-S모조"; break; // 2019.04.16 추가
					case "S7": tag = "-S크라프트"; break; // 2019.04.16 추가
				}

				string delivery_name = GetDeliveryName();

				string sub_type_code_prefix = (SubTypeCode.Length < 2 ? SubTypeCode : SubTypeCode.Substring(0, 2));
				if (sub_type_code_prefix == "SE")
				{
					if (PeriodAM == "Y")
						file_name = Convert.ToInt32(PaperQuantity).ToString() + "][" + customer_name + "-오전판][][" + order_no + "][" + (Convert.ToInt32(PaperQuantity) / 1000).ToString() + "][0][" + Convert.ToInt32(bleed_size.Width).ToString() + "x" + Convert.ToInt32(bleed_size.Height).ToString() + "][" + Standard + "][" + delivery_name + "]" + "[0]" + "[2][S][.pdf";
					else
						file_name = Convert.ToInt32(PaperQuantity).ToString() + "][" + customer_name + "][][" + order_no + "][" + (Convert.ToInt32(PaperQuantity) / 1000).ToString() + "][0][" + Convert.ToInt32(bleed_size.Width).ToString() + "x" + Convert.ToInt32(bleed_size.Height).ToString() + "][" + Standard + "][" + delivery_name + "]" + "[0]" + "[2][S][.pdf";
				}
				else
				{
					bool one_touch = ((bleed_size.Width == trim_size.Width) && (bleed_size.Height == trim_size.Height));

					if (PeriodAM == "Y")
						file_name = Convert.ToInt32(PaperQuantity).ToString() + "][" + customer_name + tag + "-오전판][][" + order_no + "][" + (Convert.ToInt32(PaperQuantity) / 1000).ToString() + "][0][" + Convert.ToInt32(bleed_size.Width).ToString() + "x" + Convert.ToInt32(bleed_size.Height).ToString() + "][" + Standard + "][" + delivery_name + "]" + (one_touch == true ? "[0]" : "[1]") + "[0][S][.pdf";
					else
						file_name = Convert.ToInt32(PaperQuantity).ToString() + "][" + customer_name + tag + "][][" + order_no + "][" + (Convert.ToInt32(PaperQuantity) / 1000).ToString() + "][0][" + Convert.ToInt32(bleed_size.Width).ToString() + "x" + Convert.ToInt32(bleed_size.Height).ToString() + "][" + Standard + "][" + delivery_name + "]" + (one_touch == true ? "[0]" : "[1]") + "[0][S][.pdf";
				}
			}
			else if (category_prefix == "B")
			{
				string customer_name = CustomerName;
				if (temp != string.Empty)
					customer_name = customer_name + temp;

				string size_name = string.Empty;
				switch (SizeCode)
				{
					case "A6": size_name = "A6"; break;
					case "A5": size_name = "A5"; break;
					case "A4": size_name = "A4"; break;
					case "A3": size_name = "A3"; break;
					case "B7": size_name = "64절"; break;
					case "B6": size_name = "32절"; break;
					case "B5": size_name = "16절"; break;
					case "B4": size_name = "8절"; break;
				}

				string delivery_name = GetDeliveryName();

				string plate_type = "지방판";
				if (delivery_name.Contains("서울") == true || delivery_name.Contains("방문") == true || delivery_name.Contains("강북2") == true || OPICustomerID.Contains("6567") == true)
					plate_type = "서울판";

				string paper_quantity = string.Empty;
				if ((PaperQuantity - Convert.ToInt32(PaperQuantity)) == 0)
					paper_quantity = Convert.ToInt32(PaperQuantity).ToString();
				else
					paper_quantity = PaperQuantity.ToString();
				paper_quantity = paper_quantity.Replace(".", "_");

				string date_info = DateTime.Now.ToString("yyMMdd");

				file_name = paper_quantity + "][" + customer_name + "-" + order_title + "][_][" + order_no + "][ART 100g][0][" + size_name + "][" + Standard + "][" + delivery_name + "][" + ColorCode + "도][" + plate_type + "][" + date_info + "][B][.pdf";
			}

			return file_name;
		}

		public bool DownloadFile()
		{
			string file_name = this.file_path.Split('/').Last();
			string download_path = "https://orderplatform.s3.ap-northeast-2.amazonaws.com";
			if (file_name.StartsWith("GP") || file_name.StartsWith("DP") || file_name.StartsWith("PP") || file_name.StartsWith("BP"))
			{
				download_path = PathHelper.ServerAddress + "/receipt";
			}

			int fileExtPos = this.file_path.LastIndexOf(".", StringComparison.Ordinal);
			string ext = this.file_path.Substring(fileExtPos);

			string outfile = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
								+ @"\receipt\" + this.order_num + @"\" + this.order_num + ext;

			if (!Directory.Exists(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
				   + @"\receipt\" + this.order_num))
			{
				Directory.CreateDirectory(System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
					+ @"\receipt\" + this.order_num);
			}

			int j = 1;

			if (File.Exists(outfile))
			{
				this.file_path = outfile;
			}
			else
			{
				while (!File.Exists(outfile))
				{
					try
					{
						using (WebClient client1 = new WebClient())
						{
							WebRequest serverRequest = WebRequest.Create(new Uri(download_path + this.file_path));
							WebResponse serverResponse;
							try //Try to get response from server  
							{
								serverResponse = serverRequest.GetResponse();
							}
							catch(Exception e11) //If could not obtain any response  
							{
								return false;
							}
							//GPT211201BL01850

							client1.DownloadFileAsync(new Uri(download_path + this.file_path), outfile);
							this.file_path = outfile;

							serverResponse.Close();
							client1.Dispose();
						}
						return true;
					}
					catch (Exception e)
					{
						Thread.Sleep(1000);
					}
				}
			}
			return true;
		}

		public string GetDeliveryName()
		{
			string delivery_name = string.Empty;

			switch (DeliveryCode)
			{
				case "001": delivery_name = "방문"; break;
				case "002": delivery_name = "경화"; break;
				case "003": delivery_name = "경택"; break;
				case "004": delivery_name = "직1"; break;
				case "005": delivery_name = "직2"; break;
				case "006": delivery_name = "직3"; break;
				case "007": delivery_name = "직4"; break;
				case "008": delivery_name = "직5"; break;
				case "009": delivery_name = "직6"; break;
				case "010": delivery_name = "직7"; break;
				case "011": delivery_name = "강북"; break;
				case "012": delivery_name = "경기"; break;
				case "013": delivery_name = "경인"; break;
				case "014": delivery_name = "한택"; break;
				case "015": delivery_name = "직8"; break;
				case "016": delivery_name = "직9"; break;
				case "017": delivery_name = "서울7"; break;
				case "018": delivery_name = "직11"; break;
				case "019": delivery_name = "직12"; break;
				case "020": delivery_name = "직13"; break;
				case "021": delivery_name = "직14"; break;
				case "022": delivery_name = "직15"; break;
				case "024": delivery_name = "퀵"; break;
				case "025": delivery_name = "서울"; break;
				case "026": delivery_name = "서울13"; break;
				case "027": delivery_name = "직16"; break;
				case "028": delivery_name = "서울2"; break;
				case "029": delivery_name = "직17"; break;
				case "030": delivery_name = "서울3"; break;
				case "031": delivery_name = "서울4"; break;
				case "032": delivery_name = "서울5"; break;
				case "033": delivery_name = "직0"; break;
				case "034": delivery_name = "직18"; break;
				case "035": delivery_name = "대신"; break;
				case "036": delivery_name = "서울6"; break;
				case "037": delivery_name = "방문필동"; break;
				case "039": delivery_name = "TNT"; break;
				case "023": delivery_name = "서울8"; break;
				case "040": delivery_name = "서울9"; break;
				case "041": delivery_name = "서울10"; break;
				case "042": delivery_name = "서울11"; break;
				case "046": delivery_name = "서울12"; break;
				case "099": delivery_name = "직배"; break;
				case "100": delivery_name = "성수"; break; // 2020.10.28 추가

			}

			return delivery_name;
		}

		

		public bool Verify()
		{
			if (order_date == string.Empty)
				return false;
			if (OPISeq == string.Empty)
				return false;
			if (OPICustomerID == string.Empty)
				return false;
			if (CustomerName == string.Empty)
				return false;
			if (DeliveryCode == string.Empty)
				return false;
			if (category_code == string.Empty)
				return false;
			if (file_path == string.Empty)
				return false;
			if (bleed_size == SizeF.Empty)
				return false;
			if (trim_size == SizeF.Empty)
				return false;
			if (side_count == 0)
				return false;
			if (item_count == 0)
				return false;

			return true;
		}

		public override string ToString()
		{
			string desc = string.Format("work: {0};{1};{2};{3};{4};{5};{6};{7};({8},{9});({10},{11});{12};{13};{14};{15}", order_date, OPISeq, OPICustomerID, CustomerName, DeliveryCode, category_code, category_name, side_count, bleed_size.Width, bleed_size.Height, trim_size.Width, trim_size.Height, file_path, item_count, Standard, magenta_outline);
			return desc;
		}
	}
}