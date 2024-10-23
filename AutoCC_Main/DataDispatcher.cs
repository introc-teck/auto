using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Text;

namespace AutoCC_Main
{
    class DataDispatcher
    {
        public const int STATUS_INSPECT_START = 1;
        public const int STATUS_INSPECT_FINISH = 2;

        public const int RESULT_NONE = 0;
        public const int RESULT_SUCCEEDED = 1;
        public const int RESULT_FAILED = -1;

        public DataDispatcher()
        {
        }

        /*
         * 
         * {
  "result": {
    "code": "200",
    "message": "ok"
  },
  "data": {
    "0": "1185321",
    "1": "3966",
    "2": "003001001",
    "3": "일반명함",
    "4": "1",
    "5": "90",
    "6": "50",
    "7": "92",
    "8": "52",
    "9": "W",
    "10": "메뉴명함 경기_수원화서시장점.pdf",
    "11": "/attach/gp/order_file/2024/08/21/1183024.pdf",
    "12": "8",
    "13": "1199",
    "14": "1000",
    "15": "90*50",
    "16": "◎타라그래픽스",
    "17": "A 9",
    "18": "GPT240821NC00108",
    "order_common_seqno": "1185321",
    "member_seqno": "3966",
    "ProductCode": "003001001",
    "ProductName": "일반명함",
    "CaseCode": "1",
    "Width": "90",
    "Height": "50",
    "PWidth": "92",
    "PHeight": "52",
    "oper_sys": "W",
    "FileName": "메뉴명함 경기_수원화서시장점.pdf",
    "FilePath": "/attach/gp/order_file/2024/08/21/1183024.pdf",
    "Tmpt": "8",
    "PaperCode": "1199",
    "QuantityValue": "1000",
    "SizeCode": "90*50",
    "UCom": "◎타라그래픽스",
    "Method": "A 9",
    "order_num": "GPT240821NC00108"
  }
}
         */

        public static int GetWorkList4(Machine machine, ref List<Work> work_list)
        {
            try
            {
                string[] param = new string[1];
                param[0] = "id=" + "auto" + (AppConfiguration.machine.index).ToString();
                string result1 = request("GetAutoWaitingOrder.php", param);

                Work work = new Work();

                JObject obj = JObject.Parse(result1);

                work.order_date = obj["data"]["order_num"].ToString().Substring(3, 6);
                work.OPISeq = obj["data"]["order_num"].ToString().Substring(11, 5);
                work.order_common_seqno = Convert.ToInt32(obj["data"]["order_common_seqno"]);
                work.order_num = obj["data"]["order_num"].ToString();
                work.OPICustomerID = obj["data"]["member_seqno"].ToString();
                work.category_code = obj["data"]["ProductCode"].ToString();
                work.category_name = obj["data"]["ProductName"].ToString();
                work.item_count = Convert.ToInt32(obj["data"]["CaseCode"].ToString());

                double bleed_width = Convert.ToDouble(obj["data"]["PWidth"].ToString());
                double bleed_height = Convert.ToDouble(obj["data"]["PHeight"].ToString());
                work.bleed_size = new System.Drawing.SizeF((float)bleed_width, (float)bleed_height);

                double trim_width = Convert.ToDouble(obj["data"]["Width"].ToString());
                double trim_height = Convert.ToDouble(obj["data"]["Height"].ToString());
                work.trim_size = new System.Drawing.SizeF((float)trim_width, (float)trim_height);

                work.Platform = obj["data"]["oper_sys"].ToString();
                work.file_path = obj["data"]["FilePath"].ToString();

                work.DownloadFile();

                work.ColorCode = obj["data"]["Tmpt"].ToString();
                if (work.ColorCode.Equals("4") == true)
                    work.side_count = 1;
                else
                    work.side_count = 2;

                work.PaperCode = obj["data"]["PaperCode"].ToString();
                work.PaperQuantity = Convert.ToDouble(obj["data"]["QuantityValue"].ToString());
                work.Standard = obj["data"]["SizeCode"].ToString() == "비규격" ? "1" : "0";

                // 아웃라인 체크
                /*
                if (work.category_code == "003002001" || work.category_code == "003003001" || work.Standard == "1")
                    work.magenta_outline = true;
                else
                    work.magenta_outline = false;

                if (work.category_code.StartsWith("004"))
                {
                    if ((bleed_width == trim_width) && (bleed_height == trim_height))
                        work.magenta_outline = false;
                    else
                        work.magenta_outline = true;
                }
                */
                work.magenta_outline = false;

                // 당일판
                work.PeriodAM = "N";
                work.ImpositionPosition = "1";
                work.CustomerName = obj["data"]["UCom"].ToString();
                work.DeliveryCode = obj["data"]["Method"].ToString();
                work.ExtraTitle = "없음";

                if (work_list == null)
                    work_list = new List<Work>();

                work_list.Add(work);

                return (work_list == null ? 0 : work_list.Count);
            }
            catch
            {
                return 0;
            }
        }

        public static void SetSuccess(int order_common_seqno)
        {
            string[] param = new string[2];
            param[0] = "result=success";
            param[1] = "order_common_seqno=" + order_common_seqno;
            string result1 = request("UpdateAutoResult.php", param);
        }

        public static bool RemoveAccept(int order_common_seqno, string error_list)
        {
            string[] param = new string[3];
            param[0] = "result=fail";
            param[1] = "order_common_seqno=" + order_common_seqno;
            param[2] = "error_list=" + error_list;
            string result1 = request("UpdateAutoResult.php", param);

            return true;
        }

        public static bool RequestQCCheck(string opi_date, string opi_seq, string message)
        {
            /*
            string query1 = @"SELECT QCCheck FROM pdfauto..TbDPAutoErrorList WHERE OPIDate = '" + opi_date + "' AND OPISeq = '" + opi_seq + "'";

            //string log1 = string.Format("[DB] RequestQCCheck. query_string: {0}", query1);
            //	Logger.Log(log1);

            bool result = false;
            DataTable data_table = QueryExecute(query1);
            if (data_table != null)
            {
                if (data_table.Rows.Count > 0)
                {
                    string query2 = @"UPDATE pdfauto..TbDPAutoErrorList SET QCCheck = '" + message + "' " +
                                                        "WHERE OPIDate = '" + opi_date + "' AND OPISeq = '" + opi_seq + "'";

                    string log2 = string.Format("[DB] RequestQCCheck. query_string: {0}", query2);
                    //		Logger.Log(log2);

                    result = UpdateExecute(query2);
                }
                else
                {
                    string query2 = @"INSERT INTO pdfauto..TbDPAutoErrorList(OPIDate, OPISeq, Status, Contents, QCCheck)
														VALUES('" + opi_date + "', '" + opi_seq + "', 'S', '', '" + message + "')";

                    string log2 = string.Format("[DB] RequestQCCheck. query_string: {0}", query2);
                    //		Logger.Log(log2);

                    result = UpdateExecute(query2);
                }
            }

            return result;
            */

            return true;
        }

        public static string request(string path, string[] strs)
        {

            string strUri = "https://www.gprinting.co.kr/moamoa/process/" + path;
            // POST, GET 보낼 데이터 입력

            StringBuilder dataParams = new StringBuilder();
            if (strs != null)
            {
                int i = 0;
                foreach (string str in strs)
                {
                    if (dataParams.Length == 0)
                    {
                        dataParams.Append(str);
                    }
                    else
                    {
                        dataParams.Append("&" + str);
                    }
                    i++;
                }
            }

            // 요청 String -> 요청 Byte 변환
            byte[] byteDataParams = UTF8Encoding.UTF8.GetBytes(dataParams.ToString());

            /* GET */
            // GET 방식은 Uri 뒤에 보낼 데이터를 입력하시면 됩니다.

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strUri + "?" + dataParams);
            request.Method = "GET";

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream stReadData = response.GetResponseStream();
            StreamReader srReadData = new StreamReader(stReadData, Encoding.UTF8);

            string strResult = srReadData.ReadToEnd();
            return strResult;
        }
    }
}
