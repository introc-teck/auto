
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using log4net.Config;
using log4net;

namespace AutoCC_Sub
{
    public partial class Form1 : Form
    {
        #region Process 통신

        const int WM_COPYDATA = 0x4A;
        Dictionary<string, string> m_work_infos = null;
        Work m_work = null;
        private ILog logger = null;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, uint wParam, ref COPYDATASTRUCT lParam);
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint Msg, UIntPtr wParam, ref COPYDATASTRUCT lParam, uint fuFlags, uint uTimeout, out UIntPtr lpdwResult);

        public struct COPYDATASTRUCT
        {
            public IntPtr dwData;
            public int cbData;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpData;
        }

        private string UpdateStatus(int step, string status)
        {
            Process[] process = Process.GetProcessesByName("AutoCC_Main");
            if (process.Length > 0)
            {
                string content = string.Format("todo=update_status;step={0};status={1}", step, status);
                byte[] buff = System.Text.Encoding.Default.GetBytes(content);

                COPYDATASTRUCT cds = new COPYDATASTRUCT();
                cds.dwData = IntPtr.Zero;
                cds.cbData = buff.Length + 1;
                cds.lpData = content;

                UIntPtr result = UIntPtr.Zero;

                SendMessageTimeout(process[0].MainWindowHandle, WM_COPYDATA, (UIntPtr)m_work.order_common_seqno, ref cds, 0x0002, 5000, out result);
            }

            return "1";
        }

        private void NotifyResult(string content)
        {
            Process[] process = Process.GetProcessesByName("AutoCC_Main");

            if (process.Length > 0)
            {
                byte[] buff = System.Text.Encoding.Default.GetBytes(content);

                COPYDATASTRUCT cds = new COPYDATASTRUCT();
                cds.dwData = IntPtr.Zero;
                cds.cbData = buff.Length + 1;
                cds.lpData = content;

                UIntPtr result = UIntPtr.Zero;

                SendMessageTimeout(process[0].MainWindowHandle, WM_COPYDATA, (UIntPtr)m_work.order_common_seqno, ref cds, 0x0002, 5000, out result);
            }
        }
        #endregion

        public Form1()
        {
            InitializeComponent();

            String logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log4NetSub.xml");
            FileInfo file = new FileInfo(logPath);
            XmlConfigurator.Configure(file);
            logger = LogManager.GetLogger(this.GetType());
        }

        private bool Search_Illustrator()
        {
            try
            {
                foreach (Process process in Process.GetProcesses())
                {
                    if (string.Equals(process.ProcessName.ToLower().Trim(), "Illustrator".ToLower().Trim()))
                        return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }


        private bool Search_CorelDRAW()
        {
            try
            {
                foreach (Process process in Process.GetProcesses())
                {
                    if (string.Equals(process.ProcessName.ToLower().Trim(), "CorelDRW".ToLower().Trim()))
                        return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }


        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(10, 10);
            this.Show();

            if (!Search_Illustrator())
            {
                Process.Start("Illustrator.exe");
                System.Threading.Thread.Sleep(12000);
            }

            if (!Search_CorelDRAW())
            {
                Process.Start("CorelDRW.exe");
                System.Threading.Thread.Sleep(12000);
            }

            ParseArguments(Environment.CommandLine, ref m_work);

            logger.Debug("--->" + Environment.CommandLine);

            List<string> error_codes = null;
            bool result = false;
            string drawer_code = string.Empty;

            System.Text.StringBuilder error_string = new System.Text.StringBuilder();

            if (m_work != null)
            {
                result = Inspect(ref m_work, ref error_codes);

                drawer_code = m_work.drawer_code;

                if (error_codes != null)
                {
                    foreach (var error_code in error_codes)
                    {
                        if (error_string.Length != 0)
                            error_string.Append("|");
                        error_string.Append(error_code);
                    }
                }
            }
            else
            {
                result = false;
                error_string.Append(Error.INVALID_WORK_DATA);
            }

            UpdateStatus(3, " 작업완료");

            NotifyResult(string.Format("todo=report;result={0};drawer_code={1};error_codes={2}", result == true ? 1 : 0, drawer_code == string.Empty ? "" : drawer_code, error_string));

            m_work = null;

            System.GC.Collect();

            Console.WriteLine("결과:" + result);
            Console.WriteLine("결과:" + error_string.ToString());
        }

        private bool Inspect(ref Work work, ref List<string> error_codes)
        {
            UpdateStatus(0, "주문 내용 확인:" + DateTime.Now.ToString("HH:mm:ss"));

            string work_file_name = System.IO.Path.GetFileName(work.file_path);

            bool fixable = false;

            String cate1 = work.category_code.Replace(" ", "").Substring(0, 3);
            switch (cate1)
            {
                case "003": // 명함
                    fixable = true;
                    break;
                case "004": // 스티커
                    fixable = true;
                    break;
                case "005": // 전단						 
                    fixable = true;
                    break;
                default:
                    Error.AddErrorCode(Error.CATEGORY_NOT_ALLOWED, false, ref error_codes);
                    return false;
            }

            char[] delimiters = { '.' };
            string[] tokens = work_file_name.Split(delimiters);

            if (tokens.Length < 2)
            {
                Error.AddErrorCode(Error.FILE_EXT_NOT_EXIST, false, ref error_codes);
                return false;
            }

            DrawerHandler drawer_handler = null;

            string ext = tokens[tokens.Length - 1].ToLower();

            #region 파일처리 클레스 할당 시작

            switch (ext)
            {
                case "cdr":
                    drawer_handler = new CorelDRAWHandler();
                    break;
                case "ai":
                    //2019-01-17 추가
                    drawer_handler = new IllustratorHandler();
                    break;
                case "ps":
                case "eps":
                    drawer_handler = new IllustratorHandler();
                    break;
                case "jpg":
                case "jpe":
                case "jpeg":
                    drawer_handler = new IllustratorHandlerForJPG(); 
                    break;
                case "pdf":
                    drawer_handler = new IllustratorHandlerForPDF();
                    ((IllustratorHandlerForPDF)drawer_handler).ItemCount = work.item_count;
                    break;
                case "zip":
                    if (ExtractWorkFile(work) == true)
                    {
                        string date_folder_path = string.Format("{0}{1}\\", m_work_infos["works_folder_path"], work.order_date);
                        string extracted_folder_path = string.Format("{0}{1}\\extracted_files\\", date_folder_path, work.order_num);
                        string extension = GetExtensionOfExtractedFiles(work);

                        switch (extension.ToLower())
                        {
                            case ".cdr":
                                drawer_handler = new CorelDRAWHandlerForZIP();
                                ((CorelDRAWHandlerForZIP)drawer_handler).ItemCount = work.item_count;
                                ((CorelDRAWHandlerForZIP)drawer_handler).FolderPath = extracted_folder_path;
                                break;
                            case ".ai":
                                drawer_handler = new IllustratorHandlerForZIP();
                                ((IllustratorHandlerForZIP)drawer_handler).ItemCount = work.item_count;
                                ((IllustratorHandlerForZIP)drawer_handler).FolderPath = extracted_folder_path;
                                break;
                            case ".eps":
                                drawer_handler = new IllustratorHandlerForZIP();
                                ((IllustratorHandlerForZIP)drawer_handler).ItemCount = work.item_count;
                                ((IllustratorHandlerForZIP)drawer_handler).FolderPath = extracted_folder_path;
                                break;
                            case ".pdf":
                                drawer_handler = new IllustratorHandlerForZIP_PDF();
                                ((IllustratorHandlerForZIP_PDF)drawer_handler).ItemCount = work.item_count;
                                ((IllustratorHandlerForZIP_PDF)drawer_handler).FolderPath = extracted_folder_path;
                                break;
                            case ".jpg":
                            /**
                             * jpg 계열은 모두 동일
                             */
                            // if (work.side_count == 1)
                            // {
                            //     drawer_handler = new IllustratorHandlerForZIP_JPG(work.side_count);
                            //     ((IllustratorHandlerForZIP_JPG)drawer_handler).ItemCount = work.item_count;
                            //     ((IllustratorHandlerForZIP_JPG)drawer_handler).FolderPath = extracted_folder_path;
                            // }
                            // break;
                            case ".jpe":
                            case ".jpeg": //2020.05.06 추가
                                if (work.side_count == 1)
                                {
                                    drawer_handler = new IllustratorHandlerForZIP_JPG(work.side_count);
                                    ((IllustratorHandlerForZIP_JPG)drawer_handler).ItemCount = work.item_count;
                                    ((IllustratorHandlerForZIP_JPG)drawer_handler).FolderPath = extracted_folder_path;
                                }
                                break;
                        }
                    }
                    break;
                default:
                    break;
            }

            #endregion 파일처리 클레스 할당 종료

            if (drawer_handler == null)
            {
                Error.AddErrorCode(Error.FILE_EXT_NOT_ALLOWED, false, ref error_codes);
                return false;
            }

            drawer_handler.Fixable = fixable;
            work.drawer_code = string.Format("{0:00}", drawer_handler.m_drawer_index);

            #region 파일존재 유뮤 확인 시작
            try
            {
                System.IO.FileInfo file_info = new System.IO.FileInfo(work.file_path);

                if (file_info.Exists == false)
                {


                    Error.AddErrorCode(Error.FILE_NOT_FOUND, false, ref error_codes);
                    return false;
                }
            }
            catch
            {
                //   UpdateStatus(0, " 에러-> 파일처리 RunTime 오류");
                Error.AddErrorCode(Error.FILE_NOT_FOUND, false, ref error_codes);
                return false;
            }
            #endregion  파일존재 유뮤 확인 종료


            // 1. open     
            if (drawer_handler.OpenDocument(work.file_path, ref error_codes) == false)
            {
                drawer_handler.CloseDocument();
                return false;
            }

            UpdateStatus(0, "파일 OPEN:" + DateTime.Now.ToString("HH:mm:ss"));

            // 2. check all
            bool result = CheckAll(drawer_handler, work, ref error_codes);

            // 3. close		
            drawer_handler.CloseDocument();

            System.GC.Collect();

            return result;
        }


        private bool CheckAll(DrawerHandler drawer_handler, Work work, ref List<string> error_codes)
        {
            // 1. prepare          
            if (drawer_handler.Prepare(ref error_codes) == false)
            {
                return false;
            }
            UpdateStatus(0, " GideLine,empt page 삭제");


            if (drawer_handler.CheckObjectCount(ref error_codes) == false)
            {
                return false;
            }


            // 2. check pages       
            if (work.file_path.Contains("jpg") || work.file_path.Contains("jpeg"))
            {

                // 2020.06.01 JPG파일 압뒤 구분이 안되서 양면이면 무조건 수동 처리
                // 단 여러건의 단면이면 처리

                if (work.side_count != 1)
                {
                    UpdateStatus(1, " 에러 -> 양면처리 불가");
                    return false;
                }
            }


            // 일러스트에 링크로 된 이미지 같은게 있을 경우 이단계에서 먼처 처리
            // 기존에는 이것이 페이지 분리 다음에 있어서 대부분의 작업이 완료된 다음 진행 되는 버그로
            // 시간낭비가 많음

            UpdateStatus(1, "벨리데이션:" + DateTime.Now.ToString("HH:mm:ss"));

            bool validation_result = drawer_handler.CheckValidations(ref error_codes, UpdateStatus);
            if (validation_result == false)
            {
                return false;
            }

            UpdateStatus(2, "페이지:" + DateTime.Now.ToString("HH:mm:ss"));

            bool pages_result = drawer_handler.CheckPages(work.category_code, work.bleed_size, work.trim_size, work.side_count, work.item_count, ref error_codes, UpdateStatus);

            if (pages_result == false)
            {
                return false;
            }

            // works folder
            if (System.IO.Directory.Exists(m_work_infos["works_folder_path"]) == false)
                System.IO.Directory.CreateDirectory(m_work_infos["works_folder_path"]);

            string work_folder_path = m_work_infos["works_folder_path"] + @"\";
            if (System.IO.Directory.Exists(work_folder_path) == false)
                System.IO.Directory.CreateDirectory(work_folder_path);

            // 4. export pdf
            bool pdf_result = false;

            UpdateStatus(3, "PDF:" + DateTime.Now.ToString("HH:mm:ss"));

            if (work.category_code == "004003009")
            {
                string sub_type_code_prefix = (work.SubTypeCode.Length < 2 ? work.SubTypeCode : work.SubTypeCode.Substring(0, 2));
                switch (sub_type_code_prefix)
                {
                    case "SB": // 원형
                    case "SA": // 정사각
                    case "SC": // 반원
                    case "SS": // 하트
                    case "SR": // 직사각
                    case "SO": // 타원
                        break;
                    default:
                        pdf_result = drawer_handler.ExportPDF(work_folder_path, work.bleed_size, work.side_count, work.magenta_outline, ref error_codes, UpdateStatus, work.item_count);

                        break;
                }
            }
            else
            {
                pdf_result = drawer_handler.ExportPDF(work_folder_path, work.bleed_size, work.side_count, work.magenta_outline, ref error_codes, UpdateStatus, work.item_count);
            }
            if (pdf_result == false)
            {
                Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                return false;
            }

            return true;
        }


        private bool ExtractWorkFile(Work work)
        {
            // works folder
            if (System.IO.Directory.Exists(m_work_infos["works_folder_path"]) == false)
                System.IO.Directory.CreateDirectory(m_work_infos["works_folder_path"]);

            // works folder + order_date
            string date_folder_path = string.Format("{0}{1}\\", m_work_infos["works_folder_path"], work.order_date);
            if (System.IO.Directory.Exists(date_folder_path) == false)
                System.IO.Directory.CreateDirectory(date_folder_path);

            // works folder + order_date + order_id
            string work_folder_path = string.Format("{0}{1}\\", date_folder_path, work.order_num);
            if (System.IO.Directory.Exists(work_folder_path) == false)
                System.IO.Directory.CreateDirectory(work_folder_path);

            // works folder + order_date + order_id + extracted_files
            string extracted_folder_path = string.Format("{0}{1}\\extracted_files\\", date_folder_path, work.order_num);
            if (System.IO.Directory.Exists(extracted_folder_path) == false)
                System.IO.Directory.CreateDirectory(extracted_folder_path);

            try
            {
                System.IO.Compression.ZipFile.ExtractToDirectory(work.file_path, extracted_folder_path);

                System.IO.DirectoryInfo directory_info = new System.IO.DirectoryInfo(extracted_folder_path);
                System.IO.FileInfo[] sub_file_infos = directory_info.GetFiles();
                System.IO.DirectoryInfo[] sub_directory_infos = directory_info.GetDirectories();
                if (sub_file_infos.Length == 0 && sub_directory_infos.Length == 1)
                {
                    System.IO.FileInfo[] file_infos_in_sub_directory = sub_directory_infos[0].GetFiles();
                    if (file_infos_in_sub_directory.Length > 0)
                    {
                        foreach (var file_info in file_infos_in_sub_directory)
                        {
                            string new_file_path = string.Format("{0}{1}", extracted_folder_path, file_info.Name);
                            file_info.MoveTo(new_file_path);
                        }
                    }
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        private string GetExtensionOfExtractedFiles(Work work)
        {
            string date_folder_path = string.Format("{0}{1}\\", m_work_infos["works_folder_path"], work.order_date);
            string extracted_folder_path = string.Format("{0}{1}\\extracted_files\\", date_folder_path, work.order_num);

            Dictionary<string, int> extension_list = new Dictionary<string, int>();
            string[] file_paths_in_directory = System.IO.Directory.GetFiles(extracted_folder_path);
            foreach (var file_path in file_paths_in_directory)
            {
                System.IO.FileInfo file_info = new System.IO.FileInfo(file_path);

                if (extension_list.ContainsKey(file_info.Extension.ToLower()) == true)
                    extension_list[file_info.Extension.ToLower()]++;
                else
                    extension_list.Add(file_info.Extension.ToLower(), 1);
            }

            string extension = string.Empty;
            if (extension_list.Count == 1)
            {
                extension = extension_list.ElementAt(0).Key;
            }
            else if (extension_list.Count == 2)
            {
                if (extension_list.ElementAt(0).Key == ".jpg")
                {
                    switch (extension_list.ElementAt(1).Key)
                    {
                        case ".ai":
                        case ".eps":
                        case ".cdr":
                        case ".pdf":
                            extension = extension_list.ElementAt(1).Key;
                            break;
                    }
                }
                else if (extension_list.ElementAt(1).Key == ".jpg")
                {
                    switch (extension_list.ElementAt(0).Key)
                    {
                        case ".ai":
                        case ".eps":
                        case ".cdr":
                        case ".pdf":
                            extension = extension_list.ElementAt(0).Key;
                            break;
                    }
                }
            }

            return extension;
        }


        private void ParseArguments(string arguments, ref Work work)
        {
            m_work_infos = new Dictionary<string, string>();

            // Console.WriteLine(arguments);
            logger.Debug(arguments);

            //error
            // arguments = @"order_id=0049;order_seq=49;order_date=20200508;category_code=N11;file_path=x:\\IBM\20200508\DP-20200508-0049-윈윈파트너.pdf;bleed_size_width=88.00;bleed_size_height=54.00;trim_size_width=86.00;trim_size_height=52.00;side_count=2;item_count=4;marker=;magenta_outline=False;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";

            // 코렐 주문건 1개의 아이템이 위 아래로 배치된 경우 처리 되도록 수정 test 파일
            // arguments = @"order_id=2735;order_seq=2735;order_date=20200513;category_code=N11;file_path=x:\\IBM\20200513\DP-20200513-2735-본래순대-.cdr;bleed_size_width=88.00;bleed_size_height=54.00;trim_size_width=86.00;trim_size_height=52.00;side_count=2;item_count=1;marker=;magenta_outline=true;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";

            //일러스트트 스티커 가로세로가 바뀐경우 sample 에러남
            //    arguments = @"order_id=2499;order_seq=2499;order_date=20200513;category_code=N13;file_path=x:\\IBM\20200513\DP-20200513-2499-신참쿠폰a.ai;bleed_size_width=92.00;bleed_size_height=52.00;trim_size_width=90.00;trim_size_height=50.00;side_count=2;item_count=1;marker=;magenta_outline=True;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";

            // 코렐 아이템이 위 아래 배치 Test
            // arguments = @" order_id=2407;order_seq=2407;order_date=20200513;category_code=N13;file_path=x:\\IBM\20200513\DP-20200513-2407-2020_.cdr;bleed_size_width=92.00;bleed_size_height=52.00;trim_size_width=90.00;trim_size_height=50.00;side_count=2;item_count=1;marker=;magenta_outline=True;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";

            //일러스트 아이템이 위 아래 배치 Test
            //  arguments = @"order_id=2211;order_seq=2211;order_date=20200513;category_code=N11;file_path=x:\\IBM\20200513\DP-20200513-2211-우진건설-error2.ai;bleed_size_width=88.00;bleed_size_height=54.00;trim_size_width=86.00;trim_size_height=52.00;side_count=2;item_count=1;marker=;magenta_outline=False;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";

            //일러스트트 전단 가로세로가 바뀐경우 sample 에러남
            //  arguments = @"order_id=1723;order_seq=1723;order_date=20200513;category_code=B01;file_path=x:\\IBM\20200513\DP-20200513-1723-0407목.ai;bleed_size_width=213.00;bleed_size_height=150.00;trim_size_width=210.00;trim_size_height=147.00;side_count=2;item_count=1;marker=;magenta_outline=False;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";

            //   arguments = @"order_id=1395;order_seq=1395;order_date=20200513;category_code=S10;file_path=x:\\IBM\20200513\DP-20200513-1395-디자인입춘.ai;bleed_size_width=96.00;bleed_size_height=116.00;trim_size_width=90.00;trim_size_height=110.00;side_count=1;item_count=1;marker=;magenta_outline=True;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";
            //일러스트트 전단 가로세로가 바뀐경우 sample 에러남
            //arguments = @" order_id=1199;order_seq=1199;order_date=20200513;category_code=S20;file_path=x:\\IBM\20200513\DP-20200513-1199-20 5.eps;bleed_size_width=90.00;bleed_size_height=60.00;trim_size_width=90.00;trim_size_height=60.00;side_count=1;item_count=2;marker=;magenta_outline=False;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";

            // 오브젝트 05 로시작 해서 실폐 하는 예제 -> 이부분 확인 부분을 PASS 하면 성공적으로 PDF 생성
            //  arguments = @"order_id=0863;order_seq=863;order_date=20200513;category_code=N20;file_path=x:\\MAC\20200513\2005130863-20051.ai;bleed_size_width=92.00;bleed_size_height=52.00;trim_size_width=90.00;trim_size_height=50.00;side_count=2;item_count=4;marker=;magenta_outline=False;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";

            // 코렐 아웃라인 나오는거 sample
            // arguments = @"order_id=0106;order_seq=106;order_date=20200622;category_code=N13;file_path=c:\DP-20200622-0106-싱싱강남어.cdr;bleed_size_width=92.00;bleed_size_height=52.00;trim_size_width=90.00;trim_size_height=50.00;side_count=2;item_count=1;marker=;magenta_outline=True;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";

            // object 12000개
            // arguments = @"order_id=1779;order_seq=1779;order_date=20200706;category_code=N13;file_path=x:\\IBM\20200706\DP-20200706-1779-윤성학-명.cdr;bleed_size_width=92.00;bleed_size_height=52.00;trim_size_width=90.00;trim_size_height=50.00;side_count=2;item_count=1;marker=;magenta_outline=True;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";
            // arguments = @"order_id=0331;order_seq=331;order_date=20200721;category_code=S20;file_path=x:\\IBM\20200721\DP-20200721-0331-워터맨하우.ai;bleed_size_width=55.00;bleed_size_height=100.00;trim_size_width=55.00;trim_size_height=100.00;side_count=1;item_count=2;marker=;magenta_outline=False;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";
            // arguments = @"order_id=1934;order_seq=1934;order_date=20200724;category_code=N20;file_path=c:\\DP-20200724-1934-07-24test.cdr;bleed_size_width=92.00;bleed_size_height=52.00;trim_size_width=90.00;trim_size_height=50.00;side_count=1;item_count=2;marker=;magenta_outline=False;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";
            // arguments = @"order_id=0394;order_seq=394;order_date=20200824;category_code=N20;file_path=x:\\IBM\20200824\DP-20200824-0394-베노 김대.cdr;bleed_size_width=92.00;bleed_size_height=52.00;trim_size_width=90.00;trim_size_height=50.00;side_count=1;item_count=1;marker=;magenta_outline=False;works_folder_path=C:\\Auto_CC\\result_files\\;log_file_path=C:\\Auto_CC\\log\";

            try
            {
                // arguments 의 마지막이 파일경로인 '\' 으로 끝나서 이것을 제외한 나머지 압쪽데이타 추출용안데
                // 그냥 다음단계에서 해도 문제는 없을듯..

                char[] separator1 = { '\"' };

                string[] file_path_and_arguments = arguments.Split(separator1);

                //logger.Debug("file_path_and_arguments:" + file_path_and_arguments.Length);

                if (file_path_and_arguments.Length > 0)
                {
                    string arguments_part = file_path_and_arguments[file_path_and_arguments.Length - 1];
                    arguments_part.Trim();

                    char[] separator2 = { ';' };
                    string[] pairs = arguments_part.Split(separator2);

                    foreach (var item in pairs)
                    {

                        int index = item.IndexOf('=');
                        if (index != -1)
                        {
                            string key = item.Substring(0, index);
                            m_work_infos[key.Trim()] = item.Substring(index + 1, item.Length - (index + 1));
                        }
                    }
                }

                if (work == null)
                {
                    work = new Work();
                }

                //  logger.Debug("work create");

                work.order_num = m_work_infos["order_id"];
                work.order_common_seqno = Convert.ToInt32(m_work_infos["order_seq"]);
                work.order_date = m_work_infos["order_date"];
                work.category_code = m_work_infos["category_code"];
                work.file_path = m_work_infos["file_path"];

                double bleed_width = Convert.ToDouble(m_work_infos["bleed_size_width"]);
                double bleed_height = Convert.ToDouble(m_work_infos["bleed_size_height"]);
                work.bleed_size = new SizeF((float)bleed_width, (float)bleed_height);

                double trim_width = Convert.ToDouble(m_work_infos["trim_size_width"]);
                double trim_height = Convert.ToDouble(m_work_infos["trim_size_height"]);
                work.trim_size = new SizeF((float)trim_width, (float)trim_height);

                work.side_count = Convert.ToInt32(m_work_infos["side_count"]);
                work.item_count = Convert.ToInt32(m_work_infos["item_count"]);

                work.marker = m_work_infos["marker"];

                // work.magenta_outline = Convert.ToBoolean(m_work_infos["magenta_outline"]);
                work.magenta_outline = false;

                UpdateStatus(0, " 2. 접수내역을 Work Class 저장");
            }
            catch
            {

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            UpdateStatus(int.Parse(textBox1.Text), textBox2.Text);
        }

        private void button2_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < 100; i++)
            {
                for (int j = 1; j < 5; j++)
                {
                    UpdateStatus(1, j.ToString());
                    System.Threading.Thread.Sleep(1000);

                    // 이걸해주면 하우스 드레그를 하는 동안에는 해당 For 문이 중단 된고
                    // 응답엄음 메시지도 출력 되지 않는다
                    //  Application.DoEvents(); 가 없으면 해당 for문제 종료 될때까지 
                    // 프로그램은 응답없음 
                    // 마치 다운된것 처럼 보임
                    Application.DoEvents();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            m_work.order_common_seqno = 123;

            List<string> error_codes = null;
            bool result = false;
            string drawer_code = string.Empty;

            System.Text.StringBuilder error_string = new System.Text.StringBuilder();
            NotifyResult(string.Format("todo=report;result={0};drawer_code={1};error_codes={2}", result == true ? 1 : 0, drawer_code == string.Empty ? "" : drawer_code, error_string));
        }
    }
}
