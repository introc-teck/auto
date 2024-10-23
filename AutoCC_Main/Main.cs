
using ImageMagick;
using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace AutoCC_Main
{
    public partial class Main : Form
    {
        #region 프로세스 통신
        const int WM_COPYDATA = 0x4A;

        public struct COPYDATASTRUCT
        {
            public IntPtr dwData;
            public int cbData;
            [MarshalAs(UnmanagedType.LPStr)]
            public string lpData;
        }

        DateTime m_today = DateTime.Now;

        int _count = 0;
        int check_interval = 25;  // 20초에 한번씩 DB조회

        List<Work> m_works = null;

        string _stepval = "";

        int Current_Step = 0;

        protected override void OnNotifyMessage(Message m)
        {
            if (m.Msg != 0x14) // WM_ERASEBKGND
                base.OnNotifyMessage(m);
        }

        protected override void WndProc(ref Message m)
        {
            try
            {
                switch (m.Msg)
                {
                    case WM_COPYDATA:
                        if ((int)m.WParam == m_current_work_order_seq_no)
                        {
                            COPYDATASTRUCT cds = (COPYDATASTRUCT)m.GetLParam(typeof(COPYDATASTRUCT));
                            Dictionary<string, string> infos = null;

                            ParseInfos(cds.lpData, ref infos);

                            switch (infos["todo"])
                            {
                                case "update_status":
                                    int step_index = Convert.ToInt32(infos["step"]);
                                    if (step_index >= 0 && step_index < 4)
                                    {
                                        if (step_index == 0) { _stepval = "기본"; Change_Step(1); }
                                        if (step_index == 1) { _stepval = "VAL"; Change_Step_old(1); Change_Step(2); };
                                        if (step_index == 2) { _stepval = "페이지"; Change_Step_old(2); Change_Step(3); };
                                        if (step_index == 3) { _stepval = "PDF"; Change_Step_old(3); Change_Step(4); };

                                        string[] row9 = { _stepval, infos["status"] };
                                        dataGridView1.Rows.Insert(0, row9);
                                        dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];
                                    }

                                    if (step_index == 99)
                                    {
                                        label15.Text = infos["status"];
                                        progressBar1.Maximum = int.Parse(infos["status"].ToString().Split(':')[1]);
                                        progressBar1.Value = int.Parse(infos["status"].ToString().Split(':')[0]);
                                    }

                                    break;

                                case "report":
                                    if ((int)m.WParam == m_current_work_order_seq_no)
                                    {
                                        WorkResult.m_result = (Convert.ToInt32(infos["result"]) == 1 ? true : false);
                                        WorkResult.m_drawer_code = infos["drawer_code"];
                                        WorkResult.SetErrorCodes(infos["error_codes"]);

                                        delay_exception = true;
                                        Console.WriteLine(WorkResult.m_result + ":결과 1이면 성공");

                                        if (!WorkResult.m_result)
                                        {
                                            Change_Error();
                                            dataGridView1.Rows.Clear();
                                            string[] _ims = WorkResult.ErrorCodesToString().Split('|');
                                            for (int i = 0; i < _ims.Length; i++)
                                            {
                                                string[] row9 = { "에러", _ims[i] };
                                                dataGridView1.Rows.Insert(0, row9);
                                                dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];
                                                dataGridView1.Refresh();
                                            }

                                            //일서스트에서 없는 글꼴 팝업이 뜨면 그이후건은 모두 PDF 생성실폐
                                            if (WorkResult.ErrorCodesToString().Contains("XXXX") || WorkResult.ErrorCodesToString().Contains("0214"))
                                            {
                                                FindAndKillProcess(AutoCC_Sub_name.Substring(0, AutoCC_Sub_name.Length - 4));
                                            }
                                        }

                                        dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];
                                        Delay(5);
                                    }

                                    break;

                                default:
                                    break;
                            }
                        }

                        Application.DoEvents();

                        break;

                    default:
                        base.WndProc(ref m);
                        break;
                }
            }
            catch (Exception ex)
            {
                logger.Debug(ex.ToString());
            }
        }

        #endregion

        #region 화면이동
        private Point mousePoint;

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            mousePoint = new Point(e.X, e.Y);
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                Location = new Point(this.Left - (mousePoint.X - e.X),
                    this.Top - (mousePoint.Y - e.Y));
            }
        }


        #endregion

        string AutoCC_Sub_name = "AutoCC_Sub.exe";

        int m_current_work_order_seq_no = -0;


        String m_category_name = string.Empty;
        int[] m_steps = new int[4];


        Pair<int, int> m_work_count = null;

        Status m_status = Status.NOT_WORKING;

        private ILog logger = null;

        private bool inti_check = true; // 환경번수 설정 오류 확인

        private static bool delay_exception = false; // 지정된 대기 시간을 빠져 나옴

        enum Status
        {
            NOT_WORKING, READY, WORKING, Stop
        }

        private int timeout = 0;
        private int timeout_base = 0;

        private int total_count = 0;
        private int total_Sucess_Count = 0;



        public Main()
        {
            InitializeComponent();

            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            this.SetStyle(ControlStyles.EnableNotifyMessage, true);

            logger = LogManager.GetLogger(this.GetType());
        }



        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point(10, 10);
            this.Show();

            #region 환경변수 초기화



            AppConfiguration.config = new Dictionary<string, string>();

            AppConfiguration.config.Add("Sunday", AppConfiguration.LoadConfig("Sunday"));
            AppConfiguration.config.Add("Monday", AppConfiguration.LoadConfig("Monday"));
            AppConfiguration.config.Add("Tuesday", AppConfiguration.LoadConfig("Tuesday"));
            AppConfiguration.config.Add("Wednesday", AppConfiguration.LoadConfig("Wednesday"));
            AppConfiguration.config.Add("Thursday", AppConfiguration.LoadConfig("Thursday"));
            AppConfiguration.config.Add("Friday", AppConfiguration.LoadConfig("Friday"));
            AppConfiguration.config.Add("Saturday", AppConfiguration.LoadConfig("Saturday"));


            AppConfiguration.config.Add("root_folder", AppConfiguration.LoadConfig("root_folder"));
            AppConfiguration.config.Add("target_folder", AppConfiguration.LoadConfig("target_folder"));
            AppConfiguration.config.Add("preview_folder", AppConfiguration.LoadConfig("preview_folder"));
            AppConfiguration.config.Add("deposit_folder", AppConfiguration.LoadConfig("deposit_folder"));
            AppConfiguration.config.Add("ordered_folder", AppConfiguration.LoadConfig("ordered_folder"));
            AppConfiguration.config.Add("result_folder", AppConfiguration.LoadConfig("result_folder"));

            AppConfiguration.config.Add("log", AppConfiguration.LoadConfig("root_folder") + "log\\");

            if (AppConfiguration.LoadConfig("Machine_info").Equals("null"))
            {
                inti_check = false;
                label1.Text = "장치 ID: 0 / 0";
                MessageBox.Show("장치 번호를 설정해 주세요.");
            }
            else
            {

                AppConfiguration.machine.count = int.Parse(AppConfiguration.LoadConfig("Machine_info").Split('/')[1]);
                AppConfiguration.machine.index = int.Parse(AppConfiguration.LoadConfig("Machine_info").Split('/')[0]);

                label1.Text = "장치 ID : " + AppConfiguration.machine.index + " / " +
                                         AppConfiguration.machine.count;

            }




            #endregion

            timer1 = new System.Windows.Forms.Timer();
            timer1.Interval = 1000;
            timer1.Tick += new EventHandler(Timer1_Tick);
            timer1.Start();

            #region 폰트설정

            //PrivateFontCollection privateFonts = new PrivateFontCollection();

            ////폰트명이 아닌 폰트의 파일명을 적음
            //privateFonts.AddFontFile("DS-DIGI.TTF");


            ////24f 는 출력될 폰트 사이즈
            //Font font = new Font(privateFonts.Families[0], 26f);
            //Font font2 = new Font(privateFonts.Families[0], 12f);


            //label1.Font = font2; label2.Font = font2; label3.Font = font; label4.Font = font2; label5.Font = font2;
            //label6.Font = font2; label7.Font = font2; label8.Font = font2; label9.Font = font2; label10.Font = font2;
            //label11.Font = font2; label12.Font = font2; button1.Font = font2;

            //dataGridView1.Font = font2;

            #endregion

            #region log4j
            String logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log4Net.xml");
            FileInfo file = new FileInfo(logPath);
            XmlConfigurator.Configure(file);
            logger = LogManager.GetLogger(this.GetType());
            #endregion

            logger.Debug("Program Start");

            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.CadetBlue;

            reflash_screen();
        }


        bool _time_stop = false;

        private void RunTask()
        {
            int work_count = GetWorkList();

            bool exception_check = false;

            if (work_count == 0)
            {
                m_status = Status.READY;

                string[] row9 = { "기본" + "", "No Data: " + label12.Text.Replace("현재시간 : ", "") };
                dataGridView1.Rows.RemoveAt(0);
                dataGridView1.Rows.Insert(0, row9);

                return;
            }

            for (int i = 0; i < work_count; i++)
            {
                label15.Text = "";
                progressBar1.Value = 0;

                _time_stop = false;

                label13.Text = i + ":" + work_count;

                delay_exception = false;

                Work work = m_works[i];

                WorkResult.Clear();
                WorkResult.m_result = false;

                label14.Text = "";

                string ext = System.IO.Path.GetExtension(work.file_path);
                string category_prefix = work.category_code.Substring(0, 3);

                bool except_customer = false;

                total_count = total_count + 1; // 전체 처리건수 저장


                #region 예외 경우나 주문 내역 오류 확인

                /**
                 * KHJ
                 * 4.0에서는 사용하지 않는 업체 코드이다. 따라서 아래 코드는 의미없다.
                 */
                #region 의미없는 코드
                if (work.OPICustomerID == "103027" || work.OPICustomerID == "105238" || work.OPICustomerID == "109111") // 부천행복인쇄 , 부천세일프린팅 , (방문)대양어워드
                    except_customer = true;
                else if (work.OPICustomerID == "108458" && ext.Equals("cdr") == true) // 일산상상디자인
                    except_customer = true;
                else if (work.OPICustomerID == "102143" || work.OPICustomerID == "110778")  // 화성디자인하우, (호주)크리에이티브컴바인
                    except_customer = true;
                else if (work.OPICustomerID == "106737")    // 중랑양지씨앤피
                    except_customer = true;
                else if (work.OPICustomerID == "108824" && category_prefix == "005" && work.item_count > 5)   // 아이디컴퍼니
                    except_customer = true;
                #endregion


                bool insufficient_work = false;

                if (work.Verify() == false) insufficient_work = true;

                bool flyer_landscape = false;
                bool sticker_narrow_gap = false;

                if (category_prefix == "005") // 전단
                {
                    if (work.bleed_size.Width > work.bleed_size.Height)
                        flyer_landscape = true;
                }
                else if (category_prefix == "004") // 스티커
                {
                    string sub_type_code_prefix = (work.SubTypeCode.Length < 2 ? work.SubTypeCode : work.SubTypeCode.Substring(0, 2));
                    if (sub_type_code_prefix != "SE")
                    {
                        if ((work.bleed_size.Width != work.trim_size.Width) || (work.bleed_size.Height != work.trim_size.Height))
                        {
                            if ((work.bleed_size.Width < (work.trim_size.Width + 6.0f)) || (work.bleed_size.Height < (work.trim_size.Height + 6.0f)))
                                sticker_narrow_gap = true;
                        }
                    }
                }

                if (except_customer == true)
                {
                    WorkResult.AddErrorCode("7000");    // 예외업체
                }
                else if (insufficient_work == true)
                {
                    WorkResult.AddErrorCode("7001");    // 불충분한 주문정보
                }
                else if (flyer_landscape == true)
                {
                    WorkResult.AddErrorCode("7002");    // 전단 가로형
                }
                else if (sticker_narrow_gap == true)
                {
                    WorkResult.AddErrorCode("7003");
                }
                #endregion 예외 경우나 주문 내역 오류 확인 종료
                else
                {
                    m_current_work_order_seq_no = work.order_common_seqno;

                    Inspect(ref work); // 여기서 Timeout 시간이 결정됨

                    dataGridView1.Rows.Clear();

                    label9.Text = "시작시간 : " + DateTime.Now.ToString("HH-mm-ss");
                    label10.Text = "종료시간 : " + timeout;
                    label2.Text = "업체명 : " + work.CustomerName;
                    label4.Text = "상품명 : " + work.category_name + "(" + work.item_count + ")";
                   
                    Delay(timeout);  // 데이터 처리하는동안 대기 상태 유지      

                    m_current_work_order_seq_no = -0;

                    _time_stop = true;
                }

                #region 서버응답 결과 처리 루틴 테스트
                //if (WorkResult.m_result == true)
                //{
                //    // 데스트를 위한 임시
                //    total_Sucess_Count = total_Sucess_Count + 1;
                //    reflash_screen();

                //    string argumentsp = string.Format("order_id={0};order_seq={1};order_date={2};category_code={3};file_path={4};" +
                //                                       "bleed_size_width={5:F2};bleed_size_height={6:F2};trim_size_width={7:F2};" +
                //                                       "trim_size_height={8:F2};side_count={9};item_count={10};marker={11};magenta_outline={12};" +
                //                                       "works_folder_path={13};log_file_path={14}",
                //                                       work.OPISeq, work.order_common_seqno, work.order_date, work.category_code, work.file_path,
                //                                       work.bleed_size.Width, work.bleed_size.Height, work.trim_size.Width,
                //                                       work.trim_size.Height, work.side_count, work.item_count, work.marker, work.magenta_outline,
                //                                       AppConfiguration.config["result_folder"], AppConfiguration.config["log"]);

                //    logger.Debug("--------------------------------------------------");
                //    logger.Debug(argumentsp);
                //    logger.Debug(WorkResult.ErrorCodesToString());
                //}
                //else
                //{
                //    string argumentsp = string.Format("order_id={0};order_seq={1};order_date={2};category_code={3};file_path={4};" +
                //                                    "bleed_size_width={5:F2};bleed_size_height={6:F2};trim_size_width={7:F2};" +
                //                                    "trim_size_height={8:F2};side_count={9};item_count={10};marker={11};magenta_outline={12};" +
                //                                    "works_folder_path={13};log_file_path={14}",
                //                                    work.OPISeq, work.order_common_seqno, work.order_date, work.category_code, work.file_path,
                //                                    work.bleed_size.Width, work.bleed_size.Height, work.trim_size.Width,
                //                                    work.trim_size.Height, work.side_count, work.item_count, work.marker, work.magenta_outline,
                //                                    AppConfiguration.config["result_folder"], AppConfiguration.config["log"]);

                //    logger.Debug("--------------------------------------------------");
                //    logger.Debug(argumentsp);
                //    logger.Debug(WorkResult.ErrorCodesToString());

                //    //  데스트를 위한 임시
                //    reflash_screen();
                //}

                //reflash_screen();
                #endregion

                #region Data 처리

                string accept_result = string.Empty;
                List<string> erp_results = new List<string>();

                if (WorkResult.m_result == true)
                {
                    #region 처리성공된 파일의 DB 처리

                    FileInfo finfo = new FileInfo(m_works[i].file_path);
                    string path = finfo.DirectoryName;  // + @"\receipt_files";

                    // 미리보기 만들어서 업로드
                    using (MagickImageCollection images = new MagickImageCollection())
                    {
                        images.Read(path + @"\" + m_works[i].order_num + ".pdf");
                        int page = 1;
                        foreach (MagickImage image in images)
                        {
                            if (page == 1)
                            {
                                work.PreviewPath = work.order_num + "_" + page.ToString().PadLeft(3, '0') + ".jpg";
                            }
                            else
                            {
                                work.PreviewPath += "|" + work.order_num + "_" + page.ToString().PadLeft(3, '0') + ".jpg";
                            }
                            image.Alpha(AlphaOption.Remove);
                            image.Write(path + @"\" + m_works[i].order_num + "_" + page.ToString().PadLeft(3, '0') + ".jpg", MagickFormat.Jpeg);
                            page++;
                        }
                    }

                    int ext1 = m_works[i].file_path.LastIndexOf(@"\");
                    string pdf_save_path = path + @"\" + m_works[i].order_num + ".pdf";

                    NameValueCollection strings = new NameValueCollection();
                    strings.Add("ordernum", m_works[i].order_num);
                    strings.Add("ID", "GP4#auto" + (AppConfiguration.machine.index).ToString());

                    NameValueCollection files = new NameValueCollection();
                    files.Add("pdf_file", pdf_save_path); // Path + filename

                    //File.
                    string[] pfiles = Directory.GetFiles(path, m_works[i].order_num + "*.jpg");
                    Array.Sort(pfiles);
                    for (int j = 0; j < pfiles.Length; j++)
                    {
                        files.Add("preview_file" + j.ToString().PadLeft(3, '0'), pfiles[j]);
                    }
                    
                    HttpToServer.Upload(strings, files).GetAwaiter().GetResult();

                    FileInfo ppp = new FileInfo(m_works[i].file_path);
                    try { 
                        Directory.GetFiles(path).ToList().ForEach(File.Delete);
                    } catch
					{

					}

                    DataDispatcher.SetSuccess(work.order_common_seqno);

                    #endregion 처리성공된 파일의 DB 처리 종료

                    #region 처리성공된 파일이동 처리

                    string local_pdf_file_path = string.Format("{0}{1}\\{2}\\{3:00}.pdf", AppConfiguration.config["result_folder"], work.order_date, work.OPISeq, "result");
                    string target_pdf_folder_path = AppConfiguration.config["target_folder"] + @"\" + DateTime.Now.ToString("yyyy") + @"\" + DateTime.Now.ToString("MM") + @"\" + DateTime.Now.ToString("dd");

                    string flyer_order_no = string.Empty;
                    string order_title = string.Empty;

                    string target_pdf_file_name =  @"\" + work.order_num + ".pdf";
                    string target_pdf_file_path = string.Format("{0}{1}", target_pdf_folder_path, target_pdf_file_name);
                    string target_pdf_file_path1 = string.Format("{0}{1}", @"\\172.16.33.135\hotfolder", target_pdf_file_name);//\\172.16.33.90\hotfolder

                    // 2020.07.06 파일정보 DB 입력이 성공 하면 파일 이동작업 시작
                    if (WorkResult.m_result)
                    {
                        string preview_pdf_date_folder_path = string.Format("{0}\\{1}", AppConfiguration.config["preview_folder"], DateTime.Now.ToString("yyyy") + @"\" + DateTime.Now.ToString("MM") + @"\" + DateTime.Now.ToString("dd"));

                        try
                        {
                            string[] param = new string[6];
                            param[0] = "OrderNum=" + work.order_num;
                            param[1] = "AcceptFilePath=" + target_pdf_folder_path.Replace("\\\\172.16.33.224\\gpnas\\ndrive\\attach\\gp", "").Replace(@"\", "/");
                            param[2] = "AcceptFileName=" + work.order_num + ".pdf";
                            param[3] = "PreviewFilePath=" + preview_pdf_date_folder_path.Replace("\\\\172.16.33.224\\gpnas\\ndrive\\attach\\gp", "").Replace(@"\", "/");
                            param[4] = "PreviewFileNames=" + work.PreviewPath.Replace("|", "||");
                            param[5] = "ID=" + "auto" + (AppConfiguration.machine.index).ToString(); ;
                            string result = DataDispatcher.request("OrderReceipt.php", param);
                        }
                        catch(Exception ex)
                        {
                            // failed to copy pdf
                            WorkResult.m_result = false;
                            WorkResult.AddErrorCode(Error.FAILED_TO_COPY_PDF);

                            break;
                        }
                    }
                    #endregion 처리성공된 파일이동 처리 종료

                    if (WorkResult.m_result)
                    {
                        total_Sucess_Count = total_Sucess_Count + 1;
                        reflash_screen();
                    }
                    else
                    {
                        // 2020.07.06 로직 추가 파일 복사나 파일이름 저장시 error 나면 수동으로
                        DataDispatcher.RemoveAccept(work.order_common_seqno, WorkResult.ErrorCodesToString());

                        //처리실페의 화면 처리
                        reflash_screen();

                        logger.Debug("------------------------< 파일 처리 오류 >--------------------------");

                        string argumentsp2 = string.Format("order_id={0};order_seq={1};order_date={2};category_code={3};file_path={4};" +
                                                  "bleed_size_width={5:F2};bleed_size_height={6:F2};trim_size_width={7:F2};" +
                                                  "trim_size_height={8:F2};side_count={9};item_count={10};marker={11};magenta_outline={12};" +
                                                  "works_folder_path={13};log_file_path={14}",
                                                  work.OPISeq, work.order_common_seqno, work.order_date, work.category_code, work.file_path,
                                                  work.bleed_size.Width, work.bleed_size.Height, work.trim_size.Width,
                                                  work.trim_size.Height, work.side_count, work.item_count, work.marker, work.magenta_outline,
                                                  AppConfiguration.config["result_folder"], AppConfiguration.config["log"]);
                        logger.Debug(argumentsp2);
                        logger.Debug(WorkResult.ErrorCodesToString());
                        logger.Debug("--------------------------------------------------------------------");
                    }
                }
                else
                {
                    DataDispatcher.RemoveAccept(work.order_common_seqno, WorkResult.ErrorCodesToString());

                    reflash_screen();
                }

                if (WorkResult.m_error_codes != null)
                {
                    string error_message = string.Empty;

                    if (WorkResult.m_error_codes.Contains(Error.OBJECT_WITH_FILL_LOWER_TONE) == true)
                    {
                        if (error_message != string.Empty)
                            error_message += "|";
                    }

                    if (error_message != string.Empty)
                        DataDispatcher.RequestQCCheck(work.order_date, work.OPISeq, error_message);

                    if (WorkResult.m_error_codes.Contains(Error.Com_Exception) == true)
                    {
                        FindAndKillProcess(AutoCC_Sub_name.Substring(0, AutoCC_Sub_name.Length - 4));
                    }
                }
                #endregion Data 처리종료

                Application.DoEvents();

                Thread.Sleep(1000);

                if (m_status != Status.WORKING && m_status != Status.Stop)
                {
                    exception_check = true; //사용자가 stop 했을경우 확인
                    break;
                }

                if (m_status == Status.Stop) // 서버쪽 응답  이 없는 경우 처리
                {
                    break;
                }
            } // end of [for (int i = 0; i < work_count; i++)]


            // DB에서 조회된 주문걸을 모두 처리 했으니 상태를 Working 에서 READY 로 변경 한다.
            // 그리고 결과릴 초기화
            if (exception_check)
            {
                m_status = Status.NOT_WORKING;
                m_works.Clear();
                button1.BackColor = Color.Red;
                button1.Text = "Stop";
                button1.Enabled = true;
                exception_check = false;
            }
            else
            {
                exception_check = false;
                m_status = Status.READY;
                m_works.Clear();
            }
        }

        private void reflash_screen()
        {
            if (total_Sucess_Count != 0)
            {
                double _ims = GetPercentage(total_Sucess_Count, total_count, 2);

                if (_ims > 70) label3.ForeColor = Color.Blue;
                if (_ims < 70 && _ims > 50) label3.ForeColor = Color.Green;
                if (_ims < 50) label3.ForeColor = Color.Red;

                label3.Text = _ims.ToString("F1") + "%";
            }
            else
            {
                label3.ForeColor = Color.Red;
                label3.Text = "  0%";
            }

            label11.Text = "처리건 : " + total_Sucess_Count + " / " + total_count;

            label9.Text = "시작시간:";
            label10.Text = "종료시간:";
            label2.Text = "업체명 : ";
            label4.Text = "상품명 : ";

            dataGridView1.Rows.Clear();

            timeout = 0;

            label5.ForeColor = Color.Black;
            label6.ForeColor = Color.Black;
            label7.ForeColor = Color.Black;
            label8.ForeColor = Color.Black;

            dataGridView1.Rows.Insert(0, "");
            dataGridView1.Rows.Insert(0, "");
            dataGridView1.Rows.Insert(0, "");
            dataGridView1.Rows.Insert(0, "");

            dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[0];

            init_step();
        }

        string _oldTime = "";
        string _curTime = "";

        int _chk_app_down = 0;

        private void Timer1_Tick(object Sender, EventArgs e)
        {
            _oldTime = label12.Text;
            _curTime = System.DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");

            if (_oldTime.Equals(_curTime))
            {
                _chk_app_down++;
                if (_chk_app_down > 10)
                {
                    MessageBox.Show("프로그램 다운");
                    _chk_app_down = 0;
                }
            }
            else
            {
                label12.Text = "현재시간 : " + _curTime;
            }

            //20초에 한번씩 작업 확인
            if (_count > check_interval)
            {
                _count = 0;
                
                if (AppConfiguration.chk_work_start())
                {
                    if (m_status == Status.READY)
                    {
                        m_status = Status.WORKING;
                        RunTask();
                    }
                }
                else
                {
                    total_count = 0;
                    total_Sucess_Count = 0;
                    reflash_screen();
                }
            }
            else
            {
                if (m_status == Status.READY) _count++;
            }

            if (m_status == Status.WORKING)
            {
                label10.Text = "남은시간 : " + --timeout + "초";

                if (timeout < -10 && !_time_stop)
                {
                    WorkResult.m_result = false;
                    WorkResult.AddErrorCode(Error.Time_Out);
                    delay_exception = true;
                    FindAndKillProcess(AutoCC_Sub_name.Substring(0, AutoCC_Sub_name.Length - 4));
                }
            }
        }

        private int GetWorkList()
        {
            DataDispatcher.GetWorkList4(AppConfiguration.machine, ref m_works);

            int work_count = 0;

            if (m_works != null)
                work_count = m_works.Count;

            return work_count;
        }

        private void ParseInfos(string arguments, ref Dictionary<string, string> infos)
        {
            infos = new Dictionary<string, string>();

            char[] separator1 = { ';' };
            string[] results = arguments.Split(separator1);
            int length = 0;
            try // temp
            {
                length = results.Length;
            }
            catch
            {

            }
            if (length < 10)
            {
                foreach (var item in results)
                {
                    char[] separator2 = { '=' };
                    string[] key_and_value = item.Split(separator2);
                    if (key_and_value.Length == 2)
                        infos[key_and_value[0]] = key_and_value[1];
                }
            }
        }

        private bool Inspect(ref Work work)
        {
            // 기본은 Sub  프로그램은 자동 종료 되지만
            // 사용자 중단이나 기타 이유로 종료가 안된 상태가 발생 하는 경우가 있으므로 시작할때 한번 삭제후 시작.
            kill_AutoCC_Sub(AutoCC_Sub_name.Substring(0, AutoCC_Sub_name.Length - 4));

            bool result = false;

            timeout = 900;
            timeout_base = timeout; // 프로세스 종료 시간을 최동응답 이후 반응이 없는 시간으로 처리 되도록 추가

            init_step();

            Process process = new Process();

            FileInfo finfo = new FileInfo(work.file_path);
            string arguments = string.Format("order_id={0};order_seq={1};order_date={2};category_code={3};file_path={4};" +
                                             "bleed_size_width={5:F2};bleed_size_height={6:F2};trim_size_width={7:F2};" +
                                             "trim_size_height={8:F2};side_count={9};item_count={10};marker={11};magenta_outline={12};" +
                                             "works_folder_path={13};log_file_path={14}",
                                             work.OPISeq, work.order_common_seqno, work.order_date, work.category_code, work.file_path,
                                             work.bleed_size.Width, work.bleed_size.Height, work.trim_size.Width,
                                             work.trim_size.Height, work.side_count, work.item_count, work.marker, work.magenta_outline,
                                             finfo.Directory, AppConfiguration.config["log"]);
            logger.Debug(arguments);
            try
            {
                process.StartInfo.FileName = AutoCC_Sub_name;
                process.StartInfo.Arguments = arguments;
                process.SynchronizingObject = this;
                process.EnableRaisingEvents = true;
                process.Start();
            }
            catch
            {
                result = false;
                WorkResult.SetErrorCodes(Error.TIMEOUT_TO_INSPECT);
                string log = string.Format("[AddErrorCode] {0} (repair: {1})", Error.ToDesc(Error.TIMEOUT_TO_INSPECT), "False");
            }

            return result;
        }


        private int GetTimeOut(string category_code, int side_count, int item_count)
        {
            double init_time = 0;
            double item_time = 0;

            switch (category_code.Substring(0, 1))
            {
                case "B":
                    init_time = (side_count == 1 ? 6.0 : 8.0);
                    item_time = (side_count == 1 ? 0.4 : 0.6);
                    break;
                default:
                    init_time = (side_count == 1 ? 2.4 : 3.6);
                    item_time = (side_count == 1 ? 0.4 : 0.8);
                    break;
            }

            double timeout = (init_time + (double)(item_count - 1) * item_time) * 60;

            return (int)timeout;
        }


        private void Change_Step(int _val)
        {
            if (_val == 1)
            {
                panel3.BackColor = Color.Green;
                label5.ForeColor = Color.White;
            }
            if (_val == 2)
            {
                panel4.BackColor = Color.Green;
                label6.ForeColor = Color.White;
            }
            if (_val == 3)
            {
                panel5.BackColor = Color.Green;
                label7.ForeColor = Color.White;

            }
            if (_val == 4)
            {
                panel6.BackColor = Color.Green;
                label8.ForeColor = Color.White;
            }

            Application.DoEvents();
        }

        private void Change_Error()
        {
            panel3.BackColor = Color.Red;
            panel4.BackColor = Color.Red;
            panel5.BackColor = Color.Red;
            panel6.BackColor = Color.Red;

            label5.ForeColor = Color.White;
            label6.ForeColor = Color.White;
            label7.ForeColor = Color.White;
            label8.ForeColor = Color.White;

            Application.DoEvents();
        }

        private void Change_Step_old(int _val)
        {

            if (_val == 1)
            {
                panel3.BackColor = Color.LightGreen;
                label5.ForeColor = Color.Black;
            }
            if (_val == 2)
            {
                panel4.BackColor = Color.LightGreen;
                label6.ForeColor = Color.Black;
            }
            if (_val == 3)
            {
                panel5.BackColor = Color.LightGreen;
                label7.ForeColor = Color.Black;
            }
            if (_val == 4)
            {
                panel6.BackColor = Color.LightGreen;
                label8.ForeColor = Color.Black;
            }

            Application.DoEvents();
        }

        private void init_step()
        {
            panel3.BackColor = Color.Silver;
            panel4.BackColor = Color.Silver;
            panel5.BackColor = Color.Silver;
            panel6.BackColor = Color.Silver;

            label5.ForeColor = Color.Black;
            label6.ForeColor = Color.Black;
            label7.ForeColor = Color.Black;
            label8.ForeColor = Color.Black;

            Application.DoEvents();
        }

        private void FindAndKillProcess(string file_name)
        {
            try
            {
                foreach (Process process in Process.GetProcesses())
                {
                    if (string.Equals(process.ProcessName, file_name)) process.Kill();
                    //    Delay(10000);
                    if (process.ProcessName.ToLower().Trim().Contains("CorelDRW".ToLower().Trim())) process.Kill();
                    if (process.ProcessName.ToLower().Trim().Contains("Illustrator".ToLower().Trim())) process.Kill();
                }
            }
            catch
            {
            }
        }

        private void kill_AutoCC_Sub(string file_name)
        {
            try
            {
                foreach (Process process in Process.GetProcesses())
                {
                    if (string.Equals(process.ProcessName, file_name)) process.Kill();
                }
            }
            catch
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (m_status == Status.NOT_WORKING)
            {
                m_status = Status.READY;

                button1.BackColor = Color.Green;
                button1.Text = "Started";
                button1.Enabled = true;
            }
            else if (m_status == Status.READY)
            {
                m_status = Status.NOT_WORKING;

                button1.BackColor = Color.Red;
                button1.Text = "Start";
                button1.Enabled = true;
            }
            else if (m_status == Status.WORKING)
            {
                m_status = Status.NOT_WORKING;

                button1.BackColor = Color.Yellow;
                button1.Text = "Stoping";
                button1.Enabled = false;
            }

            Application.DoEvents();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

            if (m_status == Status.NOT_WORKING)
            {
                Application.Exit();
            }
            else
            {
                MessageBox.Show("프로그램이 실행중 입니다.");
            }
        }


        private double GetPercentage(double value, double total, int decimalplaces)
        {
            Console.WriteLine("퍼센트:" + System.Math.Round(value * 100 / total, decimalplaces));
            return System.Math.Round(value * 100 / total, decimalplaces);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            if (m_status != Status.NOT_WORKING)
                return;

            this.TopMost = false;

            Form form_settings = new FormSettings();
            form_settings.ShowDialog();

            this.TopMost = true;
        }

        private void Delay(int MS)
        {
            while (0 < MS)
            {
                System.Windows.Forms.Application.DoEvents();

                MS--;
                System.Threading.Thread.Sleep(1000);

                if (delay_exception) break;
            }
        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_3(object sender, EventArgs e)
        {
            timeout = 30;

            Console.WriteLine("시간:" + GetTimeOut("N11", 2, 30));
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }
    }
}
