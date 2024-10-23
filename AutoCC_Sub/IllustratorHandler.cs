
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using Illustrator;
using PDFSplitMerge;
using iTextSharp.text.pdf;
using System.IO;
using log4net.Config;
using log4net;
using System.Threading.Tasks;


namespace AutoCC_Sub
{
    class IllustratorHandler : DrawerHandler
    {
        protected Application m_app = null;
        protected Document m_document = null;
        protected Layer m_layer = null;

        protected Type m_type = null;
        protected object m_app_object = null;

        protected double m_page_tolerance = 0.5;

        private ILog logger = null;

        private bool _page_change = false;


        private object lockObject = new object();

        public IllustratorHandler()
        {
            m_drawer_index = 2;

            String logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log4Net.xml");
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
                CloseDocument();
                return false;
            }

            return false;

        }

        public override bool OpenDocument(string file_path, ref List<string> error_codes)
        {

            try
            {

                CloseDocument();

                if (!Search_Illustrator())
                {

                    Process.Start("Illustrator.exe");

                };



                m_type = Type.GetTypeFromProgID("Illustrator.Application", false);

                if (m_type != null)
                {
                    m_app_object = Activator.CreateInstance(m_type);
                    m_app = m_app_object as Illustrator.Application;
                    m_app.UserInteractionLevel = AiUserInteractionLevel.aiDontDisplayAlerts;

                }
                else
                {
                    Error.AddErrorCode(Error.FAILED_TO_EXECUTE_DRAWER, false, ref error_codes);
                    return false;
                }

            }
            catch
            {

                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                Error.AddErrorCode(Error.FAILED_TO_EXECUTE_DRAWER, false, ref error_codes);
                return false;
            }

            if (m_app != null)
            {
                try
                {
                    m_document = m_app.Open(file_path);

                    if (m_document == null)
                    {
                        Error.AddErrorCode(Error.FAILED_TO_OPEN_DOCUMENT, false, ref error_codes);
                        return false;
                    }
                }
                catch
                {

                    Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                    Error.AddErrorCode(Error.FAILED_TO_OPEN_DOCUMENT, false, ref error_codes);
                    return false;
                }

            }

            m_pages = null;

            System.GC.Collect();

            return true;
        }

        public override void CloseDocument()
        {

            try
            {
                if (m_document != null) m_document.Close(2);
                m_document = null;
                m_app = null;
                System.GC.Collect();
            }
            catch (Exception ex)
            {

            }


        }

        public override bool Prepare(ref List<string> error_codes)
        {
            try
            {
                m_app.ExecuteMenuCommand("clearguide");

                if (m_document.Layers.Count > 1)
                {
                    List<int> deleted_page_indexes = null;

                    // 비어있는 페이지 삭제
                    for (int i = 0; i < m_document.Layers.Count; i++)
                    {
                        if (m_document.Layers[i + 1].PageItems.Count == 0)
                        {
                            if (deleted_page_indexes == null)
                                deleted_page_indexes = new List<int>();
                            deleted_page_indexes.Add(i);
                        }
                    }

                    if (deleted_page_indexes != null)
                    {
                        for (int i = (deleted_page_indexes.Count - 1); i >= 0; i--)
                        {
                            m_document.Layers[deleted_page_indexes[i] + 1].Delete();
                        }
                    }

                    if (m_document.Layers.Count > 1)
                    {
                        Error.AddErrorCode(Error.DOCUMENT_WITH_NOT_A_LAYER, false, ref error_codes);
                    }
                }

                if (m_document.Layers.Count == 1)
                {
                    m_layer = m_document.Layers[1];
                    m_layer.Printable = true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {

                Error.AddErrorCode(Error.DOCUMENT_WITH_NOT_A_LAYER, false, ref error_codes);
                return false;
            }

            return true;
        }


        public override bool CheckPages(string category_code, SizeF bleed_size, SizeF trim_size, int side_count, int item_count, ref List<string> error_codes, Func<int, string, string> dispatch)
        {

            if (m_document.PageItems.Count < side_count * item_count)
            {
                Error.AddErrorCode(Error.OBJECT_Count_Small, false, ref error_codes);
                return false;
            }

            dispatch(2, " 아디템 추출 시작: " + DateTime.Now.ToString("HH:mm:ss"));

            if (ExtractPages(category_code, bleed_size, trim_size, ref error_codes, dispatch, item_count) == 0)
            {
                if (!error_codes.Contains(Error.OBJECT_Opacity_Used) &&
                     !error_codes.Contains(Error.OBJECT_PatternColor_Used) &&
                     !error_codes.Contains(Error.OBJECT_RGBColor_Used) &&
                     !error_codes.Contains(Error.OBJECT_LabColor_Used) &&
                     !error_codes.Contains(Error.TEXT_OBJECT_USED)
                    )
                {

                    if (m_pages == null && m_pages_reverse != null)
                    {
                        Error.AddErrorCode(Error.PAGE_NOT_EXIST, false, ref error_codes);
                        return false;
                    }
                    else
                    {
                        if (m_pages == null)
                        {
                            Error.AddErrorCode(Error.PAGE_NOT_EXIST, false, ref error_codes);
                            return false;
                        }
                    }
                }
                else
                {
                    return false;
                }
            }

            if (m_pages.Count == side_count * item_count)
            {
            }
            else
            {
                // 주문건수와 검색된 아이템이 동일 하지 않으면 Overlappe 검사 1건일때

                if (RemoveOverlappedPages(bleed_size - trim_size) > 0)
                {
                    Error.AddErrorCode(Error.OVERLAPPED_PAGES_EXIST, true, ref error_codes);
                }
            }

            if ((m_pages.Count % side_count) != 0)
            {
                Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_DOUBLE_SIDE, false, ref error_codes);
                return false;
            }

            if (m_pages.Count != (side_count * item_count))
            {
                Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_ITEM_COUNT, false, ref error_codes);
                return false;
            }

            if (side_count != 1)
            {

                if (m_pages.Count == 2 && item_count == 1 && side_count == 2) // 2020.05.18 1개 짜리 주문일 경우
                {
                    Page _ims0 = m_pages[0];
                    Page _ims1 = m_pages[1];

                    if (_ims0.rect.Y != _ims1.rect.Y)
                    {
                        RectangleF X1 = m_pages[0].rect;
                        X1.Y = m_pages[1].rect.Y;

                        if (m_pages[0].rect.X < m_pages[1].rect.X)
                        {
                            RectangleF X2 = m_pages[1].rect;
                            X2.X = m_pages[0].rect.X + trim_size.Width + 10;
                            m_pages[1].rect = X2;
                        }
                        else
                        {
                            RectangleF X2 = m_pages[0].rect;
                            X2.X = m_pages[1].rect.X + trim_size.Width + 10;
                            m_pages[0].rect = X2;
                        }
                    }
                }

                bool result = PairPages();

                if (result == false)
                {
                    Error.AddErrorCode(Error.FAILED_TO_PAIR_PAGES, false, ref error_codes);
                    return false;
                }
            }

            return true;
        }

        List<Get_page> _page_items = new List<Get_page>();

        public override bool CheckValidations(ref List<string> error_codes, Func<int, string, string> dispatch)
        {
            bool result = true;

            try
            {
                if (m_document.Artboards.Count > 1)
                {
                    Error.AddErrorCode(Error.ArtboartCountOver, false, ref error_codes);
                    result = false;
                }

                if (m_document.RasterItems.Count > 0)
                {
                }


                if (m_document.PageItems.Count >= 45000)
                {
                    Error.AddErrorCode(Error.DOCUMENT_WITH_TOO_MANY_OBJECT, false, ref error_codes);
                    result = false;
                }

                if (m_document.DocumentColorSpace != AiDocumentColorSpace.aiDocumentCMYKColor)
                {
                    Error.AddErrorCode(Error.DOCUMENT_WITH_RGB, false, ref error_codes);
                    result = false;
                }

                if (m_document.PlacedItems.Count > 0)
                {
                    Error.AddErrorCode(Error.OBJECT_WITH_LINK, false, ref error_codes);
                    result = false;
                }

                if (m_document.NonNativeItems.Count > 0)
                {
                    Error.AddErrorCode(Error.OBJECT_WITH_PDF_LINK, false, ref error_codes);
                    result = false;
                }

                if ((((m_document.TextFrames.Count != 0) || (m_document.LegacyTextItems.Count != 0))))
                {
                    Error.AddErrorCode(Error.TEXT_OBJECT_USED, false, ref error_codes);
                    result = false;
                }

                if (m_document.Selection != null)
                {
                    foreach (var page_item in m_document.Selection)
                    {
                        page_item.Selected = false;
                    }
                }


                if (m_document.Swatches != null)
                {
                    for (int i = (m_document.Swatches.Count - 1); i >= 0; i--)
                    {
                        if (m_document.Swatches[i + 1].Color.typename == "SpotColor")
                        {

                            Error.AddErrorCode(Error.OBJECT_WITH_SPOT_COLOR, true, ref error_codes);
                            m_document.Swatches[i + 1].Delete();

                            break;
                        }
                    }
                }
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                result = false;
            }

            return result;
        }

        public override bool CheckObjectCount(ref List<string> error_codes)
        {
            if (m_document.PageItems.Count >= 45000)    // 45000
            {
                Error.AddErrorCode(Error.DOCUMENT_WITH_TOO_MANY_OBJECT, false, ref error_codes);
                return false;
            }

            return true;
        }

        public override bool ExportPDF(string work_folder_path, SizeF bleed_size, int side_count, bool magenta_outline, ref List<string> error_codes, Func<int, string, string> dispatch, int item_count)
        {
            lock (lockObject)
            {

                try
                {
                    PDFSaveOptions pdf_save_options = new PDFSaveOptions();
                    pdf_save_options.ViewAfterSaving = false;
                    pdf_save_options.PDFPreset = "접수용";
                    pdf_save_options.Compatibility = AiPDFCompatibility.aiAcrobat7;

                    int file_index = 1;

                    sort_result();

                    Dictionary<int, string> Pdf_Create = new Dictionary<int, string>();
                    // 2020-07-21 가끔 동일 면이 2번 출력 되는 경우가 있어서 확인용
                    for (int i = 0; i < m_pages.Count; i++)
                    {
                        if ((i % side_count) == 0)
                        {
                            dispatch(3, "PDF 생성중 : (" + i + 1 + " / " + item_count + ")");
                        }


                        m_document.PageItems[m_pages[i].object_index].Selected = true;

                        // check selected page item size
                        #region 선택된 항목의 error 체크
                        //if (m_document.Selection != null)
                        //{
                        //    if (((object[])m_document.Selection).Length != 1)
                        //    {
                        //        #region 1개이상 선택 되면 return False
                        //        m_pages = null;
                        //        System.GC.Collect();

                        //        Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);

                        //        return false;
                        //        #endregion
                        //    }


                        //    var selected_page_item = m_document.Selection[0];

                        //    SizeF item_size = new SizeF((float)DimUnit.PtToMM(selected_page_item.Width), (float)DimUnit.PtToMM(selected_page_item.Height));
                        //    if (IsSizeEqual(item_size, bleed_size, (float)m_page_tolerance) == false)
                        //    {
                        //        #region 선택된 이미지와 주문 크기 비교 다르면 return false;
                        //        m_pages = null;
                        //        System.GC.Collect();

                        //        Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);                          
                        //        return false;
                        //        #endregion
                        //    }
                        //}
                        //else
                        //{
                        //    #region 선택된 항목이 없다면 return false

                        //    m_pages = null;
                        //    System.GC.Collect();
                        //    Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);                      
                        //    return false;

                        //    #endregion
                        //}
                        #endregion


                        int artboard_index = m_document.Artboards.GetActiveArtboardIndex();
                        m_document.FitArtboardToSelectedArt(artboard_index);
                        m_document.PageItems[m_pages[i].object_index].Selected = false;

                        try
                        {
                            Console.WriteLine("[" + Pdf_Create[i] + "]");

                            Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                            m_pages = null;
                            System.GC.Collect();
                            return false;

                        }
                        catch (Exception ex)
                        {
                            Pdf_Create.Add(m_pages[i].object_index, i + "");
                        }

                        string ff = Pdf_Create[m_pages[i].object_index];

                        Artboard artboard = m_document.Artboards[artboard_index + 1];
                        object[] artboard_rect = artboard.ArtboardRect;
                        SizeF artboard_size = new SizeF((float)DimUnit.PtToMM(Convert.ToDouble(artboard_rect[2]) - Convert.ToDouble(artboard_rect[0])), (float)-DimUnit.PtToMM(Convert.ToDouble(artboard_rect[3]) - Convert.ToDouble(artboard_rect[1])));

                        if (IsSizeEqual(artboard_size, bleed_size, (float)m_page_tolerance) == false)
                        {
                            Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);

                            m_pages = null;
                            System.GC.Collect();

                            return false;
                        }

                        if (side_count == 1)
                        {
                            PathItem magenta_rectangle = null;
                            /*
                            if (magenta_outline == true)
                            {
                                RectangleF bleed_rect = new RectangleF((float)Convert.ToDouble(artboard_rect[0]), (float)Convert.ToDouble(artboard_rect[1]), (float)(Convert.ToDouble(artboard_rect[2]) - Convert.ToDouble(artboard_rect[0])), (float)-(Convert.ToDouble(artboard_rect[3]) - Convert.ToDouble(artboard_rect[1])));
                                magenta_rectangle = DrawOutline(m_document, bleed_rect);
                            }
                            */
                            string file_path = string.Format("{0}{1:00}.pdf", work_folder_path, file_index);
                            m_document.SaveAs(file_path, pdf_save_options);

                            System.IO.FileInfo file_info = new System.IO.FileInfo(file_path);
                            if (file_info.Exists == false)
                            {
                                m_pages = null;
                                System.GC.Collect();

                                Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                                return false;
                            }

                            if (magenta_rectangle != null)
                            {
                                magenta_rectangle.Delete();
                            }

                            file_index++;
                        }
                        else
                        {
                            string side1_file_path = string.Format("{0}{1:00}A.pdf", work_folder_path, file_index);
                            string side2_file_path = string.Format("{0}{1:00}B.pdf", work_folder_path, file_index);

                            if (i % 2 == 0) // first side
                            {
                                PathItem magenta_rectangle = null;
                                /*
                                if (magenta_outline == true)
                                {
                                    RectangleF bleed_rect = new RectangleF((float)Convert.ToDouble(artboard_rect[0]), (float)Convert.ToDouble(artboard_rect[1]), (float)(Convert.ToDouble(artboard_rect[2]) - Convert.ToDouble(artboard_rect[0])), (float)-(Convert.ToDouble(artboard_rect[3]) - Convert.ToDouble(artboard_rect[1])));
                                    magenta_rectangle = DrawOutline(m_document, bleed_rect);
                                }
                                */
                                m_document.SaveAs(side1_file_path, pdf_save_options);

                                if (magenta_rectangle != null)
                                {
                                    magenta_rectangle.Delete();
                                }
                            }
                            else            // second side
                            {
                                m_document.SaveAs(side2_file_path, pdf_save_options);
                                string file_path = string.Format("{0}{1:00}.pdf", work_folder_path, file_index);


                                string[] lstFiles = new string[3];
                                lstFiles[0] = side1_file_path;
                                lstFiles[1] = side2_file_path;

                                file_index++;
                            }
                        }
                    }

                    List<string> lstFiles1 = Directory.GetFiles(work_folder_path).ToList();

                    int ext1 = work_folder_path.LastIndexOf(@"\");
                    string filename = work_folder_path.Substring(0, ext1);
                    ext1 = filename.LastIndexOf(@"\");
                    filename = filename.Substring(ext1 + 1);
                    PdfReader reader1 = null;
                    iTextSharp.text.Document sourceDocument1 = null;
                    PdfCopy pdfCopyProvider1 = null;
                    PdfImportedPage importedPage1;
                    string outputPdfPath1 = work_folder_path + filename + ".pdf";

                    sourceDocument1 = new iTextSharp.text.Document();
                    pdfCopyProvider1 = new PdfCopy(sourceDocument1, new System.IO.FileStream(outputPdfPath1, System.IO.FileMode.Create));
                    sourceDocument1.Open();

                    try
                    {
                        for (int f = 0; f < lstFiles1.Count; f++)
                        {
                            FileInfo finfo = new FileInfo(lstFiles1[f]);
                            if (finfo.Extension != ".pdf") continue;

                            reader1 = new PdfReader(lstFiles1[f]);

                            importedPage1 = pdfCopyProvider1.GetImportedPage(reader1, 1);
                            pdfCopyProvider1.AddPage(importedPage1);

                            reader1.Close();

                            System.IO.FileInfo file_info_2 = new System.IO.FileInfo(lstFiles1[f]);
                            if (file_info_2.Exists == true)
                                file_info_2.Delete();
                        }

                        sourceDocument1.Close();
                    }
                    catch (Exception ex)
                    {
                        m_pages = null;
                        System.GC.Collect();

                        Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                        return false;
                    }

                    m_pages = null;
                    System.GC.Collect();
                }
                catch (Exception eee)
                {
                    m_pages = null;
                    System.GC.Collect();

                    Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                    Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                    return false;
                }
            }
            _page_change = false;

            return true;
        }

        private void sort_result()
        {

            if (m_pages.Count < 2) return;

            for (int i = 0; i < m_pages.Count; i = (i + 2))
            {
                Page _val1 = m_pages[i];
                Page _val2 = m_pages[i + 1];

                if (_val1.rect.X > _val2.rect.X)
                {
                    Page _val3 = m_pages[i];
                    m_pages[i] = m_pages[i + 1];
                    m_pages[i + 1] = _val3;
                }
            }
        }

        protected override int ExtractPages(string category_code, SizeF bleed_size, SizeF trim_size, ref List<string> error_codes, Func<int, string, string> dispatch, int item_count)
        {

            int reversed_page_aspect_count = 0;

            System.GC.Collect();

            try
            {
                bool _check_error = false;

                int[] _ims_error = new int[7];
                _ims_error[0] = 0; _ims_error[1] = 0; _ims_error[2] = 0; _ims_error[3] = 0; _ims_error[4] = 0; _ims_error[5] = 0; _ims_error[6] = 0;

                CancellationTokenSource cts = new CancellationTokenSource();
                ParallelOptions po = new ParallelOptions();
                po.CancellationToken = cts.Token;
                po.MaxDegreeOfParallelism = System.Environment.ProcessorCount;

                int ddd = 0;

                dispatch(2, " 파일 ITEM 체크 시작: " + DateTime.Now.ToString("HH:mm:ss"));

                System.Threading.Tasks.Parallel.For(0, m_document.PageItems.Count, (i, loopState) =>
                {

                    try
                    {

                        #region 출력
                        if (ddd % 10 == 0)
                        {
                            dispatch(99, ++ddd + ":" + m_document.PageItems.Count);

                        }
                        else
                        {
                            ++ddd;
                        }
                        #endregion


                        var page_item = m_document.PageItems[i + 1];

                        Get_page _data = new Get_page();
                        _data.item_type = page_item.typename;
                        _data.item_left = DimUnit.PtToMM(page_item.Left);
                        _data.item_top = DimUnit.PtToMM(page_item.Top);
                        _data.item_width = DimUnit.PtToMM(page_item.Width());
                        _data.item_height = DimUnit.PtToMM(page_item.Height());
                        _data.object_index = i + 1;


                        if ((_data.item_width > ((double)trim_size.Width - m_page_tolerance)) &&
                          (_data.item_width < ((double)trim_size.Width + m_page_tolerance)) &&
                          (_data.item_height > ((double)trim_size.Height - m_page_tolerance)) &&
                          (_data.item_height < ((double)trim_size.Height + m_page_tolerance)))
                        {
                            _data._in_item_check = true;
                        }
                        else
                        {
                            _data._in_item_check = false;
                        }

                        _page_items.Add(_data); // 일러스트에서 Data 를 한번만 읽고 다음 부터는 메모리 참조로


                        if (page_item.typename == "PathItem")
                        {
                            if (page_item.Opacity != 100)
                            {
                                _ims_error[0] = 1;
                            }

                            if (((page_item.Filled == true) && (page_item.FillColor.typename == "PatternColor")))
                            {
                                _ims_error[2] = 1;
                            }

                            if (((page_item.Filled == true) && (page_item.FillColor.typename == "RGBColor")))
                            {
                                _ims_error[3] = 1;
                            }

                            if (((page_item.Filled == true) && (page_item.StrokeColor.typename == "RGBColor")))
                            {
                                _ims_error[4] = 1;
                            }

                            if (((page_item.Filled == true) && (page_item.FillColor.typename == "LabColor")))
                            {
                                _ims_error[5] = 1;
                            }

                            if (((page_item.Filled == true) && (page_item.FillColor.typename == "Texture")))
                            {
                                _ims_error[6] = 1;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _check_error = true;
                    }
                });

                if (_ims_error[0] == 1) Error.AddErrorCode(Error.EFFECT_TRANSPARENCY_USED, false, ref error_codes); //0
                if (_ims_error[2] == 1) Error.AddErrorCode(Error.OBJECT_WITH_FILL_PATTERN, false, ref error_codes); //2
                if (_ims_error[3] == 1) Error.AddErrorCode(Error.OBJECT_WITH_RGB, false, ref error_codes); // FillColor 3
                if (_ims_error[4] == 1) Error.AddErrorCode(Error.OBJECT_WITH_RGB, false, ref error_codes); // StrokeColor 4
                if (_ims_error[5] == 1) Error.AddErrorCode(Error.OBJECT_WITH_LAB, false, ref error_codes); //5
                if (_ims_error[6] == 1) Error.AddErrorCode(Error.OBJECT_WITH_FILL_TEXTURE, false, ref error_codes); //6

                if (_ims_error[0] == 1 || _ims_error[2] == 1 || _ims_error[3] == 1 ||
                      _ims_error[4] == 1 || _ims_error[5] == 1 || _ims_error[6] == 1)
                {
                    return 0;
                }

                if (_check_error)
                {
                    Error.AddErrorCode(Error.Com_Exception, false, ref error_codes);

                    return 0;
                }

                dispatch(2, " 파일 ITEM 체크 종료: " + DateTime.Now.ToString("HH:mm:ss"));

                //내부 절취선 삭제
                System.Threading.Tasks.Parallel.For(0, _page_items.Count, (i, loopState) =>
                {
                    double item_left = _page_items[i].item_left;
                    double item_top = _page_items[i].item_top;
                    double item_width = _page_items[i].item_width;
                    double item_height = _page_items[i].item_height;

                    if ((item_width > ((double)bleed_size.Width - m_page_tolerance)) &&
                        (item_width < ((double)bleed_size.Width + m_page_tolerance)) &&
                        (item_height > ((double)bleed_size.Height - m_page_tolerance)) &&
                        (item_height < ((double)bleed_size.Height + m_page_tolerance)))
                    {
                        AddNewPage(_page_items[i].object_index, new RectangleF((float)item_left, (float)item_top, (float)item_width, (float)item_height));
                        EmptyPageStroke(_page_items[i].object_index);

                        string new_note = string.Format("{0}__##__{1}", _page_items[i].item_note, Convert.ToString(m_pages.Count - 1));
                        _page_items[i].item_note = new_note;
                    }

                    if ((item_width > ((double)bleed_size.Height - m_page_tolerance)) && (item_width < ((double)bleed_size.Height + m_page_tolerance)) && (item_height > ((double)bleed_size.Width - m_page_tolerance)) && (item_height < ((double)bleed_size.Width + m_page_tolerance)))
                        reversed_page_aspect_count++;
                });

                if (bleed_size.Width != trim_size.Width || bleed_size.Height != trim_size.Height)
                {
                    SizeF bleed = new SizeF((bleed_size.Width - trim_size.Width) / (float)2.0, (bleed_size.Height - trim_size.Height) / (float)2.0);

                    // trim size
                    System.Threading.Tasks.Parallel.For(0, _page_items.Count, (i, loopState) =>
                    {
                        if (_page_items[i]._in_item_check)
                        {

                            RectangleF temp_work_rect =
                            new RectangleF((float)_page_items[i].item_left - bleed.Width,
                                           (float)_page_items[i].item_top + bleed.Height,
                                           (float)(_page_items[i].item_width + (double)bleed.Width * 2.0),
                                           (float)(_page_items[i].item_height + (double)bleed.Height * 2.0));

                            bool existed = false;
                            if (m_pages != null)
                            {
                                foreach (var page in m_pages)
                                {
                                    if (page.rect.Left > (temp_work_rect.Left - m_page_tolerance / 2.0) && page.rect.Left < (temp_work_rect.Left + m_page_tolerance / 2.0)
                                        && page.rect.Top > (temp_work_rect.Top - m_page_tolerance / 2.0) && page.rect.Top < (temp_work_rect.Top + m_page_tolerance / 2.0)
                                        && page.rect.Width > (temp_work_rect.Width - m_page_tolerance) && page.rect.Width < (temp_work_rect.Width + m_page_tolerance)
                                        && page.rect.Height > (temp_work_rect.Height - m_page_tolerance) && page.rect.Height < (temp_work_rect.Height + m_page_tolerance))
                                    {
                                        existed = true;
                                        break;
                                    }
                                }
                            }

                            if (existed == false)
                            {
                                AddNewPage(-1, temp_work_rect);
                            }

                            EmptyPageStroke(_page_items[i].object_index);
                        }
                    }); // for end
                }

                if (m_pages != null)
                {
                    System.Threading.Tasks.Parallel.For(0, m_pages.Count, (i, loopState) =>
                    {
                        var page = m_pages[i];

                        if (page.object_index == -1)
                        {
                            PathItem path_item = m_document.PathItems.Add();

                            path_item.Left = DimUnit.MMToPt(page.rect.Left);
                            path_item.Top = DimUnit.MMToPt(page.rect.Top);
                            path_item.Width = DimUnit.MMToPt(page.rect.Width);
                            path_item.Height = DimUnit.MMToPt(page.rect.Height);

                            path_item.Stroked = true;
                            path_item.StrokeColor = new NoColor();
                            path_item.Filled = false;

                            string new_note = string.Format("{0}__##__{1}", path_item.Note, Convert.ToString(i));
                            path_item.Note = new_note;

                            Console.WriteLine("값:" + page.object_index);
                        }
                    });   // end of for (int i = 0; i < m_pages.Count; i++)
                }

                if (m_pages == null && reversed_page_aspect_count > 0)
                    Error.AddErrorCode(Error.REVERSED_PAGE_ASPECT, false, ref error_codes);
            }
            catch
            {
                m_pages = null;
            }

            _page_items.Clear();

            System.GC.Collect();

            return m_pages == null ? 0 : m_pages.Count;
        }

        protected override int RemoveOverlappedPages(SizeF margin)
        {
            List<int> overlapped_page_indexes = new List<int>();

            for (int i = 0; i < (m_pages.Count - 1); i++)
            {
                // get trim size
                RectangleF first_page_rect =
                    new RectangleF(m_pages[i].rect.Left + margin.Width / 2.0f,
                                   m_pages[i].rect.Top + margin.Height / 2.0f,
                                   m_pages[i].rect.Width - margin.Width,
                                   m_pages[i].rect.Height - margin.Height);

                for (int j = i + 1; j < m_pages.Count; j++)
                {
                    // get trim size
                    RectangleF second_page_rect =
                        new RectangleF(m_pages[j].rect.Left + margin.Width / 2.0f,
                                       m_pages[j].rect.Top + margin.Height / 2.0f,
                                       m_pages[j].rect.Width - margin.Width,
                                       m_pages[j].rect.Height - margin.Height);

                    if (second_page_rect.IntersectsWith(first_page_rect) == true)
                    {
                        if (overlapped_page_indexes.Contains(j) == false)
                            overlapped_page_indexes.Add(j);
                    }
                }
            }

            overlapped_page_indexes.Sort();

            for (int k = overlapped_page_indexes.Count - 1; k >= 0; k--)
            {
                m_pages.RemoveAt(overlapped_page_indexes[k]);
            }

            return overlapped_page_indexes.Count;
        }

        protected override bool PairPages()
        {
            m_pages = m_pages.OrderBy(Page => Page.rect.X).ToList();

            List<int> column_page_indexes = new List<int>();

            float left_offset = -1;

            for (int i = 0; i < m_pages.Count; i++)
            {
                if (i == 0)
                {
                    column_page_indexes.Add(i);
                    left_offset = m_pages[i].rect.Left;
                }
                else
                {
                    if (m_pages[i].rect.Left > (left_offset + (float)35.0))
                    {
                        column_page_indexes.Add(i);
                        left_offset = m_pages[i].rect.Left;
                    }
                }
            }

            if ((column_page_indexes.Count % 2) != 0)
                return false;

            List<Page> temp_pages = new List<Page>();

            for (int i = 0; i < column_page_indexes.Count; i++)
            {
                int first_page_index_in_a_column = column_page_indexes[i];
                int last_page_index_in_a_column = -1;
                if (i == (column_page_indexes.Count - 1))
                    last_page_index_in_a_column = m_pages.Count - 1;
                else
                    last_page_index_in_a_column = column_page_indexes[i + 1] - 1;

                List<Page> column_pages = new List<Page>();
                for (int j = first_page_index_in_a_column; j <= last_page_index_in_a_column; j++)
                    column_pages.Add(m_pages[j]);

                column_pages = column_pages.OrderBy(Page => Page.rect.Top).ToList();

                temp_pages.AddRange(column_pages);
            }

            m_pages.Clear();

            List<int> row_counts = new List<int>();
            for (int i = 0; i < column_page_indexes.Count; i++)
            {
                int row_count = 0;
                if (i == (column_page_indexes.Count - 1))   // last
                    row_count = temp_pages.Count - column_page_indexes[i];
                else
                    row_count = column_page_indexes[i + 1] - column_page_indexes[i];

                row_counts.Add(row_count);
            }

            for (int i = 0; i < column_page_indexes.Count; i += 2)
            {
                for (int j = column_page_indexes[i]; j < column_page_indexes[i] + row_counts[i]; j++)
                {
                    if ((j >= temp_pages.Count) || ((j + row_counts[i]) >= temp_pages.Count))
                        return false;

                    m_pages.Add(temp_pages[j]);
                    m_pages.Add(temp_pages[j + row_counts[i]]);
                }
            }

            return true;
        }

        protected void EmptyPageStroke(int index)
        {

            try
            {
                var page_item = m_document.PageItems[index];

                if (page_item.typename.Equals("PathItem") == true)
                {
                    if (page_item.Filled == true)
                    {
                        if (page_item.StrokeColor != page_item.FillColor)
                        {
                            page_item.Stroked = true;
                            page_item.StrokeColor = new NoColor();
                        }
                    }
                    else
                    {
                        page_item.Stroked = true;
                        page_item.StrokeColor = new NoColor();
                    }
                }
            }
            catch
            {

            }
        }

        protected PathItem DrawOutline(Illustrator.Document document, RectangleF rect)
        {
            Illustrator.CMYKColor magenta_color = new Illustrator.CMYKColor();
            magenta_color.Cyan = 0;
            magenta_color.Magenta = 100;
            magenta_color.Yellow = 0;
            magenta_color.Black = 0;

            PathItem path_item = document.PathItems.Rectangle(rect.Top, rect.Left, rect.Width, rect.Height);
            path_item.FillColor = new NoColor();
            path_item.Stroked = true;
            path_item.StrokeColor = magenta_color;
            path_item.StrokeWidth = 0.5;

            return path_item;
        }
    }
}
