using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Diagnostics;
using Illustrator;
using PDFSplitMerge;
using iTextSharp.text.pdf;
using log4net;
using System.IO;
using log4net.Config;

namespace AutoCC_Sub
{
    class IllustratorHandlerForZIP : IllustratorHandler
    {
        public int ItemCount { get; set; }
        public string FolderPath { get; set; }

        private string m_category_code = string.Empty;
        private SizeF m_trim_size = SizeF.Empty;

        private ILog logger = null;

        public IllustratorHandlerForZIP()
        {
            m_drawer_index = 2;

            ItemCount = 0;
            FolderPath = string.Empty;


            String logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log4Net.xml");
            FileInfo file = new FileInfo(logPath);
            XmlConfigurator.Configure(file);

            logger = LogManager.GetLogger(this.GetType());
        }

        public override bool OpenDocument(string file_path, ref List<string> error_codes)
        {
            bool result = false;
            try
            {
                string[] file_paths_in_directory = System.IO.Directory.GetFiles(FolderPath);
                if (file_paths_in_directory.Length == ItemCount)
                    result = true;
            }
            catch
            {
            }

            return result;
        }

        public override void CloseDocument()
        {
            try
            {
                if (m_document != null) m_document.Close(2); // 2: aiDoNotSaveChanges
            }
            catch (Exception ex)
            {

            }
        }

        public override bool Prepare(ref List<string> error_codes)
        {
            return true;
        }

        public override bool CheckPages(string category_code, SizeF bleed_size, SizeF trim_size, int side_count, int item_count, ref List<string> error_codes, Func<int, string, string> dispatch)
        {
            m_category_code = category_code;
            m_trim_size = trim_size;

            return true;
        }

        public override bool CheckValidations(ref List<string> error_codes, Func<int, string, string> dispatch)
        {
            return true;
        }

        public override bool CheckObjectCount(ref List<string> error_codes)
        {
            return true;
        }

        public override bool ExportPDF(string work_folder_path, SizeF bleed_size, int side_count, bool magenta_outline, ref List<string> error_codes, Func<int, string, string> dispatch, int item_count)
        {

            string[] file_paths_in_directory = System.IO.Directory.GetFiles(FolderPath);

            int Myitem_count = file_paths_in_directory.Length / item_count; //2020.05.13 주문파일이 분리 되어있는경우 기능추가

            int file_path_count = (file_paths_in_directory == null ? 0 : file_paths_in_directory.Length);

            for (int i = 0; i < file_path_count; i++)
            {
                string file_path = file_paths_in_directory[i];

                // 1. open document
                if (_OpenDocument(file_path, ref error_codes) == false)
                    return false;

                dispatch(0, "(" + i + 1 + " / " + Myitem_count + ")파일(Zip)");

                m_pages = null;
                System.GC.Collect();

                // 2. prepare
                if (_Prepare(ref error_codes) == false)
                    return false;

                if (_CheckObjectCount(ref error_codes) == false)
                    return false;

                m_pages = null;

                // 4. check validations
                if (_CheckValidations(ref error_codes, dispatch) == false)
                    return false;

                // 3. check pages
                if (_CheckPages(m_category_code, bleed_size, m_trim_size, side_count, ref error_codes, dispatch, Myitem_count) == false)
                    return false;

                // 5. export pdf
                if (_ExportPDF(work_folder_path, bleed_size, side_count, i, magenta_outline, ref error_codes) == false)
                    return false;

                // 6. close file
                _CloseDocument();

                System.Threading.Thread.Sleep(5000);
            }

            return true;
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

        private bool _LoadApp(ref List<string> error_codes)
        {

            try
            {

                CloseDocument();

                if (!Search_Illustrator()) { Process.Start("Illustrator.exe"); Delay(12); }; // 2020.04.01 n / Off 방식 변경을 위한 조건 추가 

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

            return true;
        }

        private bool _OpenDocument(string file_path, ref List<string> error_codes)
        {
            if (_LoadApp(ref error_codes) == false)
                return false;

            try
            {

                if (m_document != null) m_document.Close(2);
                m_document = null;

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

            return true;
        }

        private void _CloseDocument()
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

        private bool _Prepare(ref List<string> error_codes)
        {
            try
            {

                // 1. 모든 페이지의 가이드 라인 삭제
                // 2. empty page 삭제

                m_app.ExecuteMenuCommand("clearguide");

                if (m_document.Layers.Count > 1)
                {
                    List<int> deleted_page_indexes = null;


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
                    m_document.Layers[1].Printable = true;
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

        private void Delay(int MS)
        {
            while (0 < MS)
            {

                System.Windows.Forms.Application.DoEvents();

                MS--;
                System.Threading.Thread.Sleep(1000);
            }

        }

        private bool _CheckPages(string category_code, SizeF bleed_size, SizeF trim_size, int side_count, ref List<string> error_codes, Func<int, string, string> dispatch, int item_count)
        {
            if (ExtractPages(category_code, bleed_size, trim_size, ref error_codes, dispatch, item_count) == 0)
            {
                Error.AddErrorCode(Error.PAGE_NOT_EXIST, false, ref error_codes);
                return false;
            }

            int overlapped_page_count = RemoveOverlappedPages(bleed_size - trim_size);
            if (overlapped_page_count > 0)
            {
                Error.AddErrorCode(Error.OVERLAPPED_PAGES_EXIST, true, ref error_codes);
            }
            else if (overlapped_page_count == -1)
            {
                Error.AddErrorCode(Error.OVERLAPPED_PAGES_EXIST, false, ref error_codes);
                return false;
            }
            else if (overlapped_page_count == -2)
            {
                Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_ITEM_COUNT, false, ref error_codes);
                return false;
            }

            try
            {
                if ((m_pages.Count % side_count) != 0)
                {
                    Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_DOUBLE_SIDE, false, ref error_codes);
                    return false;
                }
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_DOUBLE_SIDE, false, ref error_codes);
                return false;
            }

            try
            {
                if (m_pages.Count != side_count)
                {
                    Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_ITEM_COUNT, false, ref error_codes);
                    return false;
                }
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_ITEM_COUNT, false, ref error_codes);
                return false;
            }

            if (side_count != 1)
            {
                bool result = PairPages();
                if (result == false)
                {
                    Error.AddErrorCode(Error.FAILED_TO_PAIR_PAGES, false, ref error_codes);
                    return false;
                }


            }

            return true;
        }

        private bool _CheckValidations(ref List<string> error_codes, Func<int, string, string> dispatch)
        {
            return base.CheckValidations(ref error_codes, dispatch);
        }

        private bool _CheckObjectCount(ref List<string> error_codes)
        {
            return base.CheckObjectCount(ref error_codes);
        }

        private bool _ExportPDF(string work_folder_path, SizeF bleed_size, int side_count, int item_index, bool magenta_outline, ref List<string> error_codes)
        {
            try
            {
                PDFSaveOptions pdf_save_options = new PDFSaveOptions();
                pdf_save_options.ViewAfterSaving = false;
                pdf_save_options.PDFPreset = "DP";

                if (m_document.Artboards.Count > 1)
                {
                    for (int i = m_document.Artboards.Count - 1; i > 0; i--)
                    {
                        m_document.Artboards[i + 1].Delete();
                    }
                }

                sort_result();

                for (int i = 0; i < m_pages.Count; i++)
                {
                    m_document.PageItems[m_pages[i].object_index].Selected = true;

                    if (m_document.Selection != null)
                    {
                        if (((object[])m_document.Selection).Length != 1)
                        {
                            Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);

                            return false;
                        }

                        var selected_page_item = m_document.Selection[0];
                        SizeF item_size = new SizeF((float)DimUnit.PtToMM(selected_page_item.Width), (float)DimUnit.PtToMM(selected_page_item.Height));
                        if (IsSizeEqual(item_size, bleed_size, (float)m_page_tolerance) == false)
                        {
                            Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);

                            return false;
                        }
                    }
                    else
                    {
                        Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);

                        return false;
                    }

                    int artboard_index = m_document.Artboards.GetActiveArtboardIndex();
                    m_document.FitArtboardToSelectedArt(artboard_index);
                    m_document.PageItems[m_pages[i].object_index].Selected = false;

                    Artboard artboard = m_document.Artboards[artboard_index + 1];
                    object[] artboard_rect = artboard.ArtboardRect;
                    SizeF artboard_size = new SizeF((float)DimUnit.PtToMM(Convert.ToDouble(artboard_rect[2]) - Convert.ToDouble(artboard_rect[0])), (float)-DimUnit.PtToMM(Convert.ToDouble(artboard_rect[3]) - Convert.ToDouble(artboard_rect[1])));
                    if (IsSizeEqual(artboard_size, bleed_size, (float)m_page_tolerance) == false)
                    {
                        Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);

                        return false;
                    }

                    if (artboard_size.Width != bleed_size.Width || artboard_size.Height != bleed_size.Height)
                    {
                        float width_gap = (float)DimUnit.MMToPt(artboard_size.Width - bleed_size.Width);
                        float height_gap = (float)DimUnit.MMToPt(artboard_size.Height - bleed_size.Height);

                        float left = (float)Convert.ToDouble(artboard_rect[0]) + width_gap / 2.0f;
                        float top = (float)Convert.ToDouble(artboard_rect[1]) - height_gap / 2.0f;
                        float right = (float)Convert.ToDouble(artboard_rect[2]) - width_gap / 2.0f;
                        float bottom = (float)Convert.ToDouble(artboard_rect[3]) + height_gap / 2.0f;
                        artboard.ArtboardRect = new object[] { left, top, right, bottom };

                        artboard_rect = artboard.ArtboardRect;
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
                        string file_path = string.Format("{0}{1:00}.pdf", work_folder_path, item_index + 1);
                        m_document.SaveAs(file_path, pdf_save_options);

                        System.IO.FileInfo file_info = new System.IO.FileInfo(file_path);
                        if (file_info.Exists == false)
                        {
                            Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                            return false;
                        }

                        if (magenta_rectangle != null)
                        {
                            magenta_rectangle.Delete();
                        }
                    }
                    else
                    {
                        string side1_file_path = string.Format("{0}{1:00}A.pdf", work_folder_path, item_index + 1);
                        string side2_file_path = string.Format("{0}{1:00}B.pdf", work_folder_path, item_index + 1);

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

                            string file_path = string.Format("{0}{1:00}.pdf", work_folder_path, item_index + 1);
                        }
                    }
                }
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                return false;
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
    }
}
