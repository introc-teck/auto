using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using System.Diagnostics;
using PDFSplitMerge;
using iTextSharp.text.pdf;
using CorelDRAW;
using System.IO;

namespace AutoCC_Sub
{
    class CorelDRAWHandlerForZIP : CorelDRAWHandler
    {
        public int ItemCount { get; set; }
        public string FolderPath { get; set; }

        private string m_category_code = string.Empty;
        private SizeF m_trim_size = SizeF.Empty;

        public CorelDRAWHandlerForZIP()
        {
            m_drawer_index = 1;

            ItemCount = 0;
            FolderPath = string.Empty;
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

        //public override bool ExportPDF(string work_folder_path, SizeF bleed_size, int side_count, bool magenta_outline, ref List<string> error_codes, Func<bool> dispatch)
        public override bool ExportPDF(string work_folder_path, SizeF bleed_size, int side_count, bool magenta_outline, ref List<string> error_codes, Func<int, string, string> dispatch, int item_count)
        {
            string[] file_paths_in_directory = System.IO.Directory.GetFiles(FolderPath);
            int file_path_count = (file_paths_in_directory == null ? 0 : file_paths_in_directory.Length);
            for (int i = 0; i < file_path_count; i++)
            {
                string file_path = file_paths_in_directory[i];

                // 1. open document
                if (_OpenDocument(file_path, ref error_codes) == false)
                    return false;

                dispatch(0, "(" + i + 1 + " / " + item_count + ")파일(Zip)");
                m_pages = null;
                System.GC.Collect();


                // 2. prepare
                if (_Prepare(ref error_codes) == false)
                    return false;

                if (_CheckObjectCount(ref error_codes) == false)
                    return false;

                m_pages = null;

                // 3. check pages
                //if (_CheckPages(m_category_code, bleed_size, m_trim_size, side_count, ref error_codes, dispatch) == false)
                if (_CheckPages(m_category_code, bleed_size, m_trim_size, side_count, ref error_codes, dispatch, item_count) == false)
                    return false;

                // 4. check validations
                //if (_CheckValidations(ref error_codes, dispatch) == false)
                if (_CheckValidations(ref error_codes, dispatch) == false)
                    return false;

                // 5. export pdf
                //if (_ExportPDF(work_folder_path, bleed_size, side_count, i, magenta_outline, ref error_codes, dispatch) == false)
                if (_ExportPDF(work_folder_path, bleed_size, side_count, i, magenta_outline, ref error_codes) == false)
                    return false;

                // 7. close file
                _CloseDocument();

                Thread.Sleep(1000);
            }

            return true;
        }

        private CorelDRAW.Application app = null; //2020.04.01 On / Off 방식 변경을 위한 추가


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
                CloseDocument();
                return false;
            }

            return false;

        }


        private bool _OpenDocument(string file_path, ref List<string> error_codes)
        {

            try
            {
                app = new CorelDRAW.Application();
                app.Refresh();
                app.InitializeVBA();
                app.Visible = true;
                m_white_color = app.CreateCMYKColor(0, 0, 0, 0);
            }
            catch
            {
                Error.AddErrorCode(Error.FAILED_TO_EXECUTE_DRAWER, false, ref error_codes);
                return false;
            }

            try
            {
                m_document = app.OpenDocument(file_path, 0);
                if (m_document == null)
                {
                    Error.AddErrorCode(Error.FAILED_TO_OPEN_DOCUMENT, false, ref error_codes);
                    return false;
                }
            }
            catch
            {
                Error.AddErrorCode(Error.FAILED_TO_OPEN_DOCUMENT, false, ref error_codes);
                return false;
            }

            // check version
            try
            {
                if ((int)m_document.SourceFileVersion > 15 && (int)m_document.SourceFileVersion != 105)
                {
                    Error.AddErrorCode(Error.DRAWER_VERSION_NOT_ALLOWED, false, ref error_codes);
                    return false;
                }
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                return false;
            }

            try
            {
                m_document.ReferencePoint = (VGCore.cdrReferencePoint)cdrReferencePoint.cdrCenter;
                m_document.Unit = (VGCore.cdrUnit)cdrUnit.cdrMillimeter;
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                return false;
            }

            return true;
        }

        private void _CloseDocument()
        {
            try
            {
                app.ActiveDocument.Close();

                if (m_document != null) m_document.Close();

                app = null;


            }
            catch (Exception ex)
            {
                Console.WriteLine("Open Document Not Found -> Pass ");
            }
        }

        private bool _Prepare(ref List<string> error_codes)
        {
            try
            {
                if (m_document.Pages.Count > 1)
                {
                    List<int> deleted_page_indexes = null;
                    for (int i = 0; i < m_document.Pages.Count; i++)
                    {
                        if (m_document.Pages[i + 1].Shapes.Count == 0)
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
                            //							m_document.DeletePages(i + 1);
                            m_document.DeletePages(deleted_page_indexes[i] + 1);
                        }
                    }

                    Error.AddErrorCode(Error.DOCUMENT_WITH_NOT_A_PAGE, false, ref error_codes);
                }

                if (m_document.Pages.Count != 1)
                {
                    return false;
                }

                m_document.ActivePage.Shapes.FindShapes("", (VGCore.cdrShapeType)cdrShapeType.cdrGuidelineShape).Shapes.All().Delete();
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                return false;
            }

            return true;
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


            if ((m_pages.Count % side_count) != 0)
            {
                Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_DOUBLE_SIDE, false, ref error_codes);
                return false;
            }
            else if (m_pages.Count != side_count)
            {
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

            try
            {
                m_group_shape = m_document.ActivePage.Shapes.FindShapes().Shapes.All().Group();
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                return false;
            }

            return true;
        }

        //		private bool _CheckValidations(ref List<string> error_codes, Func<bool> dispatch)
        private bool _CheckValidations(ref List<string> error_codes, Func<int, string, string> dispatch)
        {

            return base.CheckValidations(ref error_codes, dispatch);
        }

        private bool _CheckObjectCount(ref List<string> error_codes)
        {
            return base.CheckObjectCount(ref error_codes);
        }

        //private bool _ExportPDF(string work_folder_path, SizeF bleed_size, int side_count, int item_index, bool magenta_outline, ref List<string> error_codes, Func<bool> dispatch)
        private bool _ExportPDF(string work_folder_path, SizeF bleed_size, int side_count, int item_index, bool magenta_outline, ref List<string> error_codes)

        {
            try
            {
                if (m_document.FacingPages == true)
                    m_document.FacingPages = false;

                CorelDRAW.Page page = (CorelDRAW.Page)m_document.InsertPagesEx(1, false, m_document.Pages.Count, (double)bleed_size.Width, (double)bleed_size.Height);

                if (page.SizeWidth != bleed_size.Width || page.SizeHeight != bleed_size.Height)
                {
                    Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                    return false;
                }

                m_group_shape.MoveToLayer(m_document.Pages[m_document.Pages.Count].Layers[m_document.Pages[m_document.Pages.Count].Layers.Count]);

                SetPDFMode();


                //2020.06.01
                sort_result();

                for (int i = 0; i < m_pages.Count; i++)
                {
                    m_document.Pages[m_document.Pages.Count].Shapes.All().UngroupAll();

                    double width = m_document.Pages[m_document.Pages.Count].Shapes[m_pages[i].object_index].SizeWidth;
                    double height = m_document.Pages[m_document.Pages.Count].Shapes[m_pages[i].object_index].SizeHeight;
                    if (IsSizeEqual((float)width, (float)height, bleed_size.Width, bleed_size.Height, (float)m_page_tolerance) == false)
                    {
                        Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                        return false;
                    }

                    double dMoveX = m_document.Pages[m_document.Pages.Count].Shapes[m_pages[i].object_index].CenterX - (bleed_size.Width / 2.0) - m_document.Pages[m_document.Pages.Count].LeftX;
                    double dMoveY = m_document.Pages[m_document.Pages.Count].Shapes[m_pages[i].object_index].CenterY - (bleed_size.Height / 2.0) - m_document.Pages[m_document.Pages.Count].BottomY;

                    Thread.Sleep(1000);

                    m_document.Pages[m_document.Pages.Count].Shapes.All().Group();

                    Thread.Sleep(1000);

                    double center_x = m_document.Pages[m_document.Pages.Count].Shapes[1].CenterX;
                    m_document.Pages[m_document.Pages.Count].Shapes[1].CenterX = center_x - dMoveX;
                    if (IsValueEqual(m_document.Pages[m_document.Pages.Count].Shapes[1].CenterX, (center_x - dMoveX), m_page_tolerance) == false)
                    {
                        Thread.Sleep(1000);

                        m_document.Pages[m_document.Pages.Count].Shapes[1].CenterX = center_x;
                        m_document.Pages[m_document.Pages.Count].Shapes[1].CenterX = center_x - dMoveX;
                        if (IsValueEqual(m_document.Pages[m_document.Pages.Count].Shapes[1].CenterX, (center_x - dMoveX), m_page_tolerance) == false)
                        {
                            Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                            return false;
                        }
                    }

                    double center_y = m_document.Pages[m_document.Pages.Count].Shapes[1].CenterY;
                    m_document.Pages[m_document.Pages.Count].Shapes[1].CenterY = center_y - dMoveY;
                    if (IsValueEqual(m_document.Pages[m_document.Pages.Count].Shapes[1].CenterY, (center_y - dMoveY), m_page_tolerance) == false)
                    {
                        Thread.Sleep(1000);

                        m_document.Pages[m_document.Pages.Count].Shapes[1].CenterY = center_y;
                        m_document.Pages[m_document.Pages.Count].Shapes[1].CenterY = center_y - dMoveY;
                        if (IsValueEqual(m_document.Pages[m_document.Pages.Count].Shapes[1].CenterY, (center_y - dMoveY), m_page_tolerance) == false)
                        {
                            Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                            return false;
                        }
                    }

                    m_document.Pages[m_document.Pages.Count].Shapes.All().UngroupAll();

                    Shape background_shape = (CorelDRAW.Shape)m_document.Pages[m_document.Pages.Count].ActiveLayer.CreateRectangle2(0, 0, bleed_size.Width, bleed_size.Height);
                    ShapeRange shape_range = (CorelDRAW.ShapeRange)m_document.Pages[m_document.Pages.Count].Shapes.Range();
                    shape_range.Add((VGCore.Shape)background_shape);
                    background_shape.Outline.SetNoOutline();
                    background_shape.Fill.ApplyUniformFill(m_white_color);
                    background_shape.CenterX = m_document.Pages[m_document.Pages.Count].Shapes[1].CenterX;
                    background_shape.CenterY = m_document.Pages[m_document.Pages.Count].Shapes[1].CenterY;
                    background_shape.OrderToBack();

                    Thread.Sleep(1000);

                    m_document.Pages[m_document.Pages.Count].Shapes.All().Group();

                    if (side_count == 1)
                    {
                        VGCore.Shape magenta_rectangle = null;
                        /*
                        if (magenta_outline == true)
                        {
                            VGCore.Page _page = (VGCore.Page)m_document.Pages[m_document.Pages.Count];
                            magenta_rectangle = DrawMagentaOutline(m_document, bleed_size, _page);
                        }
                        */
                        string file_path = string.Format("{0}{1:00}.pdf", work_folder_path, item_index + 1);
                        m_document.PublishToPDF(file_path);

                        System.IO.FileInfo file_info = new System.IO.FileInfo(file_path);
                        if (file_info.Exists == false)
                        {
                            Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                            return false;
                        }

                        if (magenta_rectangle != null)
                            magenta_rectangle.Delete();
                    }
                    else
                    {
                        string side1_file_path = string.Format("{0}{1:00}A.pdf", work_folder_path, item_index + 1);
                        string side2_file_path = string.Format("{0}{1:00}B.pdf", work_folder_path, item_index + 1);

                        if (i % 2 == 0) // first side
                        {
                            VGCore.Shape magenta_rectangle = null;
                            /*
                            if (magenta_outline == true)
                            {
                                VGCore.Page _page = (VGCore.Page)m_document.Pages[m_document.Pages.Count];
                                magenta_rectangle = DrawMagentaOutline(m_document, bleed_size, _page);
                            }
                            */
                            m_document.PublishToPDF(side1_file_path);

                            if (magenta_rectangle != null)
                                magenta_rectangle.Delete();
                        }
                        else            // second side
                        {
                            m_document.PublishToPDF(side2_file_path);
                        }
                    }

                    background_shape.Delete();
                }

                m_document.DeletePages(m_document.Pages.Count);
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
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
