using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Diagnostics;

using Illustrator;
using PDFSplitMerge;

namespace AutoCC_Sub
{
    class IllustratorHandlerForZIP_PDF : IllustratorHandler
    {
        public int ItemCount { get; set; }
        public string FolderPath { get; set; }

        private string m_copied_pdf_path = string.Empty;

        public IllustratorHandlerForZIP_PDF()
        {
            m_drawer_index = 2;

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
            try
            {
                if (m_copied_pdf_path != string.Empty) System.IO.File.Delete(m_copied_pdf_path);
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
            return true;
        }

        public override bool CheckPages(string category_code, SizeF bleed_size, SizeF trim_size, int side_count, int item_count, ref List<string> error_codes, Func<int, string, string> dispatch)
        {
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
            int file_path_count = (file_paths_in_directory == null ? 0 : file_paths_in_directory.Length);
            for (int i = 0; i < file_path_count; i++)
            {
                string file_path = file_paths_in_directory[i];

                // 1. open document
                if (_OpenDocument(file_path, ref error_codes) == false)
                    return false;

                dispatch(0, "(" + i + 1 + " / " + item_count + ")파일(Zip_PDF)");
                m_pages = null;
                System.GC.Collect();

                // 3. check pages
                if (_CheckPages(bleed_size, side_count, ref error_codes) == false)
                    return false;

                // 5. export pdf
                if (_ExportPDF(work_folder_path, i) == false)
                    return false;

                // 6. close file
                _CloseDocument();

                System.Threading.Thread.Sleep(1000);
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

        private void Delay(int MS)
        {
            while (0 < MS)
            {
                System.Windows.Forms.Application.DoEvents();

                MS--;
                System.Threading.Thread.Sleep(1000);
            }
        }

        private bool _OpenDocument(string file_path, ref List<string> error_codes)
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

            m_copied_pdf_path = CopyOrgPDF(file_path);

            return (m_copied_pdf_path == string.Empty ? false : true);
        }

        private void _CloseDocument()
        {
            try
            {
                m_app.ActiveDocument.Close();

                if (m_document != null) m_document.Close(2); // 2: aiDoNotSaveChanges

                m_app = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Open Document Not Found -> Pass ");
            }
        }

        private bool _CheckPages(SizeF bleed_size, int side_count, ref List<string> error_codes)
        {
            int page_count = GetPDFPageCount(m_copied_pdf_path);
            if (page_count != side_count)
            {
                Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_ITEM_COUNT, false, ref error_codes);
                return false;
            }

            bool result = true;

            try
            {
                PDFFileOptions pdf_file_options = m_app.Preferences.PDFFileOptions;

                for (int i = 0; i < page_count; i++)
                {
                    pdf_file_options.PageToOpen = i + 1;
                    Document document = m_app.Open(m_copied_pdf_path);
                    if (document == null)
                    {
                        Error.AddErrorCode(Error.FAILED_TO_OPEN_DOCUMENT, false, ref error_codes);
                        result = false;
                    }

                    if (result == true)
                    {
                        if (_CheckObjectCount(document, ref error_codes) == false)
                            return false;
                    }

                    if (result == true)
                    {
                        if (IsSizeEqual((float)DimUnit.PtToMM(document.Width), (float)DimUnit.PtToMM(document.Height), bleed_size.Width, bleed_size.Height, (float)m_page_tolerance) == false)
                        {
                            Error.AddErrorCode(Error.PAGE_NOT_EXIST, false, ref error_codes);
                            result = false;
                        }
                    }

                    if (result == true)
                    {
                        if (_CheckValidations(document, ref error_codes) == false)
                        {
                            result = false;
                        }
                    }

                    if (document != null)
                        document.Close(AiSaveOptions.aiDoNotSaveChanges);
                }   // end of for (int i = 0; i < page_count; i++)
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                result = false;
            }

            return result;
        }

        private bool _CheckValidations(Document document, ref List<string> error_codes)
        {
            bool result = true;

            try
            {
                if (document.DocumentColorSpace != AiDocumentColorSpace.aiDocumentCMYKColor)
                {
                    Error.AddErrorCode(Error.DOCUMENT_WITH_RGB, false, ref error_codes);
                    result = false;
                }

                if (document.PlacedItems.Count > 0)
                {
                    Error.AddErrorCode(Error.OBJECT_WITH_LINK, false, ref error_codes);
                    result = false;
                }

                if (document.NonNativeItems.Count > 0)
                {
                    Error.AddErrorCode(Error.OBJECT_WITH_PDF_LINK, false, ref error_codes);
                    result = false;
                }

                if ((((document.TextFrames.Count != 0) || (document.LegacyTextItems.Count != 0))))
                {
                    {
                        Error.AddErrorCode(Error.TEXT_OBJECT_USED, false, ref error_codes);
                        result = false;
                    }
                }

                if (document.LegacyTextItems.Count != 0)
                {
                    {
                        Error.AddErrorCode(Error.TEXT_OBJECT_USED, false, ref error_codes);
                        result = false;
                    }
                }

                if (document.Selection != null)
                {
                    foreach (var page_item in document.Selection)
                    {
                        page_item.Selected = false;
                    }
                }

                if (document.Swatches != null)
                {
                    for (int i = (document.Swatches.Count - 1); i >= 0; i--)
                    {
                        if (document.Swatches[i + 1].Color.typename == "SpotColor")
                        {
                            Error.AddErrorCode(Error.OBJECT_WITH_SPOT_COLOR, true, ref error_codes);
                            break;
                        }
                    }
                }

                List<int> effected_item_list = new List<int>(); // object index

                for (int i = 0; i < document.PageItems.Count; i++)
                {
                    var page_item = document.PageItems[i + 1];

                    if (page_item.Opacity != 100)
                    {
                        Error.AddErrorCode(Error.EFFECT_TRANSPARENCY_USED, true, ref error_codes);
                    }

                    if (page_item.typename == "PathItem")
                    {
                        PathItem path_item = page_item;

                        if (((path_item.Filled == true) && (path_item.FillColor.typename == "GradientColor")))
                        {
                            if (Fixable == false)
                                result = false;
                            Error.AddErrorCode(Error.OBJECT_WITH_FILL_FOUNTAIN, result, ref error_codes);
                        }

                        if (((path_item.Filled == true) && (path_item.FillColor.typename == "PatternColor")))
                        {
                            Error.AddErrorCode(Error.OBJECT_WITH_FILL_PATTERN, false, ref error_codes);
                            result = false;
                        }

                        if (((path_item.Filled == true) && (path_item.FillColor.typename == "RGBColor")))
                        {
                            Error.AddErrorCode(Error.OBJECT_WITH_RGB, false, ref error_codes);
                            result = false;
                        }

                        if (((path_item.Filled == true) && (path_item.StrokeColor.typename == "RGBColor")))
                        {
                            Error.AddErrorCode(Error.OBJECT_WITH_RGB, false, ref error_codes);
                            result = false;
                        }

                        if (((path_item.Filled == true) && (path_item.FillColor.typename == "LabColor")))
                        {
                            Error.AddErrorCode(Error.OBJECT_WITH_LAB, false, ref error_codes);
                            result = false;
                        }

                        if (((path_item.Filled == true) && (path_item.FillColor.typename == "Texture")))
                        {
                            Error.AddErrorCode(Error.OBJECT_WITH_FILL_TEXTURE, false, ref error_codes);
                            result = false;
                        }
                    }
                    else if (page_item.typename == "RasterItem")
                    {
                        RasterItem raster_item = page_item;
                        if (raster_item.Embedded == false)
                        {
                            Error.AddErrorCode(Error.OBJECT_WITH_LINK, false, ref error_codes);
                            result = false;
                        }
                    }
                }

                if (Fixable == true)
                {
                    RasterizeOptions options = new RasterizeOptions();
                    options.Transparency = true;

                    for (int i = (effected_item_list.Count - 1); i >= 0; i--)
                    {
                        PathItem path_item = (PathItem)m_document.PageItems[effected_item_list[i]];
                        if (m_document.Rasterize(path_item, path_item.ControlBounds, options) == null)
                            result = false;
                    }
                }

                if (document.RasterItems.Count > 0)
                {
                    foreach (RasterItem raster_item in document.RasterItems)
                    {
                        if (raster_item.Embedded == false)
                        {
                            result = false;
                            break;
                        }
                    }
                }
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                return false;
            }

            return result;
        }

        private bool _CheckObjectCount(Document document, ref List<string> error_codes)
        {
            if (document.PageItems.Count >= 10000)    // 45000
            {
                Error.AddErrorCode(Error.DOCUMENT_WITH_TOO_MANY_OBJECT, false, ref error_codes);
                return false;
            }

            return true;
        }

        private bool _ExportPDF(string work_folder_path, int item_index)
        {
            string new_file_path = string.Format("{0}{1:00}.pdf", work_folder_path, item_index + 1);

            System.IO.FileInfo file_info = new System.IO.FileInfo(m_copied_pdf_path);
            file_info.MoveTo(new_file_path);

            return true;
        }

        private string CopyOrgPDF(string file_path)
        {
            string copied_file_path = string.Empty;

            try
            {
                DateTime now = DateTime.Now;
                string temp_file_title = string.Format("{0:0000}{1:00}{2:00}_{3:00}{4:00}{5:00}_{6:0000}", now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second, now.Millisecond);

                System.IO.FileInfo file_info = new System.IO.FileInfo(file_path);
                copied_file_path = string.Format("{0}\\{1}.pdf", file_info.DirectoryName, temp_file_title);
                if (System.IO.File.Exists(copied_file_path) == true)
                {
                    System.IO.File.Delete(copied_file_path);
                }
                System.IO.File.Copy(file_path, copied_file_path);
            }
            catch
            {
                copied_file_path = string.Empty;
            }

            return copied_file_path;
        }

        private int GetPDFPageCount(string file_path)
        {
            int count = 0;

            try
            {
                iTextSharp.text.pdf.PdfReader reader = null;
                System.Text.UTF8Encoding encoding = new System.Text.UTF8Encoding();
                reader = new iTextSharp.text.pdf.PdfReader(file_path);
                reader.RemoveUnusedObjects();
                count = reader.NumberOfPages;
            }
            catch
            {
            }

            return count;
        }
    }
}
