
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
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
    class IllustratorHandlerForPDF : IllustratorHandler
    {
        public int ItemCount { get; set; }

        private string m_copied_pdf_path = string.Empty;

        private ILog logger = null;

        public IllustratorHandlerForPDF()
        {
            m_drawer_index = 2;

            ItemCount = 0;

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


        private static void Delay(int MS)
        {
            while (0 < MS)
            {

                System.Windows.Forms.Application.DoEvents();

                MS--;
                System.Threading.Thread.Sleep(1000);
            }
        }



        public override bool OpenDocument(string file_path, ref List<string> error_codes)
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
                catch (Exception eeee)
                {
                    Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                    Error.AddErrorCode(Error.FAILED_TO_OPEN_DOCUMENT, false, ref error_codes);
                    return false;
                }
            }

            m_copied_pdf_path = CopyOrgPDF(file_path);

            return (m_copied_pdf_path == string.Empty ? false : true);

        }

        public override void CloseDocument()
        {
            try
            {
                if (m_copied_pdf_path != string.Empty)
                {
                    System.IO.File.Delete(m_copied_pdf_path);
                    m_copied_pdf_path = null;
                }

                if (m_document != null) m_document.Close(2);
                m_document = null;
                m_app = null;
                System.GC.Collect();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Open Document Not Found -> Pass ");
            }
        }

        public override bool Prepare(ref List<string> error_codes)
        {
            return true;
        }

        public override bool CheckPages(string category_code, SizeF bleed_size, SizeF trim_size, int side_count, int item_count, ref List<string> error_codes, Func<int, string, string> dispatch)
        {
            int page_count = GetPDFPageCount(m_copied_pdf_path);

            if (page_count != (side_count * item_count))
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
                        {
                            result = false;
                        }
                    }

                    if (result == true)
                    {
                        if (IsSizeEqual((float)DimUnit.PtToMM(document.Width),
                                         (float)DimUnit.PtToMM(document.Height),
                                         bleed_size.Width, bleed_size.Height, (float)m_page_tolerance) == false)
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

        public override bool CheckValidations(ref List<string> error_codes, Func<int, string, string> dispatch)
        {
            return true;
        }

        public override bool CheckObjectCount(ref List<string> error_codes)
        {
            return true;
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

                if (document.TextFrames.Count != 0)
                {
                    Error.AddErrorCode(Error.TEXT_OBJECT_USED, false, ref error_codes);
                    result = false;
                }

                if (document.LegacyTextItems.Count != 0)
                {
                    Error.AddErrorCode(Error.TEXT_OBJECT_USED, false, ref error_codes);
                    result = false;
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

                bool _check_error = false;

                Dictionary<string, string> _error = new Dictionary<string, string>();

                System.Threading.Tasks.Parallel.For(0, document.PageItems.Count, (i, loopState) =>
                {
                    try
                    {
                        var page_item = document.PageItems[i + 1];

                        if (page_item.Opacity != 100)
                        {
                            _error.Add(i + "0", "A1");
                        }

                        if (page_item.typename == "PathItem")
                        {

                            #region 에러체크

                            if (((page_item.Filled == true) && (page_item.FillColor.typename == "GradientColor")))
                            {
                                result = false;
                                _error.Add(i + "1", "B1");
                            }

                            if (((page_item.Filled == true) && (page_item.FillColor.typename == "PatternColor")))
                            {
                                _error.Add(i + "2", "C1");
                                result = false;
                            }

                            if (((page_item.Filled == true) && (page_item.FillColor.typename == "RGBColor")))
                            {
                                _error.Add(i + "3", "D1");
                                result = false;
                            }

                            if (((page_item.Filled == true) && (page_item.StrokeColor.typename == "RGBColor")))
                            {
                                _error.Add(i + "4", "E1");
                                result = false;
                            }

                            if (((page_item.Filled == true) && (page_item.FillColor.typename == "LabColor")))
                            {
                                _error.Add(i + "5", "F1");
                                result = false;
                            }

                            if (((page_item.Filled == true) && (page_item.FillColor.typename == "Texture")))
                            {
                                _error.Add(i + "6", "G1");
                                result = false;
                            }

                            #endregion

                        }
                        else if (page_item.typename == "RasterItem")
                        {
                            if (page_item.Embedded == false)
                            {
                                _error.Add(i + "7", "H1");
                                result = false;
                            }
                        }

                        if (!result) loopState.Break();
                        Console.WriteLine(i + ":" + result);
                        page_item = null;
                        System.GC.Collect();
                    }
                    catch (Exception ex)
                    {
                        _check_error = true;
                    }
                });

                System.GC.Collect();


                if (_error.Count != 0)
                {
                    string[] _ck = new string[8] { "0", "0", "0", "0", "0", "0", "0", "0" };

                    for (int i = 0; i < _error.Count; i++)
                    {
                        if (_error.Values.ToList()[i] == "A1" && _ck[0] == "0") { _ck[0] = "1"; Error.AddErrorCode(Error.EFFECT_TRANSPARENCY_USED, true, ref error_codes); };
                        if (_error.Values.ToList()[i] == "B1" && _ck[1] == "0") { _ck[1] = "1"; Error.AddErrorCode(Error.OBJECT_WITH_FILL_FOUNTAIN, result, ref error_codes); };
                        if (_error.Values.ToList()[i] == "C1" && _ck[2] == "0") { _ck[2] = "1"; Error.AddErrorCode(Error.OBJECT_WITH_FILL_PATTERN, false, ref error_codes); };
                        if (_error.Values.ToList()[i] == "D1" && _ck[3] == "0") { _ck[3] = "1"; Error.AddErrorCode(Error.OBJECT_WITH_RGB, false, ref error_codes); };
                        if (_error.Values.ToList()[i] == "E1" && _ck[4] == "0") { _ck[4] = "1"; Error.AddErrorCode(Error.OBJECT_WITH_RGB, false, ref error_codes); };
                        if (_error.Values.ToList()[i] == "F1" && _ck[5] == "0") { _ck[5] = "1"; Error.AddErrorCode(Error.OBJECT_WITH_LAB, false, ref error_codes); };
                        if (_error.Values.ToList()[i] == "G1" && _ck[6] == "0") { _ck[6] = "1"; Error.AddErrorCode(Error.OBJECT_WITH_FILL_TEXTURE, false, ref error_codes); };
                        if (_error.Values.ToList()[i] == "H1" && _ck[7] == "0") { _ck[6] = "1"; Error.AddErrorCode(Error.OBJECT_WITH_LINK, false, ref error_codes); };

                    }
                }

                if (_check_error)
                {
                    Error.AddErrorCode(Error.Com_Exception, false, ref error_codes);
                    result = false;
                }

                if (Fixable == true)
                {
                    RasterizeOptions options = new RasterizeOptions();
                    options.Transparency = true;

                    for (int i = (effected_item_list.Count - 1); i >= 0; i--)
                    {
                        if (m_document.Rasterize((PathItem)m_document.PageItems[effected_item_list[i]], (PathItem)m_document.PageItems[effected_item_list[i]].ControlBounds, options) == null)
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

            System.GC.Collect();


            return result;
        }

        private bool _CheckObjectCount(Document document, ref List<string> error_codes)
        {
            if (document.PageItems.Count >= 45000)    // 45000
            {
                Error.AddErrorCode(Error.DOCUMENT_WITH_TOO_MANY_OBJECT, false, ref error_codes);
                return false;
            }

            return true;
        }

        public override bool ExportPDF(string work_folder_path, SizeF bleed_size, int side_count, bool magenta_outline, ref List<string> error_codes, Func<int, string, string> dispatch, int item_count)
        {
            return true;
        }

        public string CopyOrgPDF(string file_path)
        {
            string copied_file_path = string.Empty;

            try
            {
                DateTime now = DateTime.Now;
                string temp_file_title = string.Format("{0:0000}{1:00}{2:00}_{3:00}{4:00}{5:00}", now.Year, now.Month, now.Day, now.Hour, now.Minute, now.Second);

                if (System.IO.Directory.Exists("C:\\Auto_CC\\temp\\") == false)
                    System.IO.Directory.CreateDirectory("C:\\Auto_CC\\temp\\");

                copied_file_path = string.Format("C:\\Auto_CC\\temp\\{0}.pdf", temp_file_title);
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



        public void ExtractPages(string sourcePdfPath, string outputPdfPath, int startPage, int endPage, string _filename)
        {
            PdfReader reader = null;
            iTextSharp.text.Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;

            try
            {
                // Intialize a new PdfReader instance with the contents of the source Pdf file:
                reader = new PdfReader(sourcePdfPath);

                // For simplicity, I am assuming all the pages share the same size
                // and rotation as the first page:
                sourceDocument = new iTextSharp.text.Document(reader.GetPageSizeWithRotation(startPage));

                // Initialize an instance of the PdfCopyClass with the source 
                // document and an output file stream:
                pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath + "\\" + _filename + ".pdf", System.IO.FileMode.Create));

                sourceDocument.Open();

                // Walk the specified range and add the page copies to the output file:
                for (int i = startPage; i <= endPage; i++)
                {
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                }
                sourceDocument.Close();
                reader.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool SplitPDF(string pdf_path, string target_folder_path, int side_count, int item_count)
        {
            bool result = true;

            try
            {
                if (item_count == 1)
                {
                    string new_file_path = string.Format("{0}01.pdf", target_folder_path);

                    System.IO.FileInfo file_info = new System.IO.FileInfo(pdf_path);
                    file_info.MoveTo(new_file_path);
                }
                else
                {
                    string split_method = string.Empty;

                    for (int i = 0; i < item_count; i++)
                    {
                        string sub_method = string.Empty;
                        if (side_count == 1)
                            sub_method = string.Format("{0};", i + 1);
                        else
                            sub_method = string.Format("{0}-{1};", i * 2 + 1, i * 2 + 2);

                        split_method += sub_method;
                    }

                    System.IO.FileInfo file_info = new System.IO.FileInfo(pdf_path);
                    string output_file_path = string.Format("{0}%d.pdf", target_folder_path);

                    CPDFSplitMergeObj pdf_merge = new CPDFSplitMergeObj();
                    pdf_merge.Split(pdf_path, split_method, output_file_path);

                    for (int i = (item_count - 1); i >= 0; i--)
                    {
                        string old_file_path = string.Format("{0}{1}.pdf", target_folder_path, i);

                        try
                        {
                            if (System.IO.File.Exists(old_file_path) == true)
                            {
                                string new_file_path = string.Format("{0}{1:00}.pdf", target_folder_path, i + 1);
                                System.IO.File.Move(old_file_path, new_file_path);
                            }
                        }
                        catch
                        {
                            result = false;
                            break;
                        }
                    }
                }
            }
            catch
            {
                result = false;
            }

            return result;
        }
    }
}
