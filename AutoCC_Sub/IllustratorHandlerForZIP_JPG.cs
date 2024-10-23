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

namespace AutoCC_Sub
{
    class IllustratorHandlerForZIP_JPG : IllustratorHandler
    {
        public int ItemCount { get; set; }
        public string FolderPath { get; set; }

        public IllustratorHandlerForZIP_JPG(int _val)
        {
            m_drawer_index = 2;

            ItemCount = 0;
            FolderPath = string.Empty;

            _side_count_local = _val;
        }

        private int _side_count_local = 1;

        public override bool OpenDocument(string file_path, ref List<string> error_codes)
        {
            bool result = false;
            try
            {
                string[] file_paths_in_directory = System.IO.Directory.GetFiles(FolderPath);

                if (file_paths_in_directory.Length == (ItemCount * _side_count_local))
                    result = true;
            }
            catch
            {
            }

            return result;
        }

        public override void CloseDocument()
        {
            System.GC.Collect();
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

            List<string> valid_jpg_paths = GetValidJPGPaths(FolderPath, side_count, item_count);

            _side_count_local = side_count;

            if (valid_jpg_paths.Count != (side_count * item_count))
            {
                return false;
            }

            if (valid_jpg_paths != null)
            {
                if (valid_jpg_paths.Count > 0)
                {
                    for (int i = 0; i < valid_jpg_paths.Count; i++)
                    {
                        string file_path = valid_jpg_paths[i];

                        // 1. open document
                        if (_OpenDocument(file_path, ref error_codes) == false)
                        {

                            return false;
                        }

                        // 2. check pages
                        if (_CheckPages(bleed_size, side_count, ref error_codes) == false)
                        {
                            return false;
                        }

                        // 3. check validations
                        if (_CheckValidations(file_path, ref error_codes) == false)
                        {

                            return false;
                        }

                        // 4. export pdf
                        // jpg 는 단면만 처리
                        if (_ExportPDF(work_folder_path, bleed_size, i, 0, magenta_outline, ref error_codes) == false)
                        {
                            return false;
                        }

                        // 6. close file
                        _CloseDocument();

                        System.Threading.Thread.Sleep(1000);
                    }

                    m_pages = null;
                    System.GC.Collect();
                }
            }
            else
            {
                return false;
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
                _CloseDocument();

                if (!Search_Illustrator()) { Process.Start("Illustrator.exe"); Delay(12); };

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
                Console.WriteLine("Open Document Not Found -> Pass ");
            }

        }

        private bool _CheckPages(SizeF bleed_size, int side_count, ref List<string> error_codes)
        {
            int reversed_page_aspect_count = 0;

            try
            {
                // bleed size
                if (m_document.PageItems.Count == 1)
                {
                    var page_item = m_document.PageItems[1];

                    double item_width = DimUnit.PtToMM(page_item.Width());
                    double item_height = DimUnit.PtToMM(page_item.Height());

                    double width_ratio = item_width / (double)bleed_size.Width;
                    double resized_item_height = item_height / width_ratio;

                    double resized_item_width = item_width / width_ratio;

                    if (resized_item_width > ((double)bleed_size.Width - m_page_tolerance) && (resized_item_width < ((double)bleed_size.Width + m_page_tolerance)) && resized_item_height > ((double)bleed_size.Height - m_page_tolerance) && (resized_item_height < ((double)bleed_size.Height + m_page_tolerance)))
                    {
                        double new_item_width = item_width / width_ratio;
                        double new_item_height = item_height / width_ratio;

                        page_item.Left = 0.0f;
                        page_item.Top = DimUnit.MMToPt(new_item_height);
                        page_item.Width = DimUnit.MMToPt(new_item_width);
                        page_item.Height = DimUnit.MMToPt(new_item_height);

                        AddNewPage(1, new RectangleF((float)page_item.Left, (float)page_item.Top, (float)page_item.Width(), (float)page_item.Height()));
                    }

                    if ((item_width > ((double)bleed_size.Height - m_page_tolerance)) && (item_width < ((double)bleed_size.Height + m_page_tolerance)) && (item_height > ((double)bleed_size.Width - m_page_tolerance)) && (item_height < ((double)bleed_size.Width + m_page_tolerance)))
                    {
                        reversed_page_aspect_count++;
                    }
                }


                if (reversed_page_aspect_count > 0)
                {
                    Error.AddErrorCode(Error.REVERSED_PAGE_ASPECT, false, ref error_codes);
                    m_pages = null;
                }

            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
            }

            return (m_pages == null ? false : true);
        }

        private bool _CheckValidations(string file_path, ref List<string> error_codes)
        {
            bool result = true;

            try
            {
                if (m_document.DocumentColorSpace != AiDocumentColorSpace.aiDocumentCMYKColor)
                {
                    Error.AddErrorCode(Error.DOCUMENT_WITH_RGB, false, ref error_codes);
                    result = false;
                }

                if (m_document.RasterItems.Count == 1)
                {
                    var raster_item = m_document.RasterItems[1];

                    if (raster_item.ImageColorSpace != AiImageColorSpace.aiImageCMYK)
                    {
                        Error.AddErrorCode(Error.OBJECT_WITH_RGB, false, ref error_codes);
                        result = false;
                    }

                    double item_width = DimUnit.PtToMM(raster_item.Width);
                    double item_height = DimUnit.PtToMM(raster_item.Height);

                    Size image_size = GetImageSize(file_path);
                    double dpi_x = (double)image_size.Width / DimUnit.MMToIn(item_width);
                    double dpi_y = (double)image_size.Height / DimUnit.MMToIn(item_height);


                    if (dpi_x < 300.0 || dpi_y < 300.0)
                    {
                        Error.AddErrorCode(Error.OBJECT_WITH_LOWER_DPI, false, ref error_codes);
                        result = false;
                    }
                }

            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
            }

            return result;
        }

        private bool _ExportPDF(string work_folder_path, SizeF bleed_size, int item_index, int side_index, bool magenta_outline, ref List<string> error_codes)
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

                m_document.PageItems[m_pages[item_index].object_index].Selected = true;

                // check selected page item size
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
                m_document.PageItems[m_pages[item_index].object_index].Selected = false;

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
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                return false;
            }

            return true;
        }

        private bool _RenamePDFs(string work_folder_path, int item_index, ref List<string> error_codes)
        {
            string side1_file_path = string.Format("{0}{1:00}A.pdf", work_folder_path, item_index + 1);
            string target_file_path = string.Format("{0}{1:00}.pdf", work_folder_path, item_index + 1);
            System.IO.FileInfo side1_file_info = new System.IO.FileInfo(side1_file_path);
            try
            {
                side1_file_info.MoveTo(target_file_path);
            }
            catch
            {

            }

            try
            {
                System.IO.FileInfo target_file_info = new System.IO.FileInfo(target_file_path);
                if (target_file_info.Exists == false)
                {
                    Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                    return false;
                }
            }
            catch
            {
                Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                return false;
            }

            try
            {
                System.IO.FileInfo file_info_1 = new System.IO.FileInfo(side1_file_path);
                if (file_info_1.Exists == true)
                    file_info_1.Delete();
            }
            catch
            {

            }

            return true;
        }

        private bool _MergePDFs(string work_folder_path, int item_index, ref List<string> error_codes)
        {
            string side1_file_path = string.Format("{0}{1:00}A.pdf", work_folder_path, item_index + 1);
            string side2_file_path = string.Format("{0}{1:00}B.pdf", work_folder_path, item_index + 1);

            string file_path = string.Format("{0}{1:00}.pdf", work_folder_path, item_index + 1);


            string[] lstFiles = new string[3];

            lstFiles[0] = side1_file_path;
            lstFiles[1] = side2_file_path;


            PdfReader reader = null;
            iTextSharp.text.Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage;
            string outputPdfPath = file_path;

            sourceDocument = new iTextSharp.text.Document();
            pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));
            sourceDocument.Open();

            try
            {
                for (int f = 0; f < lstFiles.Length - 1; f++)
                {
                    reader = new PdfReader(lstFiles[f]);

                    importedPage = pdfCopyProvider.GetImportedPage(reader, 1);
                    pdfCopyProvider.AddPage(importedPage);

                    reader.Close();
                }

                sourceDocument.Close();
            }
            catch (Exception ex)
            {
                Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                return false;
            }

            try
            {
                System.IO.FileInfo file_info = new System.IO.FileInfo(file_path);
                if (file_info.Exists == false)
                {
                    Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                    return false;
                }
            }
            catch
            {
                Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                return false;
            }

            try
            {
                System.IO.FileInfo file_info_1 = new System.IO.FileInfo(side1_file_path);
                if (file_info_1.Exists == true)
                    file_info_1.Delete();
                System.IO.FileInfo file_info_2 = new System.IO.FileInfo(side2_file_path);
                if (file_info_2.Exists == true)
                    file_info_2.Delete();
            }
            catch
            {

            }

            return true;
        }

        private List<string> GetValidJPGPaths(string folder_path, int side_count, int _item_count)
        {
            List<string> valid_jpg_paths = null;

            string[] file_paths_in_directory = System.IO.Directory.GetFiles(FolderPath);

            if (file_paths_in_directory.Length != side_count * _item_count) return valid_jpg_paths; // 2020.06.01 압축파일에 jpg 여러건 처리 불가

            int file_path_count = (file_paths_in_directory == null ? 0 : file_paths_in_directory.Length);
            if ((file_path_count % side_count) == 0)
            {
                for (int i = 0; i < file_path_count; i++)
                {
                    string file_path = string.Empty;
                    switch (side_count)
                    {
                        case 1:
                            file_path = string.Format("{0}____{1:00}A.jpg", FolderPath, i + 1);
                            System.IO.File.Copy(file_paths_in_directory[i], file_path, true);
                            break;
                        case 2:
                            file_path = string.Format("{0}____{1:00}{2}.jpg", FolderPath, i / 2 + 1, (i % 2) == 0 ? "A" : "B");
                            System.IO.File.Copy(file_paths_in_directory[i], file_path, true);
                            break;
                    }

                    if (System.IO.File.Exists(file_path) == true)
                    {
                        if (valid_jpg_paths == null)
                            valid_jpg_paths = new List<string>();
                        valid_jpg_paths.Add(file_path);
                    }
                    else
                    {
                        valid_jpg_paths = null;
                        break;
                    }
                }
            }

            return valid_jpg_paths;
        }

        private Size GetImageSize(string image_path)
        {
            Size size = Size.Empty;

            try
            {
                Bitmap bitmap = new Bitmap(image_path, false);
                size = new Size(bitmap.Width, bitmap.Height);
            }
            catch
            {
            }

            return size;
        }
    }
}
