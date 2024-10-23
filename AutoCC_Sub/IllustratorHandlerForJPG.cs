using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Illustrator;

namespace AutoCC_Sub
{
    class IllustratorHandlerForJPG : IllustratorHandler
    {
        string m_file_path = string.Empty;

        public IllustratorHandlerForJPG()
        {
            m_drawer_index = 2;
        }

        public override bool OpenDocument(string file_path, ref List<string> error_codes)
        {
            m_file_path = file_path;

            return base.OpenDocument(file_path, ref error_codes);
        }

        public override bool CheckObjectCount(ref List<string> error_codes)
        {
            return true;
        }

        public override bool CheckValidations(ref List<string> error_codes, Func<int, string, string> dispatch)
        {
            bool result = true;

            try
            {
                if (m_document.DocumentColorSpace != AiDocumentColorSpace.aiDocumentCMYKColor)
                {
                    Error.AddErrorCode(Error.DOCUMENT_WITH_RGB, false, ref error_codes);
                    result = false;
                }

                if (m_document.PageItems.Count == 1)
                {
                    var raster_item = m_document.PageItems[1];

                    if (raster_item.ImageColorSpace != AiImageColorSpace.aiImageCMYK)
                    {
                        Error.AddErrorCode(Error.OBJECT_WITH_RGB, false, ref error_codes);
                        result = false;
                    }

                    double item_width = DimUnit.PtToMM(raster_item.Width);
                    double item_height = DimUnit.PtToMM(raster_item.Height);

                    Size image_size = GetImageSize(m_file_path);
                    double dpi_x = (double)image_size.Width / DimUnit.MMToIn(item_width);
                    double dpi_y = (double)image_size.Height / DimUnit.MMToIn(item_height);
                    string log = string.Format("IllustratorHandlerForJPG.CheckValidations_Normal. dpi = ({0}, {1})", dpi_x, dpi_y);

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

        protected override int ExtractPages(string category_code, SizeF bleed_size, SizeF trim_size, ref List<string> error_codes, Func<int, string, string> dispatch, int item_count)
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

                    double height_ratio = item_height / (double)bleed_size.Height;
                    double resized_item_width = item_width / height_ratio;

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

            return (m_pages == null ? 0 : m_pages.Count);
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
