using CorelDRAW;
using VGCore;
using iTextSharp.text.pdf;
using log4net;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Threading;
using System.Threading.Tasks;


namespace AutoCC_Sub
{
    class CorelDRAWHandler : DrawerHandler
    {
        protected VGCore.Document m_document = null;
        protected VGCore.Shapes m_shapes = null;
        protected VGCore.Shape m_group_shape = null;
        protected VGCore.Color m_white_color = null;


        protected double m_page_tolerance = 1.0;
        protected double m_wider_page_tolerance = 2.0;

        private ILog logger = null;


        public CorelDRAWHandler()
        {
            m_drawer_index = 1;

            String logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Log4Net.xml");
            FileInfo file = new FileInfo(logPath);
            XmlConfigurator.Configure(file);

            logger = LogManager.GetLogger(this.GetType());
        }


        private VGCore.Application app = null; //2020.03.31 On / Off 방식 변경을 위한 추가


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
                foreach (Process process in Process.GetProcesses())
                {
                    if (string.Equals(process.ProcessName.ToLower().Trim(), "CorelDRW".ToLower().Trim()))
                        process.Kill();
                }

                return false;
            }

            return false;

        }

        public override bool OpenDocument(string file_path, ref List<string> error_codes)
        {
            CloseDocument();

            if (Search_CorelDRAW())
            {
                Type pia_type = Type.GetTypeFromProgID("CorelDRAW.Application");
                app = Activator.CreateInstance(pia_type) as VGCore.Application;

                if (app == null) return false;
                app.Refresh();
                app.InitializeVBA();
                app.Visible = true;
                m_white_color = app.CreateCMYKColor(0, 0, 0, 0);


            }
            else
            {

                Process.Start("CorelDRAW.exe");
                //Delay(7000);
                System.Threading.Thread.Sleep(12000);

                Type pia_type = Type.GetTypeFromProgID("CorelDRAW.Application");
                app = Activator.CreateInstance(pia_type) as VGCore.Application;


                if (app == null) return false;
                app.Refresh();
                app.InitializeVBA();
                app.Visible = true;
                m_white_color = app.CreateCMYKColor(0, 0, 0, 0);

            }

            // open document
            try
            {
                m_document = app.OpenDocument(file_path);
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
                string _ims = m_document.SourceFileVersion.ToString().Replace("cdrVersion", "");

                if (int.Parse(_ims) > 20)
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
                m_document.ReferencePoint = VGCore.cdrReferencePoint.cdrCenter;
                m_document.Unit = VGCore.cdrUnit.cdrMillimeter;
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                return false;
            }

            return true;
        }

        public override void CloseDocument()
        {
            try
            {
                if (m_document != null) m_document.Close();

            }
            catch (Exception ex)
            {

            }

        }

        public override bool Prepare(ref List<string> error_codes)
        {

            try
            {
                if (m_document.Pages.Count > 1)
                {
                    List<int> deleted_page_indexes = null;

                    bool _check_error = false;

                    Parallel.For(0, m_document.Pages.Count, (i) =>
                    {

                        try
                        {
                            if (m_document.Pages[i + 1].Shapes.Count == 0)
                            {
                                if (deleted_page_indexes == null)
                                    deleted_page_indexes = new List<int>();
                                deleted_page_indexes.Add(i);
                            }

                        }
                        catch (Exception ex)
                        {
                            _check_error = true;
                        }
                    });

                    if (deleted_page_indexes != null)
                    {

                        for (int i = (deleted_page_indexes.Count - 1); i >= 0; i--)
                        {

                            m_document.DeletePages(deleted_page_indexes[i] + 1);
                        }
                    }

                    if (_check_error)
                    {
                        Error.AddErrorCode(Error.Com_Exception, false, ref error_codes);
                        return false;
                    }

                    Error.AddErrorCode(Error.DOCUMENT_WITH_NOT_A_PAGE, false, ref error_codes);
                }

                if (m_document.Pages.Count != 1)
                {
                    return false;
                }

                m_document.ActivePage.Shapes.FindShapes("", VGCore.cdrShapeType.cdrGuidelineShape).Shapes.All().Delete();
            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, false, ref error_codes);
                return false;
            }


            return true;
        }

        public override bool CheckPages(string category_code, SizeF bleed_size, SizeF trim_size, int side_count, int item_count, ref List<string> error_codes, Func<int, string, string> dispatch)
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


            if (m_pages != null)
            {
                string cate = category_code.Substring(0, 3);
                if (cate == "004")//스티커
                {

                    int _retFlag = 0;  //2020.04.03 람다식 내부에서 return 이 안된서 추가

                    try
                    {
                        SizeF gap = bleed_size - trim_size;

                        bool _check_error = false;

                        Parallel.ForEach(m_pages, page =>
                        {
                            try
                            {

                                RectangleF page_rect = page.rect;
                                RectangleF outer_rect = new RectangleF(page_rect.Left - 3, page_rect.Top - 3, page_rect.Width + 6, page_rect.Height + 6);

                                VGCore.Shapes shapes1 = (VGCore.Shapes)m_document.ActivePage.Shapes.FindShapes().Shapes;

                                foreach (VGCore.Shape shape in shapes1)
                                {
                                    RectangleF item_rect = new RectangleF((float)shape.LeftX, (float)(m_document.ActivePage.SizeHeight - shape.TopY), (float)shape.SizeWidth, (float)shape.SizeHeight);
                                    if (page_rect.IntersectsWith(item_rect) == true && outer_rect.Contains(item_rect) == false && item_rect.Width < page_rect.Width && item_rect.Height < page_rect.Height)
                                    {
                                        m_pages = null;
                                        break;
                                    }
                                };

                                VGCore.Shapes shapes2 = (VGCore.Shapes)m_document.ActivePage.Shapes.FindShapes("", VGCore.cdrShapeType.cdrCurveShape).Shapes;
                                foreach (VGCore.Shape shape in shapes2)
                                {
                                    if (shape.SizeWidth > 100)
                                    {
                                        RectangleF item_rect = new RectangleF((float)shape.LeftX, (float)(m_document.ActivePage.SizeHeight - shape.TopY), (float)shape.SizeWidth, (float)shape.SizeHeight);
                                        if (page_rect.IntersectsWith(item_rect) == true && outer_rect.Contains(item_rect) == false)
                                        {
                                            m_pages = null;
                                            break;
                                        }
                                    }
                                }

                                if (m_pages == null)
                                {
                                    _retFlag = 1;
                                }

                            }
                            catch (Exception ex)
                            {
                                _check_error = true;
                            }
                        });


                        if (_check_error)
                        {
                            Error.AddErrorCode(Error.Com_Exception, false, ref error_codes);
                            return false;
                        }

                        if (_retFlag == 1)
                        {

                            Error.AddErrorCode(Error.DUMMY_OBJECT_IN_PAGE, false, ref error_codes);
                            return false;
                        }
                    }
                    catch
                    {

                    }
                }
            }

            if (m_pages != null)
            {
                if ((m_pages.Count % side_count) != 0)
                {
                    Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_DOUBLE_SIDE, false, ref error_codes);
                    return false;
                }
                else if (m_pages.Count != (side_count * item_count))
                {
                    Error.AddErrorCode(Error.PAGE_COUNT_NOT_MATCH_TO_ITEM_COUNT, false, ref error_codes);
                    return false;
                }
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

        public override bool CheckValidations(ref List<string> error_codes, Func<int, string, string> dispatch)
        {
            bool result = true;

            try
            {
                m_shapes = m_document.ActivePage.Shapes.FindShapes().Shapes;


                if (m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrContourGroupShape).Count > 0)
                {
                    Error.AddErrorCode(Error.EFFECT_CONTOUR_USED, false, ref error_codes);
                    result = false;
                }

                if (m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "@com.transparency.type").Count > 0)
                {

                    Error.AddErrorCode(Error.EFFECT_TRANSPARENCY_USED, false, ref error_codes);
                    result = false;

                }
                if (m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrTextShape, true).Count > 0)
                {
                    Error.AddErrorCode(Error.TEXT_OBJECT_USED, false, ref error_codes);
                    result = false;

                }
                if (m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "@fill.type='uniform'").Count > 0)
                {
                    VGCore.Shapes shapes = (VGCore.Shapes)m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "@fill.type='uniform'").Shapes;

                    bool er1 = false;
                    bool er2 = false;

                    System.Threading.Tasks.Parallel.For(0, shapes.Count, (i, loopState) =>
                    {
                        VGCore.Shape shape = (VGCore.Shape)shapes[i + 1];

                        if (shape.Fill.UniformColor.IsCMYK == false)
                        {
                            if (shape.Fill.UniformColor.IsSpot == true)
                            {
                                er1 = true;
                            }
                            else
                            {
                                er2 = true;
                            }

                            if (Fixable == true)
                                shape.Fill.UniformColor.ConvertToCMYK();
                            else
                                result = false;
                        }
                    });

                    if (er1) Error.AddErrorCode(Error.OBJECT_WITH_SPOT_COLOR, true, ref error_codes);
                    if (er2) Error.AddErrorCode(Error.OBJECT_WITH_NOT_CMYK, true, ref error_codes);

                    m_shapes = m_document.ActivePage.Shapes.FindShapes().Shapes;
                }
                if (m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "@outline.type='solid'").Count > 0)
                {
                    VGCore.Shapes shapes = (VGCore.Shapes)m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "@outline.type='solid'").Shapes;

                    bool er1 = false;
                    bool er2 = false;

                    //foreach (Shape shape in shapes)
                    System.Threading.Tasks.Parallel.For(0, shapes.Count, (i, loopState) =>
                    {
                        VGCore.Shape shape = (VGCore.Shape)shapes[i + 1];

                        if (shape.Outline.Color.IsCMYK == false)
                        {
                            if (shape.Outline.Color.IsSpot == true)
                            {
                                er1 = true;
                            }
                            else
                            {
                                er2 = true;
                            }
                            if (Fixable == true)
                                shape.Outline.Color.ConvertToCMYK();
                            else
                                result = false;
                        }
                    });

                    if (er1) Error.AddErrorCode(Error.OBJECT_WITH_SPOT_COLOR, true, ref error_codes);
                    if (er2) Error.AddErrorCode(Error.OBJECT_WITH_NOT_CMYK, true, ref error_codes);


                    m_shapes = m_document.ActivePage.Shapes.FindShapes().Shapes;
                }


                if (m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "@fill.type='Fountain'").Count > 0)
                {
                    if (Fixable == false) result = false;

                    Error.AddErrorCode(Error.OBJECT_WITH_FILL_FOUNTAIN, result, ref error_codes);
                }
                if (m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "@fill.type='Hatch'").Count > 0)
                {

                    Error.AddErrorCode(Error.OBJECT_WITH_FILL_HATCH, false, ref error_codes);
                    result = false;
                }
                if (m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "@fill.type='Postscript'").Count > 0)
                {

                    Error.AddErrorCode(Error.OBJECT_WITH_FILL_POSTSCRIPT, false, ref error_codes);
                    result = false;
                }
                if (m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "@fill.type='Texture'").Count > 0)
                {
                    Error.AddErrorCode(Error.OBJECT_WITH_FILL_TEXTURE, false, ref error_codes);
                    result = false;
                }
                if (m_shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, "@com.effects <> null and @com.effects.BlendEffects.Count > 0").Count > 0)
                {


                    Error.AddErrorCode(Error.EFFECT_BLEND_USED, false, ref error_codes);
                    result = false;
                }
                if (m_shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, "@com.effects <> null and @com.effects.CustomEffects.Count > 0").Count > 0)
                {


                    Error.AddErrorCode(Error.EFFECT_CUSTOM_USED, false, ref error_codes);
                    result = false;
                }

                if (m_shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, "@com.effects <> null and @com.effects.DistortionEffects.Count > 0").Count > 0)
                {

                    Error.AddErrorCode(Error.EFFECT_DISTORTION_USED, false, ref error_codes);
                    result = false;
                }

                if (m_shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, "@com.effects <> null and @com.effects.DropShadowEffect <> null").Count > 0)
                {

                    Error.AddErrorCode(Error.EFFECT_DROPSHADOW_USED, false, ref error_codes);
                    result = false;
                }
                if (m_shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, "@com.effects <> null and @com.effects.EnvelopeEffects.Count > 0").Count > 0)
                {
                    Error.AddErrorCode(Error.EFFECT_ENVELOPE_USED, false, ref error_codes);
                    result = false;
                }
                if (m_shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, "@com.effects <> null and @com.effects.ExtrudeEffect <> null").Count > 0)
                {

                    Error.AddErrorCode(Error.EFFECT_EXTRUDE_USED, false, ref error_codes);
                    result = false;
                }
                if (m_shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, "@com.effects <> null and @com.effects.PerspectiveEffects.Count > 0").Count > 0)
                {
                    Error.AddErrorCode(Error.EFFECT_PERSPECTIVE_USED, false, ref error_codes);
                    result = false;
                }
                if (m_shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, "@type = 'effect:plugin:Bevel'").Count > 0)
                {

                    Error.AddErrorCode(Error.EFFECT_BEVEL_USED, false, ref error_codes);
                    result = false;
                }
                if (m_shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, "@com.PowerClip <> null").Count > 0)
                {

                    Error.AddErrorCode(Error.EFFECT_POWERCLIP_USED, false, ref error_codes);
                    result = false;
                }
                if (m_shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, "@com.effects <> null and @com.effects.LensEffect <> null").Count > 0)
                {

                    Error.AddErrorCode(Error.EFFECT_LENS_USED, false, ref error_codes);
                    result = false;
                }


                VGCore.Shapes bitmap_shapes = (VGCore.Shapes)m_shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrBitmapShape, false).Shapes;

                bool _er_c = false;
                System.Threading.Tasks.Parallel.For(0, bitmap_shapes.Count, (i, loopState) =>
                {
                    VGCore.Shape shape = (VGCore.Shape)bitmap_shapes[i + 1];

                    if (shape.Bitmap.Mode == VGCore.cdrImageType.cdrRGBColorImage)
                    {
                        _er_c = true;

                        if (Fixable == true)
                            shape.Bitmap.ConvertTo(VGCore.cdrImageType.cdrCMYKColorImage);
                        else
                            result = false;

                        loopState.Break();
                    }
                });

                if (_er_c) Error.AddErrorCode(Error.OBJECT_WITH_RGB, true, ref error_codes);

            }
            catch
            {
                Error.AddErrorCode(Error.EXCEPTION_OCCURRED_IN_DRAWER, true, ref error_codes);
                result = false;
            }

            return result;
        }

        public override bool CheckObjectCount(ref List<string> error_codes)
        {
            VGCore.Shapes shapes = (VGCore.Shapes)m_document.ActivePage.Shapes.FindShapes().Shapes;

            Console.WriteLine("카운트:" + shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "").Count);

            if (shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "").Count > 10000)
            {
                Error.AddErrorCode(Error.DOCUMENT_WITH_TOO_MANY_OBJECT, false, ref error_codes);
                return false;
            }

            return true;
        }

        public override bool ExportPDF(string work_folder_path, SizeF bleed_size, int side_count, bool magenta_outline, ref List<string> error_codes, Func<int, string, string> dispatch, int item_count)
        {
            try
            {
                if (m_document.FacingPages == true) m_document.FacingPages = false;

                VGCore.Page page = (VGCore.Page)m_document.InsertPagesEx(1, false, m_document.Pages.Count, (double)bleed_size.Width, (double)bleed_size.Height);

                if (page.SizeWidth != bleed_size.Width || page.SizeHeight != bleed_size.Height)
                {


                    Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                    return false;
                }

                m_group_shape.MoveToLayer(m_document.Pages[m_document.Pages.Count].Layers[m_document.Pages[m_document.Pages.Count].Layers.Count]);

                m_document.Pages[m_document.Pages.Count].Layers[m_document.Pages[m_document.Pages.Count].Layers.Count].Printable = true;

                SetPDFMode();

                int file_index = 1;

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

                    VGCore.Shape background_shape = (VGCore.Shape)m_document.Pages[m_document.Pages.Count].ActiveLayer.CreateRectangle2(0, 0, bleed_size.Width, bleed_size.Height);
                    VGCore.ShapeRange shape_range = (VGCore.ShapeRange)m_document.Pages[m_document.Pages.Count].Shapes.Range();
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

                        string file_path = string.Format("{0}{1:00}.pdf", work_folder_path, file_index);
                        m_document.PublishToPDF(file_path);

                        System.IO.FileInfo file_info = new System.IO.FileInfo(file_path);
                        if (file_info.Exists == false)
                        {


                            Error.AddErrorCode(Error.FAILED_TO_CREATE_PDF, false, ref error_codes);
                            return false;
                        }

                        if (magenta_rectangle != null)
                            magenta_rectangle.Delete();

                        file_index++;
                    }
                    else
                    {
                        string side1_file_path = string.Format("{0}{1:00}A.pdf", work_folder_path, file_index);
                        string side2_file_path = string.Format("{0}{1:00}B.pdf", work_folder_path, file_index);

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

                            file_index++;
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

            try
            {

                foreach (VGCore.Shape shape in
                        m_document.ActivePage.Shapes.FindShapes().Shapes.FindShapes(null, VGCore.cdrShapeType.cdrGuidelineShape, false).Shapes)
                {
                    shape.Locked = false;
                }

                m_document.ActivePage.Shapes.FindShapes().Shapes.FindShapes(null, VGCore.cdrShapeType.cdrGuidelineShape, false).Shapes.All().Delete();

                VGCore.Shapes group_shapes = (VGCore.Shapes)m_document.ActivePage.Shapes.FindShapes().Shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrGroupShape, true, "").Shapes;

                int ddd = 0;

                for (int i = 0; i < group_shapes.Count; i++)
                {


                    VGCore.Shape group_shape = (VGCore.Shape)group_shapes[i + 1];
                    if (IsSizeEqual((float)group_shape.SizeWidth, (float)group_shape.SizeHeight, bleed_size.Width, bleed_size.Height, (float)m_page_tolerance) == true)
                    {
                        VGCore.Shape rect_shape = (VGCore.Shape)m_document.Pages[m_document.Pages.Count].ActiveLayer.CreateRectangle2(group_shape.LeftX, group_shape.TopY - group_shape.SizeHeight, group_shape.SizeWidth, group_shape.SizeHeight);
                        VGCore.ShapeRange shape_range = (VGCore.ShapeRange)m_document.Pages[m_document.Pages.Count].Shapes.Range();
                        shape_range.Add((VGCore.Shape)rect_shape);
                        rect_shape.Outline.SetNoOutline();
                        rect_shape.Fill.ApplyNoFill();
                        rect_shape.OrderToBack();

                        m_shapes = m_document.ActivePage.Shapes.FindShapes().Shapes;
                    }
                }

                m_document.ActivePage.Shapes.All().UngroupAll();

                VGCore.Shapes shapes = (VGCore.Shapes)m_document.ActivePage.Shapes.FindShapes().Shapes;

                int error_code1 = 0;
                int error_code2 = 0;

                bool _check_error = false;

                Parallel.For(0, shapes.Count, (i, loopState) =>
                {

                    try
                    {


                        if (ddd % 10 == 0)
                        {
                            dispatch(99, ++ddd + ":" + shapes.Count);//  Console.Write

                        }
                        else
                        {
                            ++ddd;
                        }


                        VGCore.Shape shape = (VGCore.Shape)shapes[i + 1];

                        if (bleed_size.Width == trim_size.Width && bleed_size.Height == trim_size.Height)   // one touch
                        {
                            String cate1 = category_code.Substring(0, 1);
                            if (cate1 == "S")
                            {
                                if (IsSizeEqual((float)shape.SizeWidth, (float)shape.SizeHeight, bleed_size.Width + 6.0f, bleed_size.Height + 6.0f, 2.0f) == true)
                                {
                                    m_pages = null;
                                    //break;
                                    loopState.Stop();
                                }
                            }
                        }

                        if ((shape.SizeWidth > ((double)bleed_size.Width - m_page_tolerance)) &&
                           (shape.SizeWidth < ((double)bleed_size.Width + m_page_tolerance)) &&
                           (shape.SizeHeight > ((double)bleed_size.Height - m_page_tolerance)) &&
                           (shape.SizeHeight < ((double)bleed_size.Height + m_page_tolerance)))
                        {
                            if (CheckTone(shape) == false)
                            {
                                error_code1 = 1;
                            }

                            if (shape.Outline.Width > 0.5)
                            {
                                error_code2 = 1;
                                m_pages = null;
                                loopState.Stop();
                            }

                            String cate1 = category_code.Substring(0, 1);
                            if (cate1 == "S")
                            {
                                if (bleed_size.Width == trim_size.Width && bleed_size.Height == trim_size.Height)   // one touch
                                {
                                    if (shape.Fill.Type == VGCore.cdrFillType.cdrUniformFill)
                                    {
                                        int cyan = shape.Fill.UniformColor.CMYKCyan;
                                        int magenta = shape.Fill.UniformColor.CMYKMagenta;
                                        int yellow = shape.Fill.UniformColor.CMYKYellow;
                                        int black = shape.Fill.UniformColor.CMYKBlack;

                                        if (cyan > 0 || magenta > 0 || yellow > 0 || black > 0)
                                        {
                                            m_pages = null;
                                            loopState.Stop();
                                        }
                                    }
                                }
                            }

                            AddNewPage(i + 1, new RectangleF((float)shape.LeftX, (float)shape.TopY, (float)shape.SizeWidth, (float)shape.SizeHeight));

                            shape.Outline.SetNoOutline();

                        }
                        else if (IsSizeEqual((float)shape.SizeWidth, (float)shape.SizeHeight, trim_size.Width, trim_size.Height, (float)m_page_tolerance) == true)
                        {
                            shape.Outline.SetNoOutline();
                        }
                        else if (IsSizeEqual((float)shape.SizeWidth, (float)shape.SizeHeight, bleed_size.Width, bleed_size.Height, (float)m_wider_page_tolerance) == true)
                        {
                            Page page = new Page();
                            page.object_index = i + 1;
                            page.rect = new RectangleF((float)shape.LeftX, (float)shape.TopY, (float)shape.SizeWidth, (float)shape.SizeHeight);

                            if (m_wrong_pages == null)
                                m_wrong_pages = new List<Page>();
                            m_wrong_pages.Add(page);

                            dispatch(2, "페이지 추출:(" + m_wrong_pages.Count + "/" + item_count + ")");

                        }

                        if ((shape.SizeWidth > ((double)bleed_size.Height - m_page_tolerance)) && (shape.SizeWidth < ((double)bleed_size.Height + m_page_tolerance)) && (shape.SizeHeight > ((double)bleed_size.Width - m_page_tolerance)) && (shape.SizeHeight < ((double)bleed_size.Width + m_page_tolerance)))
                        {
                            if (category_code != "S30")
                                reversed_page_aspect_count++;
                        }
                    }
                    catch (Exception ex)
                    {
                        _check_error = true;
                    }
                });

                if (_check_error)
                {
                    Error.AddErrorCode(Error.Com_Exception, false, ref error_codes);
                    m_pages = null;

                }


                if (error_code1 == 1) Error.AddErrorCode(Error.OBJECT_WITH_FILL_LOWER_TONE, false, ref error_codes);
                if (error_code2 == 1) Error.AddErrorCode(Error.OBJECT_WITH_STROKE_TOO_THICK, false, ref error_codes);

                try
                {
                    VGCore.Shapes all_shapes = (VGCore.Shapes)m_document.ActivePage.Shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, false, string.Empty).Shapes;

                    //VGCore.Shape l_shape = null;


                    // 2020.04.22 병렬처리로 변경 필요
                    Console.WriteLine("시작:" + DateTime.Now.ToString("hh-mm-ss"));

                    //foreach (VGCore.Shape shape in all_shapes)
                    Parallel.For(0, all_shapes.Count, (i, loopState) =>
                    {
                        VGCore.Shape shape = (VGCore.Shape)all_shapes[i + 1];

                        if (shape.Outline.Type != VGCore.cdrOutlineType.cdrNoOutline)
                        {
                            if (shape.Outline.Width > 0.0)
                            {
                                if (shape.Outline.LineJoin != VGCore.cdrOutlineLineJoin.cdrOutlineRoundLineJoin)
                                    shape.Outline.LineJoin = VGCore.cdrOutlineLineJoin.cdrOutlineRoundLineJoin;
                            }
                        }
                    });


                    Console.WriteLine("종료:" + DateTime.Now.ToString("hh-mm-ss"));

                }
                catch
                {

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
                m_pages = null;
            }

            return (m_pages == null ? 0 : m_pages.Count);
        }

        protected override int RemoveOverlappedPages(SizeF margin)
        {
            List<int> overlapped_page_indexes = new List<int>();

            for (int i = 0; i < (m_pages.Count - 1); i++)
            {
                // get trim size
                RectangleF first_page_rect = new RectangleF(m_pages[i].rect.Left + margin.Width / 2.0f, m_pages[i].rect.Top + margin.Height / 2.0f, m_pages[i].rect.Width - margin.Width, m_pages[i].rect.Height - margin.Height);

                for (int j = i + 1; j < m_pages.Count; j++)
                {
                    // get trim size
                    RectangleF second_page_rect = new RectangleF(m_pages[j].rect.Left + margin.Width / 2.0f, m_pages[j].rect.Top + margin.Height / 2.0f, m_pages[j].rect.Width - margin.Width, m_pages[j].rect.Height - margin.Height);

                    if (second_page_rect.IntersectsWith(first_page_rect) == true)
                    {
                        if (IsRectEqual(first_page_rect, second_page_rect, (float)m_page_tolerance) == true)
                        {
                            if (overlapped_page_indexes.Contains(j) == false)
                                overlapped_page_indexes.Add(j);
                        }
                        else
                        {
                            return -1;
                        }
                    }
                }
            }

            overlapped_page_indexes.Sort();


            for (int k = overlapped_page_indexes.Count - 1; k >= 0; k--)
            {
                m_pages.RemoveAt(overlapped_page_indexes[k]);
            }

            int wrong_page_count = (m_wrong_pages == null ? 0 : m_wrong_pages.Count);
            for (int i = (wrong_page_count - 1); i >= 0; i--)
            {
                RectangleF first_page_rect = new RectangleF(m_wrong_pages[i].rect.Left + margin.Width / 2.0f, m_wrong_pages[i].rect.Top + margin.Height / 2.0f, m_wrong_pages[i].rect.Width - margin.Width, m_wrong_pages[i].rect.Height - margin.Height);

                int page_count = (m_pages == null ? 0 : m_pages.Count);
                for (int j = 0; j < page_count; j++)
                {
                    RectangleF second_page_rect = new RectangleF(m_pages[j].rect.Left + margin.Width / 2.0f, m_pages[j].rect.Top + margin.Height / 2.0f, m_pages[j].rect.Width - margin.Width, m_pages[j].rect.Height - margin.Height);
                    if (second_page_rect.IntersectsWith(first_page_rect) == true)
                    {
                        wrong_page_count--;
                        break;
                    }
                }
            }

            if (wrong_page_count > 0)
                return -2;

            return overlapped_page_indexes.Count;
        }

        protected override bool PairPages()
        {
            if (m_pages.Count == 2) return true;


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

        private bool ConvertAllTransparenciesToBitmaps()
        {
            try
            {
                VGCore.Shapes shapes = (VGCore.Shapes)m_document.ActivePage.Shapes.FindShapes().Shapes;

                foreach (VGCore.Shape shape in shapes)
                {
                    if (shape.Transparency.Type != VGCore.cdrTransparencyType.cdrNoTransparency)
                        shape.ConvertToBitmapEx(VGCore.cdrImageType.cdrCMYKColorImage, false, true, 300);
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool ConvertAllFountainsToBitmaps()
        {
            try
            {

                VGCore.Shapes shapes = (VGCore.Shapes)m_document.ActivePage.Shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrNoShape, true, "@fill.type='Fountain'").Shapes;

                foreach (VGCore.Shape shape in shapes)
                {
                    if (shape.Type == VGCore.cdrShapeType.cdrGroupShape)
                        return false;

                    shape.ConvertToBitmapEx(VGCore.cdrImageType.cdrCMYKColorImage, false, true, 300);
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        private bool IsTextShapeInPageRectExist()
        {
            try
            {

                VGCore.Shapes shapes = (VGCore.Shapes)m_document.ActivePage.Shapes.FindShapes(string.Empty, VGCore.cdrShapeType.cdrTextShape, true).Shapes;

                foreach (VGCore.Shape shape in shapes)
                {
                    RectangleF shape_rect = new RectangleF((float)shape.LeftX, (float)-shape.TopY, (float)shape.SizeWidth, (float)shape.SizeHeight);

                    foreach (var page in m_pages)
                    {
                        RectangleF page_rect = new RectangleF(page.rect.Left, -page.rect.Top, page.rect.Width, page.rect.Height);
                        if (page_rect.Contains(shape_rect) == true || page_rect.IntersectsWith(shape_rect) == true)
                            return true;
                    }
                }
            }
            catch
            {

            }

            return false;
        }

        public bool CheckPDFIntegrity(string work_folder_path, int item_count)
        {
            bool result = true;

            try
            {
                m_document.InsertPagesEx(1, false, m_document.Pages.Count, 210.0, 297.0);

                for (int i = 0; i < item_count; i++)
                {
                    string file_path = string.Format("{0}{1:00}.pdf", work_folder_path, i + 1);
                    m_document.Pages[m_document.Pages.Count].ActiveLayer.Import(file_path);

                    VGCore.Shapes shapes = (VGCore.Shapes)m_document.Pages[m_document.Pages.Count].Shapes.FindShapes().Shapes;
                    if (shapes.Count == 0)
                    {
                        result = false;
                        break;
                    }
                }

                m_document.DeletePages(m_document.Pages.Count);
            }
            catch
            {
                result = false;
            }

            return result;
        }

        protected void SetPDFMode()
        {
            m_document.PDFSettings.PublishRange = VGCore.pdfExportRange.pdfCurrentPage;
            m_document.PDFSettings.Keywords = string.Empty;
            m_document.PDFSettings.BitmapCompression = VGCore.pdfBitmapCompressionType.pdfZIP;
            m_document.PDFSettings.JPEGQualityFactor = 2;
            m_document.PDFSettings.TextAsCurves = false;
            m_document.PDFSettings.EmbedFonts = false;
            m_document.PDFSettings.EmbedBaseFonts = false;
            m_document.PDFSettings.TrueTypeToType1 = false;
            m_document.PDFSettings.SubsetFonts = false;
            m_document.PDFSettings.SubsetPct = 80;
            m_document.PDFSettings.CompressText = false;
            m_document.PDFSettings.Encoding = VGCore.pdfEncodingType.pdfBinary;
            m_document.PDFSettings.DownsampleColor = true;
            m_document.PDFSettings.DownsampleGray = true;
            m_document.PDFSettings.DownsampleMono = true;
            m_document.PDFSettings.ColorResolution = 300;
            m_document.PDFSettings.MonoResolution = 300;
            m_document.PDFSettings.GrayResolution = 1200;
            m_document.PDFSettings.Hyperlinks = false;
            m_document.PDFSettings.Bookmarks = false;
            m_document.PDFSettings.Thumbnails = false;
            m_document.PDFSettings.Startup = 0;
            m_document.PDFSettings.ComplexFillsAsBitmaps = true;
            m_document.PDFSettings.Overprints = false;
            m_document.PDFSettings.Halftones = true;
            m_document.PDFSettings.SpotColors = true;
            m_document.PDFSettings.MaintainOPILinks = false;
            m_document.PDFSettings.FountainSteps = 256;
            m_document.PDFSettings.EPSAs = VGCore.pdfEPSAs.pdfPreview;
            m_document.PDFSettings.pdfVersion = VGCore.pdfVersion.pdfVersionPDFX1a;
            m_document.PDFSettings.IncludeBleed = false;
            m_document.PDFSettings.Bleed = 31750;
            m_document.PDFSettings.Linearize = false;
            m_document.PDFSettings.CropMarks = false;
            m_document.PDFSettings.RegistrationMarks = false;
            m_document.PDFSettings.DensitometerScales = false;
            m_document.PDFSettings.FileInformation = false;
            m_document.PDFSettings.ColorMode = VGCore.pdfColorMode.pdfCMYK;
            m_document.PDFSettings.ColorProfile = VGCore.pdfColorProfile.pdfSeparationProfile;
            m_document.PDFSettings.EmbedFilename = string.Empty;
            m_document.PDFSettings.EmbedFile = false;
            m_document.PDFSettings.JP2QualityFactor = 2;
            m_document.PDFSettings.TextExportMode = 0;
            m_document.PDFSettings.PrintPermissions = 0;
            m_document.PDFSettings.EditPermissions = 0;
            m_document.PDFSettings.ContentCopyingAllowed = false;
            m_document.PDFSettings.OpenPassword = string.Empty;
            m_document.PDFSettings.PermissionPassword = string.Empty;
            m_document.PDFSettings.ConvertSpotColors = false;
        }

        protected VGCore.Shapes GetShapesWithSize(VGCore.Document document, VGCore.Page page, SizeF size, double tolerance)
        {
            string query = string.Format("@width > {{{0} mm}} and @width < {{{1} mm}} and @height > {{{2} mm}} and @height < {{{3} mm}}", ((double)size.Width - tolerance), ((double)size.Width + tolerance), ((double)size.Height - tolerance), ((double)size.Height + tolerance));

            VGCore.Shapes shapes = null;
            if (page == null)
                shapes = (VGCore.Shapes)document.ActivePage.Shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, query).Shapes;
            else
                shapes = (VGCore.Shapes)page.Shapes.FindShapes(null, VGCore.cdrShapeType.cdrNoShape, true, query).Shapes;

            return shapes;
        }

        protected VGCore.Shape DrawMagentaOutline(VGCore.Document document, SizeF size, VGCore.Page _val)
        {
            VGCore.Shape magenta_rectangle = (VGCore.Shape)document.ActiveLayer.CreateRectangle(_val.LeftX, _val.TopY, _val.LeftX + size.Width, _val.TopY - size.Height);

            magenta_rectangle.Fill.ApplyNoFill();
            magenta_rectangle.Outline.Color.CMYAssign(0, 255, 0);
            magenta_rectangle.Outline.Width = 0.2;
            magenta_rectangle.AlignToPageCenter(VGCore.cdrAlignType.cdrAlignHCenter, VGCore.cdrTextAlignOrigin.cdrTextAlignBoundingBox);
            magenta_rectangle.AlignToPageCenter(VGCore.cdrAlignType.cdrAlignVCenter, VGCore.cdrTextAlignOrigin.cdrTextAlignBoundingBox);

            return magenta_rectangle;
        }

        protected bool CheckTone(VGCore.Shape shape)
        {
            if (shape.Fill.Type == VGCore.cdrFillType.cdrUniformFill)
            {
                if (shape.Fill.UniformColor.IsCMYK == true)
                {
                    int cyan = shape.Fill.UniformColor.CMYKCyan;
                    int magenta = shape.Fill.UniformColor.CMYKMagenta;
                    int yellow = shape.Fill.UniformColor.CMYKYellow;
                    int black = shape.Fill.UniformColor.CMYKBlack;

                    if (cyan > 0 && magenta == 0 && yellow == 0 && black == 0)
                    {
                        if (cyan <= 7)
                            return false;
                    }
                    else if (cyan == 0 && magenta > 0 && yellow == 0 && black == 0)
                    {
                        if (magenta <= 7)
                            return false;
                    }
                    else if (cyan == 0 && magenta == 0 && yellow > 0 && black == 0)
                    {
                        if (yellow <= 7)
                            return false;
                    }
                    else if (cyan == 0 && magenta == 0 && yellow == 0 && black > 0)
                    {
                        if (black <= 7)
                            return false;
                    }
                }
            }

            return true;
        }
    }
}
