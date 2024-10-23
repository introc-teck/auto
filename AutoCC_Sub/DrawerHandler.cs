using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace AutoCC_Sub
{
	abstract class DrawerHandler
	{
		public bool Fixable { get; set; }

		public int m_drawer_index = -1;
		//		public bool m_preview_needed = false;

		protected List<Page> m_pages = null;
		protected List<Page> m_pages_reverse = null; //2020.05.19 가로세가가 바뀐경우 처리를 위해 추가

		protected List<Page> m_wrong_pages = null;

		public abstract bool OpenDocument(string file_path,  ref List<string> error_codes);
		public abstract void CloseDocument();
		public abstract bool Prepare(ref List<string> error_codes);
		public abstract bool CheckPages(string category_code, SizeF bleed_size, SizeF trim_size, int side_count, int item_count, ref List<string> error_codes, Func<int, string, string> dispatch);
		public abstract bool CheckValidations(ref List<string> error_codes, Func<int, string, string> dispatch);
		public abstract bool CheckObjectCount(ref List<string> error_codes);
		public abstract bool ExportPDF(string work_folder_path, SizeF bleed_size, int side_count, bool magenta_outline, ref List<string> error_codes, Func<int, string, string> dispatch,int item_count);

		protected abstract int ExtractPages(string category_code, SizeF bleed_size, SizeF trim_size, ref List<string> error_codes, Func<int, string, string> dispatch,int item_count);
		protected abstract int RemoveOverlappedPages(SizeF margin);
		protected abstract bool PairPages();

		protected class Page
		{
			public int object_index { get; set; }
			public RectangleF rect { get; set; }

			public Page()
			{
				object_index = -1;
				rect = RectangleF.Empty;
			}
		}

		protected void AddNewPage(int object_index, RectangleF rect)
		{
			Page page = new Page();
			page.object_index = object_index;
			page.rect = rect;

			if (m_pages == null)
				m_pages = new List<Page>();
			m_pages.Add(page);
		}

		protected void AddNewPage_reverse(int object_index, RectangleF rect)
		{
			Page page = new Page();
			page.object_index = object_index;
			page.rect = rect;

			if (m_pages_reverse == null)
				m_pages_reverse = new List<Page>();
			m_pages_reverse.Add(page);
		}

		protected bool IsSizeEqual(SizeF size1, SizeF size2, float tolerance)
		{
			if ((size1.Width > (size2.Width - tolerance)) && (size1.Width < (size2.Width + tolerance)) && (size1.Height > (size2.Height - tolerance)) && (size1.Height < (size2.Height + tolerance)))
				return true;

			return false;
		}

		protected bool IsSizeEqual(float width1, float height1, float width2, float height2, float tolerance)
		{
			if ((width1 > (width2 - tolerance)) && (width1 < (width2 + tolerance)) && (height1 > (height2 - tolerance)) && (height1 < (height2 + tolerance)))
				return true;

			return false;
		}

		protected bool IsRectEqual(RectangleF rect1, RectangleF rect2, float tolerance)
		{
			if ((rect1.Left > (rect2.Left - tolerance)) && (rect1.Left < (rect2.Left + tolerance)) && (rect1.Top > (rect2.Top - tolerance)) && (rect1.Top < (rect2.Top + tolerance)))
			{
				if ((rect1.Width > (rect2.Width - tolerance)) && (rect1.Width < (rect2.Width + tolerance)) && (rect1.Height > (rect2.Height - tolerance)) && (rect1.Height < (rect2.Height + tolerance)))
					return true;
			}

			return false;
		}

		protected bool IsPointEqual(PointF point1, PointF point2, float tolerance)
		{
			if ((point1.X > (point2.X - tolerance)) && (point1.X < (point2.X + tolerance)) && (point1.Y > (point2.Y - tolerance)) && (point1.Y < (point2.Y + tolerance)))
				return true;

			return false;
		}

		protected bool IsPointEqual(float x1, float y1, float x2, float y2, float tolerance)
		{
			if ((x1 > (x2 - tolerance)) && (x1 < (x2 + tolerance)) && (y1 > (y2 - tolerance)) && (y1 < (y2 + tolerance)))
				return true;

			return false;
		}

		protected bool IsValueEqual(float a, float b, float tolerance)
		{
			if ((a > (b - tolerance)) && (a < (b + tolerance)))
				return true;

			return false;
		}

		protected bool IsValueEqual(double a, double b, double tolerance)
		{
			if ((a > (b - tolerance)) && (a < (b + tolerance)))
				return true;

			return false;
		}
	}
}
