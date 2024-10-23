using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoCC_Sub
{
    public class Get_page
    {
        public int object_index { get; set; }
        public double item_left { get; set; }
        public double item_top { get; set; }
        public double item_width { get; set; }
        public double item_height { get; set; }

        public string item_type { get; set; }

        public string item_note { get; set; }

        public bool _in_item_check { get; set; }
    }
}
