var bleed_width =  parseFloat ($.getenv("bleed_width"));
var bleed_height =  parseFloat ($.getenv("bleed_height"));
var tolerance =  parseFloat ($.getenv("tolerance"));

var first_page_item = app.activeDocument.selection[0];
			var selected_page_top = first_page_item.top * 0.352777778;
			for (var i = 0; i < app.activeDocument.pageItems.length; i++)
			{
				var page_item = app.activeDocument.pageItems[i];
				if (page_item.typename == "GroupItem" && page_item.selected == false)
				{
					var page_left = page_item.left * 0.352777778;
					var page_top = page_item.top * 0.352777778;
					var page_width = page_item.width * 0.352777778;
					var page_height = page_item.height * 0.352777778;
					if ((page_width > (bleed_width - tolerance)) && (page_width < (bleed_width + tolerance)) && (page_height > (bleed_height - tolerance)) && (page_height < (bleed_height + tolerance)))
					{
						if (page_top > (selected_page_top - 10.0) && page_top < (selected_page_top + 10))
						{
							first_page_item.note = first_page_item.note + i + "," + page_left + "," + page_top + "," + page_width + "," + page_height +  ";";
						}
					}
				}
			}