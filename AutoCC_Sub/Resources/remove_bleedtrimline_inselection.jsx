var bleed_width =  parseFloat ($.getenv("bleed_width"));
var bleed_height =  parseFloat ($.getenv("bleed_height"));
var trim_width =  parseFloat ($.getenv("trim_width"));
var trim_height =  parseFloat ($.getenv("trim_height"));
var tolerance =  parseFloat ($.getenv("tolerance"));

for (var i = 0; i < app.activeDocument.pageItems.length; i++)
			{
				var page_item = app.activeDocument.pageItems[i];
				var page_left = page_item.left * 0.352777778;
				var page_top = page_item.top * 0.352777778;
				var page_width = page_item.width * 0.352777778;
				var page_height = page_item.height * 0.352777778;
				if ((page_width > (bleed_width - tolerance)) && (page_width < (bleed_width + tolerance)) && (page_height > (bleed_height - tolerance)) && (page_height < (bleed_height + tolerance)))
				{
					if (page_item.typename == "PathItem")
					{
						{
							page_item.stroked = true;
							page_item.strokeColor = new NoColor();
						}
					}
				}
				else if ((page_width > (trim_width - tolerance)) && (page_width < (trim_width + tolerance)) && (page_height > (trim_height - tolerance)) && (page_height < (trim_height + tolerance)))
				{
					if (page_item.typename == "PathItem")
					{
						{
							page_item.stroked = true;
							page_item.strokeColor = new NoColor();
						}
					}
				}
			}