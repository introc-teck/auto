
var page_left = parseFloat ($.getenv("page_left"));
var page_top = parseFloat ($.getenv("page_top"));
var page_width = parseFloat ($.getenv("page_width"));
var page_height = parseFloat ($.getenv("page_height"));

/*
var page_left = parseFloat ('2.03817');
var page_top = parseFloat ('1.00027');
var page_width = parseFloat ('87.999');
var page_height = parseFloat ('53.999');
*/
var selected = false;
			for (var i = 0; i < app.activeDocument.groupItems.length; i++)
			{
				var group_item = app.activeDocument.groupItems[i];
				var p_left = group_item.left * 0.352777778;
				var p_top = group_item.top * 0.352777778;
				var p_width = group_item.width * 0.352777778;
				var p_height = group_item.height * 0.352777778;

				if ((p_left > (page_left - 0.1)) && (p_left < (page_left + 0.1)) && (p_top > (page_top - 0.1)) && (p_top < (page_top + 0.1)) && (p_width > (page_width - 0.1)) && (p_width < (page_width + 0.1)) && (p_height > (page_height - 0.1)) && (p_height < (page_height + 0.1)))
				{
					group_item.selected = true;
					selected = true;
					break;
				}
			}
			if (selected == false)
			{
				for (var i = 0; i < app.activeDocument.rasterItems.length; i++)
				{
					var raster_item = app.activeDocument.rasterItems[i];
					var p_left = raster_item.left * 0.352777778;
					var p_top = raster_item.top * 0.352777778;
					var p_width = raster_item.width * 0.352777778;
					var p_height = raster_item.height * 0.352777778;
					if ((p_left > (page_left - 0.1)) && (p_left < (page_left + 0.1)) && (p_top > (page_top - 0.1)) && (p_top <  (page_top + 0.1)) && (p_width > (page_width - 0.1)) && (p_width < (page_width + 0.1)) && (p_height > (page_height - 0.1)) && (p_height < (page_height + 0.1)))
					{
						raster_item.selected = true;
						break;
					}
				}
			}