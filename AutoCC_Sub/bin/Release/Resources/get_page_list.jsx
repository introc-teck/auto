

var bleed_width =  parseFloat ($.getenv("bleed_width"));
var bleed_height =  parseFloat ($.getenv("bleed_height"));
var tolerance =  parseFloat ($.getenv("tolerance"));

/*
var bleed_width = 66;
var bleed_height =66;
var tolerance = 0.5;
*/

var first_page_item = app.activeDocument.pageItems[0];
var x_gap = [];
				for (var i = 0; i < app.activeDocument.pageItems.length; i++)
				{
					var page_item = app.activeDocument.pageItems[i];
					if (page_item.typename == "GroupItem")
					{
						var page_left = page_item.left * 0.352777778;
						var page_top = page_item.top * 0.352777778;
						var page_width = page_item.width * 0.352777778;
						var page_height = page_item.height * 0.352777778;
                        
						if ((page_width > (bleed_width - tolerance)) && (page_width < (bleed_width + tolerance)) && (page_height > (bleed_height - tolerance)) && (page_height < (bleed_height + tolerance)))
						{
                            page_item.selected = true;
                                first_page_item.note = first_page_item.note + i + "," + page_left + "," + page_top + "," + page_width + "," + page_height +  ";";
						}
					}
				}
            
            $.setenv("status", first_page_item.note);