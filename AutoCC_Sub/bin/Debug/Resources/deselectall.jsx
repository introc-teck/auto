if (app.activeDocument.selection != null)
				{
					var selection_length = app.activeDocument.selection.length;
					for (var i = (selection_length - 1); i >= 0; i--)
					{
						app.activeDocument.selection[i].selected = false;
					}
				}