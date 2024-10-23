var activeDocument = app.activeDocument;
var save_path = "";
var jpg_file_path =  $.getenv("jpg_file_path");
        var jpegFolder = Folder(jpg_file_path);
        if (!jpegFolder.exists)
            jpegFolder.create();
        var codeStart = 97; // for a;
        for (var j = 0; j < activeDocument.artboards.length; j++) {
            var activeArtboard = activeDocument.artboards[j];
            activeDocument.artboards.setActiveArtboardIndex(j);
            var bounds = activeArtboard.artboardRect;
            var left = bounds[0];
            var top = bounds[1];
            var right = bounds[2];
            var bottom = bounds[3];
            var width = right - left;
            var height = top - bottom;
            if (app.activeDocument.rulerUnits == RulerUnits.Points) { //Add more if for more conversions
                width = width / 36;
                height = height / 36;
            }
            var fileName = activeDocument.name.split('.')[0] + '_' +pad(j+1, 3); + ".jpg";
            var destinationFile = File(jpegFolder + "/" + fileName);
            var type = ExportType.JPEG;
            var options = new ExportOptionsJPEG();
            options.antiAliasing = true;
            options.artBoardClipping = true;
            options.optimization = true;
            options.qualitySetting = 100; // Set Quality Setting
            activeDocument.exportFile(destinationFile, type, options);
            codeStart++;
            
            if(j != 0)
                save_path += ';';
            save_path += fileName;
        }
    
    $.setenv("thumbnails", fileName);
    
    
function pad(n, width, z) {
  z = z || '0';
  n = n + '';
  return n.length >= width ? n : new Array(width - n.length + 1).join(z) + n;
}