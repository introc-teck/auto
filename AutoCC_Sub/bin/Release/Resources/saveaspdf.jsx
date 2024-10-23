var pdf_file_path =  $.getenv("pdf_file_path");
//var pdf_file_path = 'C://AcceptManual/result_files/210713/GPT210713NC00011/A.pdf';
var pdf_save_options = new PDFSaveOptions();
			pdf_save_options.viewAfterSaving = false;
            var destFile = new File(pdf_file_path);
			app.activeDocument.saveAs(destFile, pdf_save_options);