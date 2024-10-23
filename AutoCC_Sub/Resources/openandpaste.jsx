var newItem;
var docSelected = app.activeDocument.selection;
if (docSelected.length > 0) {
    // Create a new document and move the selected items to it.
    var docPreset = new DocumentPreset;
    docPreset.units = RulerUnits.Millimeters;
    docPreset.previewMode = DocumentPreviewMode.PixelPreview;
    docPreset.rasterResolution = DocumentRasterResolution.MediumResolution;
    docPreset.presetType = DocumentPresetType.Mobile;

    var newDoc = app.documents.addDocument(DocumentColorSpace.CMYK, docPreset);
    newDoc.defaultFillOverprint = false;
    newDoc.defaultStrokeOverprint = false;
    newDoc.activate();
    
    if (docSelected.length > 0) {
        for (var i = 0; i < docSelected.length; i++) {
            
            docSelected[i].selected = false;
            newItem = docSelected[i].duplicate(newDoc, ElementPlacement.PLACEATBEGINNING);
        }
    } else {
        docSelected.selected = false;
        newItem = docSelected.parent.duplicate(newDoc, ElementPlacement.PLACEATBEGINNING);
    }

app.executeMenuCommand('copy');
} else {
alert("Please select one or more art objects");
}
