//수정 - allow -> Modify
//경고 - allow -> Warn
//무시 - allow -> Ignore
var PatternColor1 = "0";
var FillRGBColor1 = "0";
var StrokeRGBColor1 = "0";
var LabColor1 = "0";
var Texture1 = "0";
var Opacity1 = "0";


for (var i = 0; i < app.activeDocument.pageItems.length; i++)
{
    var page_item = app.activeDocument.pageItems[i];
    if (page_item.typename == "PathItem")
    {
        // 투명도
        if(page_item.opacity != 100) {
           Opacity1 = "1";
        }

        if (page_item.typename == "PatternColor")
            PatternColor1 = "1";

        if (page_item.fillColor.typename == "RGBColor")
            FillRGBColor1 = "1";

        if (page_item.strokeColor.typename == "RGBColor")
            StrokeRGBColor1 = "1";

        if (page_item.fillColor.typename == "LabColor")
            LabColor1 = "1";
        
        if (page_item.fillColor.typename == "Texture")
            Texture1 = "1";

    }
}

$.setenv("Opacity", Opacity1);
$.setenv("PatternColor", PatternColor1);
$.setenv("FillRGBColor", FillRGBColor1);
$.setenv("StrokeRGBColor", StrokeRGBColor1);
$.setenv("LabColor", LabColor1);
$.setenv("Texture", Texture1);